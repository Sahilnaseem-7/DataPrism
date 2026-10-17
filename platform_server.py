"""DataPrism multi-user analytics web platform (single-node deployable MVP)."""
from datetime import datetime, date, timedelta
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO, StringIO
from pathlib import Path
from zipfile import ZipFile
import csv, hashlib, hmac, json, math, mimetypes, os, re, secrets, sqlite3, sys, urllib.parse

ROOT = Path(__file__).resolve().parent
DB = Path(os.environ.get("DATAPRISM_DB", ROOT / "data" / "platform.sqlite3"))
MAX_UPLOAD = 50 * 1024 * 1024
MAX_ROWS = 100_000
SESSION_DAYS = 30
PBKDF2_ROUNDS = 310_000
CATEGORIES = {
    "ecommerce": "E-commerce & sales",
    "retail": "Retail & inventory",
    "finance": "Finance & accounting",
    "marketing": "Marketing & campaigns",
    "product": "Product & SaaS",
    "operations": "Operations & workforce",
    "logistics": "Logistics & delivery",
    "education": "Academic institution",
    "customers": "Customer & growth analytics",
    "research": "Research & surveys",
    "general": "General / custom",
}

def connect():
    DB.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB, timeout=30)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    con.execute("PRAGMA journal_mode=WAL")
    con.executescript("""
      CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, email TEXT NOT NULL UNIQUE, salt TEXT NOT NULL, password_hash TEXT NOT NULL, category TEXT, created_at TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS sessions(token_hash TEXT PRIMARY KEY, user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE, expires_at TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS datasets(id INTEGER PRIMARY KEY, user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE, name TEXT NOT NULL, category TEXT NOT NULL, created_at TEXT NOT NULL, row_count INTEGER NOT NULL, columns_json TEXT NOT NULL, parent_id INTEGER, transform_note TEXT NOT NULL DEFAULT '');
      CREATE TABLE IF NOT EXISTS dataset_rows(dataset_id INTEGER NOT NULL REFERENCES datasets(id) ON DELETE CASCADE, row_num INTEGER NOT NULL, payload TEXT NOT NULL, PRIMARY KEY(dataset_id,row_num));
      CREATE TABLE IF NOT EXISTS saved_charts(id INTEGER PRIMARY KEY, user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE, dataset_id INTEGER NOT NULL REFERENCES datasets(id) ON DELETE CASCADE, title TEXT NOT NULL, config_json TEXT NOT NULL, created_at TEXT NOT NULL);
      CREATE INDEX IF NOT EXISTS ix_datasets_user ON datasets(user_id,id DESC);
      CREATE INDEX IF NOT EXISTS ix_saved_charts_dataset ON saved_charts(user_id,dataset_id,id DESC);
    """)
    dataset_fields={r[1] for r in con.execute("PRAGMA table_info(datasets)")}
    if "parent_id" not in dataset_fields: con.execute("ALTER TABLE datasets ADD COLUMN parent_id INTEGER")
    if "transform_note" not in dataset_fields: con.execute("ALTER TABLE datasets ADD COLUMN transform_note TEXT NOT NULL DEFAULT ''")
    con.commit()
    return con

def utcnow(): return datetime.utcnow().replace(microsecond=0)
def norm_name(value, used):
    base = re.sub(r"[^a-z0-9]+", "_", str(value or "").strip().lower()).strip("_") or "column"
    name, i = base, 2
    while name in used:
        name, i = f"{base}_{i}", i + 1
    used.add(name)
    return name
def to_number(value):
    if isinstance(value, bool) or value is None: return None
    if isinstance(value, (int, float)): return float(value)
    s = str(value).strip()
    if not s: return None
    s = s.replace(",", "").replace("$", "").replace("£", "").replace("€", "")
    if s.startswith("(") and s.endswith(")"): s = "-" + s[1:-1]
    try:
        n=float(s)
        return n if math.isfinite(n) else None
    except (ValueError, OverflowError): return None
def to_date(value):
    if isinstance(value, (datetime, date)): return value.strftime("%Y-%m-%d")
    if value is None: return None
    s = str(value).strip()
    if not s: return None
    try: return datetime.fromisoformat(s.replace("Z", "+00:00")).strftime("%Y-%m-%d")
    except ValueError: pass
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%m/%d/%Y %H:%M", "%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try: return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError: pass
    return None
def infer_kind(name, values):
    vals = [v for v in values if v is not None and str(v).strip()]
    if not vals: return "text"
    key = name.lower()
    if re.search(r"(^|_)(id|code|zip|postal|phone|index|number)$", key) or key.endswith("_id"): return "category"
    nums = sum(to_number(v) is not None for v in vals)
    dates = sum(to_date(v) is not None for v in vals)
    if dates / len(vals) >= .75 and not all(to_number(v) is not None for v in vals): return "date"
    if nums / len(vals) >= .75: return "number"
    if dates / len(vals) >= .75: return "date"
    return "category" if len(set(str(v) for v in vals)) <= min(100, max(20, len(vals)//4)) else "text"
def decode_upload(filename, blob):
    ext = Path(filename).suffix.lower()
    if ext in (".csv", ".tsv"):
        decoded = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try: decoded = blob.decode(encoding); break
            except UnicodeDecodeError: pass
        delimiter="\t" if ext==".tsv" else ","
        if ext==".csv":
            try: delimiter=csv.Sniffer().sniff((decoded or "")[:8192],delimiters=",;\t|").delimiter
            except csv.Error: pass
        rows = list(csv.reader(StringIO(decoded or ""),delimiter=delimiter))
    elif ext in (".json", ".jsonl", ".ndjson"):
        decoded=blob.decode("utf-8-sig")
        try:
            parsed=[json.loads(line) for line in decoded.splitlines() if line.strip()] if ext in (".jsonl",".ndjson") else json.loads(decoded)
        except json.JSONDecodeError as e: raise ValueError(f"Invalid JSON near line {e.lineno}, column {e.colno}.") from e
        if isinstance(parsed,dict): parsed=parsed.get("records",parsed.get("data",parsed))
        if not isinstance(parsed,list) or not parsed or not all(isinstance(x,dict) for x in parsed): raise ValueError("JSON uploads must contain an array of records, or an object with a records/data array.")
        headers=[]
        for obj in parsed:
            for key in obj:
                if key not in headers: headers.append(key)
        rows=[headers]+[[json.dumps(obj.get(h),ensure_ascii=False) if isinstance(obj.get(h),(dict,list)) else obj.get(h) for h in headers] for obj in parsed]
    elif ext in (".xlsx", ".xlsm"):
        try: from openpyxl import load_workbook
        except ImportError as e: raise ValueError("Excel uploads need openpyxl. Install requirements and restart.") from e
        book = load_workbook(BytesIO(blob), read_only=True, data_only=True)
        rows = [list(r) for r in book.active.iter_rows(values_only=True)]
        book.close()
    else: raise ValueError("Choose a CSV, TSV, JSON, JSONL or Excel (.xlsx) file.")
    if len(rows) < 2: raise ValueError("Add a header row and at least one row of data.")
    if len(rows) - 1 > MAX_ROWS: raise ValueError(f"This MVP supports up to {MAX_ROWS:,} rows per upload.")
    headers, used = [], set()
    for raw in rows[0]: headers.append(norm_name(raw, used))
    if not any(str(x or "").strip() for x in rows[0]): raise ValueError("The first row must contain column names.")
    cols = []
    for i, name in enumerate(headers):
        vals = [row[i] if i < len(row) else None for row in rows[1:]]
        kind = infer_kind(name, vals)
        cols.append({"name": name, "label": str(rows[0][i] or name).strip(), "kind": kind, "invalid_count": 0})
    records = []
    for row in rows[1:]:
        if not any(v is not None and str(v).strip() for v in row): continue
        item = {}
        for i, col in enumerate(cols):
            raw = row[i] if i < len(row) else None
            if raw is None or str(raw).strip() == "": item[col["name"]] = None
            elif col["kind"] == "number":
                parsed=to_number(raw)
                item[col["name"]] = parsed if parsed is not None else str(raw).strip()
                if parsed is None: col["invalid_count"] += 1
            elif col["kind"] == "date":
                parsed=to_date(raw)
                item[col["name"]] = parsed if parsed is not None else str(raw).strip()
                if parsed is None: col["invalid_count"] += 1
            else: item[col["name"]] = str(raw).strip()
        records.append(item)
    if not records: raise ValueError("No non-empty data rows were found.")
    return cols, records
def serialize_dataset(row):
    d = dict(row); d["columns"] = json.loads(d.pop("columns_json")); return d
def dataset_row(con, user_id, did):
    return con.execute("SELECT * FROM datasets WHERE id=? AND user_id=?", (did,user_id)).fetchone()
def read_records(con, did):
    return [json.loads(r[0]) for r in con.execute("SELECT payload FROM dataset_rows WHERE dataset_id=? ORDER BY row_num", (did,))]
def save_dataset(con, user_id, name, category, cols, records, parent_id=None, transform_note=""):
    cur = con.execute("INSERT INTO datasets(user_id,name,category,created_at,row_count,columns_json,parent_id,transform_note) VALUES(?,?,?,?,?,?,?,?)", (user_id,Path(name).name[:180],category,utcnow().isoformat(),len(records),json.dumps(cols,ensure_ascii=False),parent_id,transform_note))
    did = cur.lastrowid
    con.executemany("INSERT INTO dataset_rows(dataset_id,row_num,payload) VALUES(?,?,?)", [(did,i,json.dumps(r,ensure_ascii=False)) for i,r in enumerate(records)])
    con.commit()
    return did
def apply_filter(rows, cols, params):
    field=params.get("filter_field",[""])[0]
    if not field: return rows
    if field not in {c["name"] for c in cols}: return rows
    op=params.get("filter_op",["eq"])[0]; raw=params.get("filter_value",[""])[0]
    if op=="is_empty": return [r for r in rows if r.get(field) is None or str(r.get(field)).strip()==""]
    if op=="not_empty": return [r for r in rows if r.get(field) is not None and str(r.get(field)).strip()!=""]
    num=to_number(raw)
    def match(row):
        value=row.get(field)
        if value is None: return False
        if op=="contains": return raw.casefold() in str(value).casefold()
        if op=="neq": return str(value).casefold()!=raw.casefold()
        if op in ("gt","gte","lt","lte"):
            left=to_number(value)
            if left is None or num is None: return False
            return {"gt":left>num,"gte":left>=num,"lt":left<num,"lte":left<=num}[op]
        return str(value).casefold()==raw.casefold()
    return [r for r in rows if match(r)]
def profile_data(rows, cols):
    seen=set(); duplicate_rows=0
    for row in rows:
        key=json.dumps(row,sort_keys=True,ensure_ascii=False)
        if key in seen: duplicate_rows+=1
        seen.add(key)
    profiles=[]
    for col in cols:
        name=col["name"]; vals=[r.get(name) for r in rows]; present=[v for v in vals if v is not None and str(v).strip()!=""]
        counts={}
        for v in present: counts[str(v)]=counts.get(str(v),0)+1
        invalid=sum(1 for v in present if (col["kind"]=="number" and to_number(v) is None) or (col["kind"]=="date" and to_date(v) is None))
        item={"name":name,"label":col["label"],"kind":col["kind"],"missing":len(vals)-len(present),"invalid":invalid,"distinct":len(counts),"unique":sum(1 for c in counts.values() if c==1),"top_values":sorted(counts.items(),key=lambda x:(-x[1],x[0]))[:5]}
        if col["kind"]=="number":
            xs=sorted(float(v) for v in present if isinstance(v,(int,float)) and math.isfinite(float(v)))
            if xs:
                def quantile(p):
                    position=(len(xs)-1)*p; lower=int(math.floor(position)); upper=min(len(xs)-1,lower+1)
                    return xs[lower]+(xs[upper]-xs[lower])*(position-lower)
                n=len(xs); mean=sum(xs)/n; q1,q3=quantile(.25),quantile(.75); iqr=q3-q1
                pop_var=sum((x-mean)**2 for x in xs)/n
                sample_var=sum((x-mean)**2 for x in xs)/(n-1) if n>1 else 0
                item.update({"count":n,"min":xs[0],"max":xs[-1],"range":xs[-1]-xs[0],"mean":mean,"median":quantile(.5),"q1":q1,"q3":q3,"iqr":iqr,"variance_population":pop_var,"variance_sample":sample_var,"stddev_population":math.sqrt(pop_var),"stddev_sample":math.sqrt(sample_var),"outliers":sum(1 for x in xs if x<q1-1.5*iqr or x>q3+1.5*iqr)})
        profiles.append(item)
    complete=sum(max(0,len(rows)-p["missing"]-p["invalid"]) for p in profiles)
    denom=max(1,len(rows)*max(1,len(cols)))
    label_counts={}
    for col in cols: label_counts[str(col["label"]).strip().casefold()]=label_counts.get(str(col["label"]).strip().casefold(),0)+1
    duplicate_columns=sum(count-1 for count in label_counts.values() if count>1)
    missing_cells=sum(p["missing"] for p in profiles)
    rows_with_gaps=sum(1 for row in rows if any(row.get(c["name"]) is None or str(row.get(c["name"])).strip()=="" or (c["kind"]=="number" and to_number(row.get(c["name"])) is None) or (c["kind"]=="date" and to_date(row.get(c["name"])) is None) for c in cols))
    columns_with_missing=sum(1 for p in profiles if p["missing"]+p["invalid"]>0)
    return {"rows":len(rows),"columns":len(cols),"rows_with_gaps":rows_with_gaps,"columns_with_missing":columns_with_missing,"exact_duplicate_rows":duplicate_rows,"duplicate_columns":duplicate_columns,"missing_cells":missing_cells,"completeness_pct":round(complete/denom*100,1),"column_profiles":profiles}

def rfm_analysis(rows, cols):
    customer=next((c for c in cols if re.search(r"(^|_)(customer|client|account|user)(_id)$",c["name"])),None)
    datecol=next((c for c in cols if c["kind"]=="date" and re.search(r"order|purchase|transaction|invoice|date",c["name"])),None)
    money=next((c for c in cols if c["kind"]=="number" and re.search(r"sales|revenue|amount|spend|monetary|total|value",c["name"])),None)
    order=next((c for c in cols if re.search(r"(^|_)(order|transaction|invoice)(_id)$",c["name"])),None)
    if not customer or not datecol or not money: return None
    activity={}
    for i,row in enumerate(rows):
        cid=row.get(customer["name"]); date_value=to_date(row.get(datecol["name"])); amount=row.get(money["name"])
        if cid in (None,"") or not date_value: continue
        item=activity.setdefault(str(cid),{"dates":[],"orders":set(),"frequency":0,"monetary":0.0})
        item["dates"].append(date_value)
        item["orders"].add(str(row.get(order["name"])) if order and row.get(order["name"]) not in (None,"") else f"row-{i}")
        if isinstance(amount,(int,float)): item["monetary"]+=float(amount)
    if not activity: return None
    reference=max(max(item["dates"]) for item in activity.values())
    ref_date=datetime.strptime(reference,"%Y-%m-%d").date()
    records=[]
    for cid,item in activity.items():
        recency=(ref_date-datetime.strptime(max(item["dates"]),"%Y-%m-%d").date()).days
        records.append({"customer":cid,"recency":recency,"frequency":len(item["orders"]),"monetary":item["monetary"]})
    def quintile(values,value):
        lower=sum(1 for x in values if x<value); equal=sum(1 for x in values if x==value)
        return max(1,min(5,int(((lower+equal/2)/len(values))*5)+1))
    for row in records:
        r=6-quintile([x["recency"] for x in records],row["recency"])
        f=quintile([x["frequency"] for x in records],row["frequency"])
        m=quintile([x["monetary"] for x in records],row["monetary"])
        row["r_score"],row["f_score"],row["m_score"]=r,f,m
        row["segment"]="Champions" if r>=4 and f>=4 and m>=4 else "Loyal" if f>=4 and r>=3 else "New / promising" if r>=4 and f<=2 else "At risk" if r<=2 and f>=3 else "Hibernating" if r<=2 and f<=2 else "Needs attention"
    segments={}
    for row in records:
        item=segments.setdefault(row["segment"],{"customers":0,"frequency":0,"monetary":0.0,"recency":0})
        item["customers"]+=1; item["frequency"]+=row["frequency"]; item["monetary"]+=row["monetary"]; item["recency"]+=row["recency"]
    result=[]
    for label,item in sorted(segments.items(),key=lambda x:(-x[1]["customers"],x[0])):
        n=item["customers"]
        result.append({"segment":label,"customers":n,"share_pct":round(n/len(records)*100,1),"avg_orders":round(item["frequency"]/n,2),"avg_monetary":round(item["monetary"]/n,2),"avg_recency_days":round(item["recency"]/n,1)})
    return {"customer_count":len(records),"reference_date":reference,"customer_label":customer["label"],"date_label":datecol["label"],"monetary_label":money["label"],"frequency_label":order["label"] if order else "Distinct data rows per customer","segments":result,"method":"Recency uses the latest date in the filtered data. Frequency counts distinct order IDs when present; otherwise it counts customer rows. Monetary value sums the selected amount field. Quintile scores use this dataset and are relative, not industry benchmarks."}

def correlation_data(rows, cols):
    fields=[c for c in cols if c["kind"]=="number"][:20]
    values={c["name"]:[float(row[c["name"]]) if isinstance(row.get(c["name"]),(int,float)) and math.isfinite(float(row[c["name"]])) else None for row in rows] for c in fields}
    matrix=[[None for _ in fields] for _ in fields]
    for i,a in enumerate(fields):
        for j in range(i,len(fields)):
            b=fields[j]; n=sx=sy=sxx=syy=sxy=0.0
            for x,y in zip(values[a["name"]],values[b["name"]]):
                if x is None or y is None: continue
                n+=1; sx+=x; sy+=y; sxx+=x*x; syy+=y*y; sxy+=x*y
            dx=n*sxx-sx*sx; dy=n*syy-sy*sy
            corr=(n*sxy-sx*sy)/math.sqrt(dx*dy) if n>=2 and dx>0 and dy>0 else None
            matrix[i][j]=matrix[j][i]=round(max(-1.0,min(1.0,corr)),4) if corr is not None else None
    return {"fields":[{"name":c["name"],"label":c["label"]} for c in fields],"matrix":matrix,"truncated":sum(1 for c in cols if c["kind"]=="number")>20}

def summary(rows, cols, category):
    nums=[c for c in cols if c["kind"]=="number"]
    dims=[c for c in cols if c["kind"] in ("category","text")]
    datecol=next((c for c in cols if c["kind"]=="date"),None)
    def field(pattern, pool=nums): return next((c for c in pool if re.search(pattern,c["name"])),None)
    def values(c): return [r.get(c["name"]) for r in rows if c and isinstance(r.get(c["name"]),(int,float))]
    def total(c): return sum(values(c))
    def mean(c):
        xs=values(c); return sum(xs)/len(xs) if xs else None
    def distinct(pattern):
        c=field(pattern,cols)
        vals={str(r.get(c["name"])) for r in rows if c and r.get(c["name"]) not in (None,"")}
        return len(vals),c
    primary=field(r"sales|revenue|amount|total|price|score|grade|mark|attendance|spend|actual") or (nums[0] if nums else None)
    cards=[{"label":"Records","value":len(rows),"detail":f"Across {len(cols)} columns"}]
    if primary:
        vals=values(primary); cards.append({"label":f"Total {primary['label']}","value":sum(vals),"detail":f"Average {sum(vals)/len(vals):,.2f}" if vals else "No numeric values"})
    if dims:
        c=next((x for x in dims if re.search(r"customer|student|learner|account|user|member|response|order|invoice",x["name"])),dims[0]); unique=len({str(r.get(c["name"])) for r in rows if r.get(c["name"]) is not None})
        cards.append({"label":"Students" if category=="education" else "Customers" if category=="customers" else "Unique IDs" if re.search(r"_id$",c["name"]) else "Groups","value":unique,"detail":f"Unique values in {c['label']}"})
    if datecol:
        ds=sorted(str(r.get(datecol["name"])) for r in rows if r.get(datecol["name"]))
        cards.append({"label":"Date range","value":f"{ds[0]} – {ds[-1]}" if ds else "—","detail":datecol["label"]})
    else: cards.append({"label":"Numeric fields","value":len(nums),"detail":"Available for analysis"})
    def put(i,label,value,detail):
        while len(cards)<=i: cards.append({"label":"—","value":"—","detail":""})
        cards[i]={"label":label,"value":value if value is not None else "—","detail":detail}
    def total_card(i,label,c,detail=None): put(i,label,total(c),detail or (c["label"] if c else "No matching numeric column"))
    if category=="ecommerce":
        order_n,order_col=distinct(r"(^|_)(order|invoice|transaction)(_id)?$")
        sales=field(r"sales|revenue|gross_amount|order_value"); profit=field(r"profit|margin")
        put(0,"Orders",order_n if order_col else len(rows),f"Unique {order_col['label']}" if order_col else "Rows in this dataset")
        total_card(1,"Revenue",sales)
        if profit: total_card(2,"Profit",profit)
        else: put(2,"Products",distinct(r"product(_id|_name)?$")[0],"Distinct products")
        margin=total(profit)/total(sales)*100 if profit and sales and total(sales) else None
        put(3,"Profit margin",f"{margin:.1f}%" if margin is not None else "—","Total profit ÷ total revenue" if margin is not None else "Needs revenue and profit columns")
    elif category=="retail":
        n,t=distinct(r"(^|_)(transaction|receipt|order)(_id)?$"); sales=field(r"sales|revenue|amount|total"); qty=field(r"quantity|units_sold|units")
        put(0,"Transactions",n if t else len(rows),f"Unique {t['label']}" if t else "Rows in this dataset"); total_card(1,"Sales",sales); total_card(2,"Units sold",qty)
        pcount,_=distinct(r"product(_id|_name)?$"); put(3,"Products",pcount,"Distinct products")
        stock=field(r"stock_on_hand|on_hand|inventory|(^|_)stock$"); reorder=field(r"reorder_point|reorder_level|min_stock|minimum_stock")
        if stock and reorder:
            sku=field(r"product(_id|_name)?$|sku",cols); levels={}
            for row in rows:
                key=str(row.get(sku["name"])) if sku and row.get(sku["name"]) is not None else str(id(row))
                if isinstance(row.get(stock["name"]),(int,float)) and isinstance(row.get(reorder["name"]),(int,float)): levels[key]=(row[stock["name"]],row[reorder["name"]])
            put(3,"SKUs to reorder",sum(1 for a,b in levels.values() if a<=b),f"{stock['label']} at or below {reorder['label']}")
    elif category=="finance":
        actual=field(r"actual|spent|expense"); budget=field(r"budget|planned|target"); a,b=total(actual),total(budget)
        if actual and budget:
            variance=a-b; put(0,"Budgeted amount",b,budget["label"]); put(1,"Actual amount",a,actual["label"]); put(2,"Budget variance",variance,"Actual less budget; positive means over plan"); put(3,"Variance vs budget",f"{variance/b*100:.1f}%" if b else "—","Actual less budget ÷ budget")
    elif category=="marketing":
        spend=field(r"spend|cost|ad_cost"); revenue=field(r"revenue|sales|attributed_value"); conv=field(r"conversion|conversions|leads")
        total_card(1,"Marketing spend",spend); total_card(2,"Attributed revenue",revenue)
        put(3,"Revenue / spend",f"{total(revenue)/total(spend):.2f}×" if spend and revenue and total(spend) else "—",f"Blended return · cost per conversion {total(spend)/total(conv):,.2f}" if spend and conv and total(conv) else "Blended return; attribution follows the source data")
        if conv: put(0,conv["label"],total(conv),"Total recorded conversions / leads")
    elif category=="product":
        users=field(r"active_users|users|user_count"); events=field(r"events|sessions|usage|event_count"); feature_n,_=distinct(r"feature(_name)?$")
        put(0,"Features",feature_n,"Distinct features observed"); put(1,f"Average {users['label']}" if users else "Active users",mean(users) if users else None,"Mean across reporting rows" if users else "Add an active-user field"); total_card(2,"Usage events",events,"Recorded activity")
    elif category=="operations":
        units=field(r"completed|throughput|units|tasks|output"); defects=field(r"defect|error|rework")
        total_card(1,"Completed",units,"Operational throughput"); total_card(2,"Defects / errors",defects)
        put(3,"Defect rate",f"{total(defects)/total(units)*100:.2f}%" if units and defects and total(units) else "—","Defects ÷ completed units")
    elif category=="logistics":
        delivery=field(r"shipping_days|delivery_days|transit_days|lead_time"); ontime=field(r"on_time|ontime")
        put(0,"Shipments",len(rows),"Rows in this dataset"); put(2,f"Average {delivery['label']}" if delivery else "Delivery time",mean(delivery) if delivery else None,delivery["label"] if delivery else "Add delivery or transit days")
        ov=values(ontime); rate=sum(ov)/len(ov) if ov else None; pct=(rate*100 if max(ov)<=1 else rate) if ov else None; put(3,"On-time rate",f"{pct:.1f}%" if pct is not None else "—","Average of the on-time indicator")
    elif category=="education":
        n,student=distinct(r"student(_id|_name)?$"); score=field(r"score|mark|grade|gpa"); attendance=field(r"attendance")
        put(0,"Students",n if student else len(rows),f"Unique {student['label']}" if student else "Rows in this dataset"); put(1,f"Average {score['label']}" if score else "Average score",mean(score) if score else None,"Mean across recorded outcomes" if score else "Add a score / grade field"); put(2,f"Average {attendance['label']}" if attendance else "Attendance",mean(attendance) if attendance else None,"Mean across reporting rows" if attendance else "Add an attendance field")
    elif category=="customers":
        n,cust=distinct(r"customer(_id|_name)?$|(^|_)client(_id)?$"); spend=field(r"lifetime_value|ltv|annual_spend|order_value|spend|revenue|value"); order_n,order_col=distinct(r"(^|_)(order|transaction|invoice)(_id)?$"); segments,segment_col=distinct(r"segment|cohort|group")
        customer_n=n if cust else len(rows); put(0,"Customers",customer_n,f"Unique {cust['label']}" if cust else "Rows in this dataset"); put(1,f"Average {spend['label']}" if spend else "Average customer value",mean(spend) if spend else None,"Mean per row; RFM monetary totals are customer-level" if spend else "Add a customer value field"); put(2,"Orders per customer",order_n/customer_n if order_col and customer_n else len(rows)/customer_n if customer_n else None,"Distinct orders ÷ unique customers" if order_col else "Rows ÷ unique customers"); put(3,"Customer segments",segments if segment_col else "—",f"Distinct values in {segment_col['label']}" if segment_col else "No segment field detected")
    elif category=="research":
        n,respondent=distinct(r"respondent(_id|_name)?$|participant(_id)?$"); score=field(r"response_score|score|rating|satisfaction"); items,itemcol=distinct(r"question|item|prompt")
        put(0,"Responses",len(rows),"Rows in this dataset"); put(1,"Unique respondents",n if respondent else "—",respondent["label"] if respondent else "No respondent ID detected"); put(2,f"Mean {score['label']}" if score else "Mean response",mean(score) if score else None,"Average observed score" if score else "Add a numeric response field"); put(3,"Questions / items",items if itemcol else "—",f"Distinct values in {itemcol['label']}" if itemcol else "No question/item field detected")
    return cards[:4],primary,dims,datecol

def category_analysis(rows, cols, category):
    """Return one domain-specific analysis only when the source fields support it."""
    nums=[c for c in cols if c["kind"]=="number"]
    dims=[c for c in cols if c["kind"] in ("category","text")]
    def field(pattern,pool): return next((c for c in pool if re.search(pattern,c["name"],re.I)),None)
    def aggregate(group, metrics):
        values={}
        for row in rows:
            key=row.get(group["name"])
            if key in (None,""): continue
            item=values.setdefault(str(key),{"n":0,"sums":{m["name"]:0.0 for m in metrics},"counts":{m["name"]:0 for m in metrics}})
            item["n"]+=1
            for metric in metrics:
                value=row.get(metric["name"])
                if isinstance(value,(int,float)):
                    item["sums"][metric["name"]]+=float(value); item["counts"][metric["name"]]+=1
        return values
    if category=="ecommerce":
        discount=field(r"discount|markdown",nums); sales=field(r"sales|revenue|gross_amount|order_value",nums); profit=field(r"profit",nums)
        if discount and profit:
            buckets={"No discount":{"rows":0,"sales":0.0,"profit":0.0},"1–10%":{"rows":0,"sales":0.0,"profit":0.0},"11–20%":{"rows":0,"sales":0.0,"profit":0.0},"21–30%":{"rows":0,"sales":0.0,"profit":0.0},"Over 30%":{"rows":0,"sales":0.0,"profit":0.0}}
            for row in rows:
                d=row.get(discount["name"]); p=row.get(profit["name"]); s=row.get(sales["name"]) if sales else None
                if not isinstance(d,(int,float)) or not isinstance(p,(int,float)): continue
                rate=float(d)/100 if d>1 else float(d)
                label="No discount" if rate<=0 else "1–10%" if rate<=.10 else "11–20%" if rate<=.20 else "21–30%" if rate<=.30 else "Over 30%"
                item=buckets[label]; item["rows"]+=1; item["profit"]+=float(p)
                if isinstance(s,(int,float)): item["sales"]+=float(s)
            result=[]
            for label,item in buckets.items():
                if not item["rows"]: continue
                margin=item["profit"]/item["sales"]*100 if item["sales"] else None
                result.append([label,item["rows"],round(item["sales"],2) if sales else None,round(item["profit"],2),round(margin,2) if margin is not None else None])
            return {"title":"Discount and margin analysis","description":"Compare profit and margin across discount bands. Discount values above 1 are treated as percentages.","headers":["Discount band","Rows","Revenue","Profit","Profit margin %"],"rows":result,"footnote":"Aggregated from the active filtered rows; discounts are grouped by observed row value."}
    elif category=="retail":
        product=field(r"product(_name)?$|item(_name)?$|sku$",dims); sales=field(r"sales|revenue|amount|total",nums)
        if product and sales:
            grouped=aggregate(product,[sales]); ranked=sorted(((k,v["sums"][sales["name"]]) for k,v in grouped.items()),key=lambda x:-x[1]); total=sum(v for _,v in ranked); cumulative=0; result=[]
            for name,value in ranked[:100]:
                cumulative+=value; share=value/total*100 if total else 0; cumshare=cumulative/total*100 if total else 0
                tier="A" if cumshare<=80 or (not result and share>=80) else "B" if cumshare<=95 else "C"
                result.append([name,round(value,2),round(share,2),round(cumshare,2),tier])
            return {"title":"Product contribution (ABC) analysis","description":"Rank products by sales and classify cumulative contribution: A through 80%, B through 95%, then C.","headers":["Product","Sales","Sales share %","Cumulative %","ABC class"],"rows":result,"footnote":"Top 100 products shown. ABC thresholds are a prioritization convention; apply business-specific thresholds before operational use."}
    elif category=="finance":
        actual=field(r"actual|spent|expense",nums); budget=field(r"budget|planned|target",nums); group=field(r"department|cost_center|account|category|team",dims)
        if actual and budget and group:
            grouped=aggregate(group,[actual,budget]); result=[]
            for name,item in sorted(grouped.items(),key=lambda x:-(x[1]["sums"][actual["name"]]-x[1]["sums"][budget["name"]])):
                a=item["sums"][actual["name"]]; b=item["sums"][budget["name"]]; variance=a-b
                result.append([name,round(b,2),round(a,2),round(variance,2),round(variance/b*100,2) if b else None,"Over budget" if variance>0 else "Under / on budget"])
            return {"title":f"Budget variance by {group['label'].lower()}","description":"Separate overspend from underspend and rank the largest overspends first.","headers":[group["label"],"Budget","Actual","Variance","Variance %","Status"],"rows":result[:100],"footnote":"Variance is actual minus budget; positive variance indicates overspend."}
    elif category=="marketing":
        group=field(r"channel|source|medium|campaign|ad_group",dims); spend=field(r"spend|cost|ad_cost",nums); revenue=field(r"revenue|sales|attributed_value",nums); conv=field(r"conversion|conversions|leads|signups",nums); impressions=field(r"impression",nums)
        if group and (spend or revenue or conv):
            metrics=[m for m in (spend,revenue,conv,impressions) if m]; grouped=aggregate(group,metrics); result=[]
            for name,item in sorted(grouped.items(),key=lambda x:-(x[1]["sums"].get(revenue["name"],0) if revenue else x[1]["sums"].get(conv["name"],0))):
                s=item["sums"].get(spend["name"],0) if spend else None; r=item["sums"].get(revenue["name"],0) if revenue else None; c=item["sums"].get(conv["name"],0) if conv else None; imp=item["sums"].get(impressions["name"],0) if impressions else None
                result.append([name,round(s,2) if s is not None else None,round(r,2) if r is not None else None,round(r/s,2) if r is not None and s else None,round(c,2) if c is not None else None,round(s/c,2) if s is not None and c else None,round(c/imp*100,2) if c is not None and imp else None])
            return {"title":f"Efficiency by {group['label'].lower()}","description":"Compare investment, attributed revenue, return on ad spend (ROAS), cost per conversion and conversion rate when the fields exist.","headers":[group["label"],"Spend","Revenue","ROAS ×","Conversions","Cost / conversion","Conversion rate %"],"rows":result[:100],"footnote":"ROAS and conversion metrics are calculated from the uploaded attribution fields; they are not incrementality estimates."}
    elif category=="product":
        group=field(r"feature(_name)?$|plan|tier|cohort",dims); users=field(r"active_users|user_count|users",nums); events=field(r"events|sessions|usage|event_count",nums); tickets=field(r"tickets|support_requests",nums)
        metrics=[m for m in (users,events,tickets) if m]
        if group and metrics:
            grouped=aggregate(group,metrics); result=[]
            for name,item in sorted(grouped.items(),key=lambda x:-(x[1]["sums"].get(events["name"],0) if events else x[1]["sums"].get(users["name"],0))):
                user_avg=item["sums"][users["name"]]/item["counts"][users["name"]] if users and item["counts"][users["name"]] else None
                event_total=item["sums"][events["name"]] if events else None
                ticket_total=item["sums"][tickets["name"]] if tickets else None
                ratio=event_total/user_avg if event_total is not None and user_avg else None
                result.append([name,item["n"],round(user_avg,2) if user_avg is not None else None,round(event_total,2) if event_total is not None else None,round(ratio,2) if ratio is not None else None,round(ticket_total,2) if ticket_total is not None else None])
            return {"title":f"Product adoption by {group['label'].lower()}","description":"Compare active users, usage volume and support demand across product features or plans.","headers":[group["label"],"Reporting rows","Mean active users","Usage events","Events / active user","Support tickets"],"rows":result[:100],"footnote":"Events per active user divides grouped usage events by mean active users; compare only when both metrics share the same reporting grain."}
    elif category=="customers":
        group=field(r"segment|cohort|source|channel|region",dims); customer=field(r"(^|_)(customer|client|account|user)(_id|_name)?$",cols); value=field(r"lifetime_value|ltv|order_value|spend|revenue|value",nums); order=field(r"(^|_)(order|transaction|invoice)(_id)?$",cols)
        if group and customer and value:
            levels={}
            for row in rows:
                label=row.get(group["name"]); cid=row.get(customer["name"]); amount=row.get(value["name"])
                if label in (None,"") or cid in (None,""): continue
                item=levels.setdefault(str(label),{"customers":set(),"orders":set(),"values":[]})
                item["customers"].add(str(cid))
                if order and row.get(order["name"]) not in (None,""): item["orders"].add(str(row[order["name"]]))
                if isinstance(amount,(int,float)): item["values"].append(float(amount))
            result=[]
            for name,item in sorted(levels.items(),key=lambda x:-sum(x[1]["values"])):
                vals=item["values"]
                result.append([name,len(item["customers"]),len(item["orders"]) if order else None,round(sum(vals),2) if vals else None,round(sum(vals)/len(vals),2) if vals else None])
            return {"title":f"Customer value by {group['label'].lower()}","description":"Compare customer counts, orders and observed transaction or customer value across segments.","headers":[group["label"],"Unique customers","Orders","Total value","Mean value per row"],"rows":result[:100],"footnote":"The value field is summarized at its source row grain; use the RFM panel for customer-level monetary totals when transaction data is available."}
    elif category in ("education","research"):
        group=field(r"course|subject|class|cohort|group|condition|treatment|question|item",dims); score=field(r"score|mark|grade|gpa|rating|response",nums); attendance=field(r"attendance",nums)
        metric=score or attendance
        if group and metric:
            grouped=aggregate(group,[metric]+([attendance] if attendance and attendance!=metric else [])); result=[]
            for name,item in sorted(grouped.items(),key=lambda x:-x[1]["counts"][metric["name"]]):
                n=item["counts"][metric["name"]]; avg=item["sums"][metric["name"]]/n if n else None
                row=[name,n,round(avg,3) if avg is not None else None]
                if attendance and attendance!=metric:
                    an=item["counts"][attendance["name"]]; row.append(round(item["sums"][attendance["name"]]/an,3) if an else None)
                result.append(row)
            headers=[group["label"],"Valid responses",f"Mean {metric['label']}"]+([f"Mean {attendance['label']}"] if attendance and attendance!=metric else [])
            title="Learning outcomes by " if category=="education" else "Response summary by "
            return {"title":title+group["label"].lower(),"description":"Compare group averages with the number of valid observations beside each estimate.","headers":headers,"rows":result[:100],"footnote":"Group averages are descriptive. Small group sizes should be interpreted cautiously."}
    elif category=="logistics":
        group=field(r"carrier|ship_mode|service|route|region|warehouse",dims); days=field(r"shipping_days|delivery_days|transit_days|lead_time",nums); ontime=field(r"on_time|ontime",nums)
        if group and (days or ontime):
            metrics=[m for m in (days,ontime) if m]; grouped=aggregate(group,metrics); result=[]
            for name,item in sorted(grouped.items(),key=lambda x:x[1]["sums"][days["name"]]/x[1]["counts"][days["name"]] if days and x[1]["counts"][days["name"]] else 0,reverse=True):
                av=item["sums"][days["name"]]/item["counts"][days["name"]] if days and item["counts"][days["name"]] else None
                raw_rate=item["sums"][ontime["name"]]/item["counts"][ontime["name"]] if ontime and item["counts"][ontime["name"]] else None
                rate=(raw_rate*100 if raw_rate<=1 else raw_rate) if raw_rate is not None else None
                result.append([name,item["n"],round(av,2) if av is not None else None,round(rate,2) if rate is not None else None])
            return {"title":f"Delivery service by {group['label'].lower()}","description":"Surface slower routes or carriers and compare on-time percentages where an indicator is available.","headers":[group["label"],"Shipments",f"Mean {days['label']}" if days else "Mean delivery days","On-time %"],"rows":result[:100],"footnote":"On-time fields should be encoded as 0/1 or percentages from 0–100."}
    elif category=="operations":
        group=field(r"team|department|shift|line|site",dims); output=field(r"completed|throughput|units|tasks|output",nums); defects=field(r"defect|error|rework",nums); cycle=field(r"cycle_time|processing_time|duration",nums)
        if group and (output or defects or cycle):
            metrics=[m for m in (output,defects,cycle) if m]; grouped=aggregate(group,metrics); result=[]
            for name,item in sorted(grouped.items(),key=lambda x:-x[1]["sums"].get(output["name"],0) if output else x[0]):
                out=item["sums"].get(output["name"],0) if output else None; defect=item["sums"].get(defects["name"],0) if defects else None
                n=item["counts"].get(cycle["name"],0) if cycle else 0; cyc=item["sums"].get(cycle["name"],0)/n if n else None
                result.append([name,round(out,2) if out is not None else None,round(defect,2) if defect is not None else None,round(defect/out*100,2) if defect is not None and out else None,round(cyc,2) if cyc is not None else None])
            return {"title":f"Operational performance by {group['label'].lower()}","description":"Compare throughput, defect rate and average cycle time where available.","headers":[group["label"],"Throughput","Defects","Defect rate %","Mean cycle time"],"rows":result[:100],"footnote":"Defect rate is total defects divided by total output. Confirm that both fields use compatible units."}
    return None

def make_sample(category):
    import random
    rng=random.Random(24); rows=[]
    if category=="education":
        cols=[{"name":"student_id","label":"Student ID","kind":"category"},{"name":"course","label":"Course","kind":"category"},{"name":"term","label":"Term","kind":"date"},{"name":"score","label":"Score","kind":"number"},{"name":"attendance_pct","label":"Attendance %","kind":"number"}]
        for i in range(180): rows.append({"student_id":f"STU-{i+1:03}","course":rng.choice(["Mathematics","Biology","History","Computer Science"]),"term":f"2025-{1+(i%2)*6:02}-01","score":round(max(25,min(100,rng.gauss(76,12))),1),"attendance_pct":round(rng.uniform(62,100),1)})
        return "Sample academic outcomes.csv",cols,rows
    if category=="customers":
        cols=[{"name":"customer_id","label":"Customer ID","kind":"category"},{"name":"order_id","label":"Order ID","kind":"category"},{"name":"order_date","label":"Order date","kind":"date"},{"name":"segment","label":"Current segment","kind":"category"},{"name":"age","label":"Age","kind":"number"},{"name":"order_value","label":"Order value","kind":"number"}]
        profiles={i:{"segment":rng.choice(["Loyal","New","At risk","High value"]),"age":rng.randint(18,76)} for i in range(1,121)}
        for i in range(300):
            cid=rng.randint(1,120); profile=profiles[cid]; value=round(rng.lognormvariate(4.4,.65),2); day=rng.randint(0,364)
            rows.append({"customer_id":f"C-{cid:04}","order_id":f"ORD-{i+1:05}","order_date":(datetime(2025,1,1)+timedelta(days=day)).strftime("%Y-%m-%d"),"segment":profile["segment"],"age":profile["age"],"order_value":value})
        return "Sample RFM customer orders.csv",cols,rows
    if category=="retail":
        cols=[{"name":"date","label":"Date","kind":"date"},{"name":"product","label":"Product","kind":"category"},{"name":"category","label":"Category","kind":"category"},{"name":"quantity","label":"Quantity","kind":"number"},{"name":"sales","label":"Sales","kind":"number"},{"name":"payment_method","label":"Payment method","kind":"category"}]
        for i in range(240):
            q=rng.randint(1,8); p=rng.choice([("Bread","Grocery",2.5),("Coffee","Grocery",4.2),("Notebook","Stationery",3.5),("Soap","Household",2.8),("Milk","Grocery",1.9)])
            rows.append({"date":(datetime(2025,1,1)+timedelta(days=i%90)).strftime("%Y-%m-%d"),"product":p[0],"category":p[1],"quantity":q,"sales":round(q*p[2],2),"payment_method":rng.choice(["Cash","Card","UPI"])})
        return "Sample neighbourhood shop.csv",cols,rows
    if category=="finance":
        cols=[{"name":"date","label":"Month","kind":"date"},{"name":"department","label":"Department","kind":"category"},{"name":"account","label":"Account","kind":"category"},{"name":"actual","label":"Actual spend","kind":"number"},{"name":"budget","label":"Budget","kind":"number"}]
        for i in range(150):
            budget=rng.randint(7000,42000); actual=budget*rng.uniform(.68,1.3)
            rows.append({"date":(datetime(2025,1,1)+timedelta(days=30*(i%12))).strftime("%Y-%m-%d"),"department":rng.choice(["Sales","Engineering","People","Operations","Marketing"]),"account":rng.choice(["Payroll","Software","Travel","Supplies"]),"actual":round(actual,2),"budget":budget})
        return "Sample budget vs actual.csv",cols,rows
    if category=="marketing":
        cols=[{"name":"date","label":"Date","kind":"date"},{"name":"campaign","label":"Campaign","kind":"category"},{"name":"channel","label":"Channel","kind":"category"},{"name":"spend","label":"Spend","kind":"number"},{"name":"impressions","label":"Impressions","kind":"number"},{"name":"conversions","label":"Conversions","kind":"number"},{"name":"revenue","label":"Attributed revenue","kind":"number"}]
        for i in range(220):
            spend=round(rng.uniform(40,950),2); conv=rng.randint(0,55)
            rows.append({"date":(datetime(2025,1,1)+timedelta(days=i%120)).strftime("%Y-%m-%d"),"campaign":rng.choice(["Spring launch","Always-on","Referral","Retargeting"]),"channel":rng.choice(["Search","Social","Email","Partner"]),"spend":spend,"impressions":rng.randint(2000,60000),"conversions":conv,"revenue":round(conv*rng.uniform(18,95),2)})
        return "Sample marketing performance.csv",cols,rows
    if category=="product":
        cols=[{"name":"date","label":"Date","kind":"date"},{"name":"feature","label":"Feature","kind":"category"},{"name":"plan","label":"Plan","kind":"category"},{"name":"active_users","label":"Active users","kind":"number"},{"name":"events","label":"Events","kind":"number"},{"name":"tickets","label":"Support tickets","kind":"number"}]
        for i in range(180): rows.append({"date":(datetime(2025,1,1)+timedelta(days=i%90)).strftime("%Y-%m-%d"),"feature":rng.choice(["Search","Reports","Sharing","Import","Automation"]),"plan":rng.choice(["Free","Team","Business"]),"active_users":rng.randint(25,1200),"events":rng.randint(120,9000),"tickets":rng.randint(0,60)})
        return "Sample product usage.csv",cols,rows
    if category=="operations":
        cols=[{"name":"date","label":"Date","kind":"date"},{"name":"team","label":"Team","kind":"category"},{"name":"shift","label":"Shift","kind":"category"},{"name":"tasks_completed","label":"Tasks completed","kind":"number"},{"name":"hours","label":"Hours","kind":"number"},{"name":"defects","label":"Defects","kind":"number"}]
        for i in range(180): rows.append({"date":(datetime(2025,1,1)+timedelta(days=i%60)).strftime("%Y-%m-%d"),"team":rng.choice(["Alpha","Bravo","Charlie","Delta"]),"shift":rng.choice(["Morning","Evening","Night"]),"tasks_completed":rng.randint(20,180),"hours":round(rng.uniform(5.5,9),1),"defects":rng.randint(0,12)})
        return "Sample operations metrics.csv",cols,rows
    if category=="logistics":
        cols=[{"name":"date","label":"Ship date","kind":"date"},{"name":"route","label":"Route","kind":"category"},{"name":"carrier","label":"Carrier","kind":"category"},{"name":"delivery_days","label":"Delivery days","kind":"number"},{"name":"distance_km","label":"Distance km","kind":"number"},{"name":"delivered_on_time","label":"On time deliveries","kind":"number"}]
        for i in range(200): rows.append({"date":(datetime(2025,1,1)+timedelta(days=i%90)).strftime("%Y-%m-%d"),"route":rng.choice(["North","South","East","West"]),"carrier":rng.choice(["Fleet A","Fleet B","Fleet C"]),"delivery_days":round(rng.uniform(1,9),1),"distance_km":rng.randint(30,1800),"delivered_on_time":rng.choice([0,1])})
        return "Sample delivery performance.csv",cols,rows
    if category=="research":
        cols=[{"name":"response_id","label":"Response ID","kind":"category"},{"name":"cohort","label":"Cohort","kind":"category"},{"name":"question","label":"Question","kind":"category"},{"name":"response_score","label":"Response score","kind":"number"},{"name":"duration_seconds","label":"Duration seconds","kind":"number"}]
        for i in range(240): rows.append({"response_id":f"R-{i+1:04}","cohort":rng.choice(["Control","Variant A","Variant B"]),"question":rng.choice(["Ease of use","Trust","Value","Satisfaction"]),"response_score":rng.randint(1,10),"duration_seconds":rng.randint(8,220)})
        return "Sample survey responses.csv",cols,rows
    cols=[{"name":"date","label":"Order date","kind":"date"},{"name":"ship_date","label":"Ship date","kind":"date"},{"name":"order_id","label":"Order ID","kind":"category"},{"name":"product","label":"Product","kind":"category"},{"name":"category","label":"Category","kind":"category"},{"name":"channel","label":"Channel","kind":"category"},{"name":"customer_id","label":"Customer ID","kind":"category"},{"name":"quantity","label":"Quantity","kind":"number"},{"name":"sales","label":"Sales","kind":"number"},{"name":"profit","label":"Profit","kind":"number"},{"name":"discount","label":"Discount","kind":"number"}]
    for i in range(240):
        product=rng.choice(["Canvas tote","Desk lamp","Water bottle","Planner"]); q=rng.randint(1,5); price={"Canvas tote":18,"Desk lamp":42,"Water bottle":24,"Planner":15}[product]; order_date=datetime(2025,1,1)+timedelta(days=i%180); discount=rng.choice([0,.05,.1,.15,.2]); sales=round(q*price*(1-discount),2); profit=round(sales*rng.uniform(.12,.34)-discount*price*q*.25,2)
        rows.append({"date":order_date.strftime("%Y-%m-%d"),"ship_date":(order_date+timedelta(days=rng.randint(1,8))).strftime("%Y-%m-%d"),"order_id":f"ORD-{i+1:05}","product":product,"category":"Accessories" if product in ("Canvas tote","Planner") else "Home & lifestyle","channel":rng.choice(["Website","Marketplace","Social"]),"customer_id":f"C-{rng.randint(1,95):03}","quantity":q,"sales":sales,"profit":profit,"discount":discount})
    return "Sample ecommerce orders.csv",cols,rows

class Handler(BaseHTTPRequestHandler):
    server_version = "DataPrism/2"
    def log_message(self, fmt, *args): print(f"[{self.log_date_time_string()}] {fmt % args}")
    def send_body(self, status, body, content_type="application/json; charset=utf-8", headers=None):
        data=body if isinstance(body,bytes) else body.encode("utf-8")
        self.send_response(status); self.send_header("Content-Type",content_type); self.send_header("Content-Length",str(len(data))); self.send_header("Cache-Control","no-store")
        for k,v in (headers or {}).items(): self.send_header(k,v)
        self.end_headers(); self.wfile.write(data)
    def json(self, status, data, headers=None): self.send_body(status,json.dumps(data,ensure_ascii=False),headers=headers)
    def body_json(self):
        length=int(self.headers.get("Content-Length",0))
        if length>MAX_UPLOAD: raise ValueError("Request is too large.")
        return json.loads(self.rfile.read(length) or b"{}")
    def user(self, con):
        raw=next((x.strip() for x in self.headers.get("Cookie","").split(";") if x.strip().startswith("mm_session=")),None)
        if not raw: return None
        token=raw.split("=",1)[1]
        r=con.execute("SELECT u.* FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token_hash=? AND s.expires_at>?",(hashlib.sha256(token.encode()).hexdigest(),utcnow().isoformat())).fetchone()
        return dict(r) if r else None
    def require_user(self, con):
        u=self.user(con)
        if not u: self.json(401,{"error":"Please sign in to continue."}); return None
        return u
    def session_cookie(self, user_id):
        token=secrets.token_urlsafe(36); con=connect(); con.execute("DELETE FROM sessions WHERE expires_at<=?",(utcnow().isoformat(),)); con.execute("INSERT INTO sessions(token_hash,user_id,expires_at) VALUES(?,?,?)",(hashlib.sha256(token.encode()).hexdigest(),user_id,(utcnow()+timedelta(days=SESSION_DAYS)).isoformat())); con.commit(); con.close()
        secure="; Secure" if self.headers.get("X-Forwarded-Proto","").lower()=="https" else ""
        return {"Set-Cookie":f"mm_session={token}; HttpOnly; SameSite=Lax; Path=/; Max-Age={SESSION_DAYS*86400}{secure}"}
    def do_GET(self):
        path=urllib.parse.urlparse(self.path).path
        if path=="/api/health":
            return self.json(200,{"app":"DataPrism","build":"5.0"})
        if path=="/api/me":
            con=connect(); u=self.user(con); con.close()
            if not u: return self.json(200,{"user":None})
            return self.json(200,{"user":{"email":u["email"],"category":u["category"],"category_name":CATEGORIES.get(u["category"],"")}})
        if path=="/api/datasets":
            con=connect(); u=self.require_user(con)
            if not u: con.close(); return
            ds=[serialize_dataset(x) for x in con.execute("SELECT * FROM datasets WHERE user_id=? ORDER BY id DESC",(u["id"],))]; con.close(); return self.json(200,{"datasets":ds})
        if path=="/api/dashboard": return self.get_dashboard()
        if path=="/api/preview": return self.get_preview()
        if path=="/api/chart": return self.get_chart()
        if path=="/api/saved-charts": return self.get_saved_charts()
        if path=="/api/export": return self.get_export()
        if path=="/" or path=="/index.html":
            target=ROOT/"platform.html"
        else: target=(ROOT/path.lstrip("/")).resolve()
        if ROOT not in target.parents and target!=ROOT: return self.send_error(403)
        if not target.is_file(): return self.send_error(404)
        return self.send_body(200,target.read_bytes(),mimetypes.guess_type(target.name)[0] or "application/octet-stream")
    def get_dashboard(self):
        q=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query); did=int(q.get("dataset",[0])[0] or 0)
        con=connect(); u=self.require_user(con)
        if not u: con.close(); return
        if not u["category"]: con.close(); return self.json(409,{"error":"Choose an analytics workspace first."})
        ds=con.execute("SELECT * FROM datasets WHERE id=? AND user_id=? AND category=?",(did,u["id"],u["category"])).fetchone() if did else con.execute("SELECT * FROM datasets WHERE user_id=? AND category=? ORDER BY id DESC LIMIT 1",(u["id"],u["category"])).fetchone()
        all_ds=[serialize_dataset(x) for x in con.execute("SELECT * FROM datasets WHERE user_id=? AND category=? ORDER BY id DESC",(u["id"],u["category"]))]
        if not ds: con.close(); return self.json(200,{"user":{"email":u["email"],"category":u["category"],"category_name":CATEGORIES[u["category"]]},"datasets":all_ds,"dataset":None})
        d=serialize_dataset(ds); all_rows=read_records(con,d["id"]); params=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query); rows=apply_filter(all_rows,d["columns"],params)
        charts=[dict(x) for x in con.execute("SELECT id,title,config_json,created_at FROM saved_charts WHERE user_id=? AND dataset_id=? ORDER BY id DESC",(u["id"],d["id"]))]
        for ch in charts: ch["config"]=json.loads(ch.pop("config_json"))
        con.close()
        cards,primary,dims,dt=summary(rows,d["columns"],d["category"]); profile=profile_data(rows,d["columns"])
        bykind={k:[c for c in d["columns"] if c["kind"]==k] for k in ("number","category","text","date")}
        specialized={}
        if d["category"] in ("customers","ecommerce"): specialized["rfm"]=rfm_analysis(rows,d["columns"])
        specialized["category_analysis"]=category_analysis(rows,d["columns"],d["category"])
        correlations=correlation_data(rows,d["columns"])
        return self.json(200,{"user":{"email":u["email"],"category":u["category"],"category_name":CATEGORIES[u["category"]]},"datasets":all_ds,"dataset":d,"cards":cards,"columns":d["columns"],"by_kind":bykind,"rows":rows[:80],"total_rows":len(all_rows),"filtered_rows":len(rows),"preview_truncated":len(rows)>80,"profile":profile,"specialized":specialized,"correlations":correlations,"saved_charts":charts,"active_filter":{"field":params.get("filter_field",[""])[0],"op":params.get("filter_op",["eq"])[0],"value":params.get("filter_value",[""])[0]},"defaults":{"x":(dt or (dims[0] if dims else d["columns"][0]))["name"],"y":primary["name"] if primary else "","chart":"bar","aggregation":"sum"}})
    def get_preview(self):
        q=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        try: did=int(q.get("dataset",[0])[0] or 0); page=max(1,int(q.get("page",[1])[0])); size=min(100,max(10,int(q.get("page_size",[50])[0])))
        except (TypeError,ValueError): return self.json(400,{"error":"Preview page settings must be whole numbers."})
        term=q.get("search",[""])[0].strip().casefold()
        con=connect(); u=self.require_user(con)
        if not u: con.close(); return
        d=dataset_row(con,u["id"],did)
        if not d: con.close(); return self.json(404,{"error":"Dataset not found."})
        cols=json.loads(d["columns_json"]); has_filter=bool(q.get("filter_field",[""])[0])
        if not has_filter and not term:
            total=d["row_count"]; pages=max(1,math.ceil(total/size)); page=min(page,pages); start=(page-1)*size
            page_rows=[json.loads(row[0]) for row in con.execute("SELECT payload FROM dataset_rows WHERE dataset_id=? ORDER BY row_num LIMIT ? OFFSET ?",(did,size,start))]
            con.close(); return self.json(200,{"columns":cols,"rows":page_rows,"total_rows":total,"page":page,"page_size":size,"pages":pages,"search":q.get("search",[""])[0]})
        all_rows=apply_filter(read_records(con,did),cols,q); con.close()
        if term: all_rows=[r for r in all_rows if any(term in str(r.get(c["name"],"")).casefold() for c in cols)]
        total=len(all_rows); pages=max(1,math.ceil(total/size)); page=min(page,pages); start=(page-1)*size
        return self.json(200,{"columns":cols,"rows":all_rows[start:start+size],"total_rows":total,"page":page,"page_size":size,"pages":pages,"search":q.get("search",[""])[0]})
    def get_chart(self):
        q=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query); did=int(q.get("dataset",[0])[0] or 0); x=q.get("x",[""])[0]; y=q.get("y",[""])[0]; typ=q.get("type",["bar"])[0]; agg=q.get("agg",["sum"])[0]
        con=connect(); u=self.require_user(con)
        if not u: con.close(); return
        d=dataset_row(con,u["id"],did)
        if not d: con.close(); return self.json(404,{"error":"Dataset not found."})
        cols=json.loads(d["columns_json"]); names={c["name"]:c for c in cols}
        if typ not in ("bar","line","scatter","histogram","donut"): con.close(); return self.json(400,{"error":"Choose a supported chart type."})
        if x not in names and typ not in ("histogram",): con.close(); return self.json(400,{"error":"Select an x-axis column."})
        if typ not in ("histogram",) and y not in names and agg!="count": con.close(); return self.json(400,{"error":"Select a y-axis column."})
        if typ in ("bar","line","donut") and agg!="count" and names[y]["kind"]!="number": con.close(); return self.json(400,{"error":"Choose a numeric field for the value axis."})
        rows=apply_filter(read_records(con,did),cols,q); con.close()
        if typ=="scatter":
            if names[x]["kind"]!="number" or names[y]["kind"]!="number": return self.json(400,{"error":"Scatter plots need numeric x and y columns."})
            pts=[{"x":r.get(x),"y":r.get(y)} for r in rows if isinstance(r.get(x),(int,float)) and isinstance(r.get(y),(int,float))]
            if len(pts)>900:
                step=len(pts)/900; pts=[pts[int(i*step)] for i in range(900)]
            return self.json(200,{"type":typ,"x_field":x,"x_label":names[x]["label"],"y_label":names[y]["label"],"points":pts})
        if typ=="histogram":
            if y not in names or names[y]["kind"]!="number": return self.json(400,{"error":"Choose a numeric column to show its distribution."})
            vals=[r[y] for r in rows if isinstance(r.get(y),(int,float))]
            if not vals: return self.json(200,{"type":typ,"x_label":names[y]["label"],"points":[]})
            lo,hi=min(vals),max(vals); n=1 if lo==hi else 18; width=(hi-lo)/n or 1; bins=[0]*n
            for v in vals: bins[min(n-1,int((v-lo)/width))]+=1
            pts=[{"x":f"{lo+i*width:.3g}–{lo+(i+1)*width:.3g}","y":v} for i,v in enumerate(bins)]
            return self.json(200,{"type":typ,"x_label":names[y]["label"],"y_label":"Records","points":pts})
        groups={}
        for r in rows:
            key=r.get(x)
            if key is None: key="(blank)"
            key=str(key)
            g=groups.setdefault(key,[])
            v=r.get(y) if y else None
            if agg=="count": g.append(1)
            elif isinstance(v,(int,float)): g.append(v)
        def aggregate(v):
            if not v:return 0
            if agg=="avg":return sum(v)/len(v)
            if agg=="median":return sorted(v)[len(v)//2] if len(v)%2 else (sorted(v)[len(v)//2-1]+sorted(v)[len(v)//2])/2
            if agg=="min":return min(v)
            if agg=="max":return max(v)
            if agg=="distinct":return len(set(v))
            return sum(v)
        points=[{"x":k,"y":aggregate(v)} for k,v in groups.items()]
        if names.get(x,{}).get("kind")=="date": points.sort(key=lambda p:p["x"])
        else: points.sort(key=lambda p:p["y"],reverse=True)
        if len(points)>80: points=points[:80]
        return self.json(200,{"type":typ,"x_field":x,"x_label":names.get(x,{}).get("label",x),"y_label":"Records" if agg=="count" else names.get(y,{}).get("label",y),"aggregation":agg,"points":points})
    def get_export(self):
        q=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query); did=int(q.get("dataset",[0])[0] or 0)
        con=connect(); u=self.require_user(con)
        if not u: con.close(); return
        d=dataset_row(con,u["id"],did)
        if not d: con.close(); return self.json(404,{"error":"Dataset not found."})
        cols=json.loads(d["columns_json"]); rows=apply_filter(read_records(con,did),cols,q); con.close(); out=StringIO(); w=csv.writer(out); w.writerow([c["label"] for c in cols]); w.writerows([[r.get(c["name"]) for c in cols] for r in rows]); blob=out.getvalue().encode("utf-8-sig")
        return self.send_body(200,blob,"text/csv; charset=utf-8",{"Content-Disposition":f'attachment; filename="{Path(d["name"]).stem}-export.csv"'})
    def get_saved_charts(self):
        q=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query); did=int(q.get("dataset",[0])[0] or 0)
        con=connect(); u=self.require_user(con)
        if not u: con.close(); return
        if not dataset_row(con,u["id"],did): con.close(); return self.json(404,{"error":"Dataset not found."})
        charts=[dict(x) for x in con.execute("SELECT id,title,config_json,created_at FROM saved_charts WHERE user_id=? AND dataset_id=? ORDER BY id DESC",(u["id"],did))]
        for ch in charts: ch["config"]=json.loads(ch.pop("config_json"))
        con.close(); return self.json(200,{"charts":charts})
    def do_POST(self):
        path=urllib.parse.urlparse(self.path).path
        try:
            if path in ("/api/auth/signup","/api/auth/login"):
                data=self.body_json(); email=str(data.get("email","")).strip().lower(); password=str(data.get("password",""))
                if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+",email): return self.json(400,{"error":"Enter a valid email address."})
                if len(password)<10: return self.json(400,{"error":"Use a password with at least 10 characters."})
                con=connect()
                if path.endswith("signup"):
                    salt=secrets.token_bytes(16); digest=hashlib.pbkdf2_hmac("sha256",password.encode(),salt,PBKDF2_ROUNDS)
                    try: cur=con.execute("INSERT INTO users(email,salt,password_hash,category,created_at) VALUES(?,?,?,NULL,?)",(email,salt.hex(),digest.hex(),utcnow().isoformat())); con.commit()
                    except sqlite3.IntegrityError: con.close(); return self.json(409,{"error":"An account with this email already exists. Sign in instead."})
                    uid=cur.lastrowid
                else:
                    row=con.execute("SELECT id,salt,password_hash FROM users WHERE email=?",(email,)).fetchone()
                    digest=hashlib.pbkdf2_hmac("sha256",password.encode(),bytes.fromhex(row["salt"]),PBKDF2_ROUNDS) if row else b""
                    if not row or not hmac.compare_digest(digest.hex(),row["password_hash"]): con.close(); return self.json(401,{"error":"Email or password is incorrect."})
                    uid=row["id"]
                con.close(); return self.json(200,{"ok":True},self.session_cookie(uid))
            if path=="/api/auth/logout":
                con=connect(); raw=next((x.strip().split("=",1)[1] for x in self.headers.get("Cookie","").split(";") if x.strip().startswith("mm_session=")),None)
                if raw: con.execute("DELETE FROM sessions WHERE token_hash=?",(hashlib.sha256(raw.encode()).hexdigest(),)); con.commit()
                con.close(); return self.json(200,{"ok":True},{"Set-Cookie":"mm_session=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0"})
            con=connect(); u=self.require_user(con)
            if not u: con.close(); return
            if path=="/api/category":
                data=self.body_json(); cat=data.get("category")
                if cat not in CATEGORIES: con.close(); return self.json(400,{"error":"Choose one of the listed workspace categories."})
                con.execute("UPDATE users SET category=? WHERE id=?",(cat,u["id"])); con.commit(); con.close(); return self.json(200,{"ok":True})
            if path=="/api/chart/save":
                data=self.body_json(); did=int(data.get("dataset",0)); ds=dataset_row(con,u["id"],did)
                if not ds: con.close(); return self.json(404,{"error":"Dataset not found."})
                cols=json.loads(ds["columns_json"]); names={c["name"] for c in cols}; config=data.get("config",{})
                if config.get("type") not in ("bar","line","scatter","histogram","donut") or config.get("x") not in names or (config.get("y") and config["y"] not in names): con.close(); return self.json(400,{"error":"Chart settings are incomplete."})
                title=str(data.get("title","")).strip()[:100]
                if not title: con.close(); return self.json(400,{"error":"Give this saved chart a name."})
                cur=con.execute("INSERT INTO saved_charts(user_id,dataset_id,title,config_json,created_at) VALUES(?,?,?,?,?)",(u["id"],did,title,json.dumps(config),utcnow().isoformat())); con.commit(); chart_id=cur.lastrowid; con.close(); return self.json(200,{"ok":True,"id":chart_id})
            if path=="/api/chart/delete":
                data=self.body_json(); cur=con.execute("DELETE FROM saved_charts WHERE id=? AND user_id=?",(int(data.get("id",0)),u["id"])); con.commit(); con.close(); return self.json(200,{"ok":True}) if cur.rowcount else self.json(404,{"error":"Saved chart not found."})
            if path=="/api/transform":
                data=self.body_json(); did=int(data.get("dataset",0)); ds=dataset_row(con,u["id"],did)
                if not ds: con.close(); return self.json(404,{"error":"Dataset not found."})
                cols=json.loads(ds["columns_json"]); rows=read_records(con,did); by={c["name"]:c for c in cols}; action=data.get("action"); note=""
                if action=="remove_duplicates":
                    seen=set(); cleaned=[]
                    for row in rows:
                        key=json.dumps(row,sort_keys=True,ensure_ascii=False)
                        if key not in seen: seen.add(key); cleaned.append(row)
                    note=f"Removed {len(rows)-len(cleaned)} exact duplicate rows"; rows=cleaned
                elif action in ("trim","lower","upper","title","fill_missing","drop_missing"):
                    field=data.get("field"); value=data.get("value","")
                    if field not in by: con.close(); return self.json(400,{"error":"Choose a valid column."})
                    if action=="drop_missing":
                        c=by[field]; before=len(rows)
                        rows=[r for r in rows if r.get(field) is not None and str(r.get(field)).strip()!="" and not (c["kind"]=="number" and to_number(r.get(field)) is None) and not (c["kind"]=="date" and to_date(r.get(field)) is None)]
                        c["invalid_count"]=0; note=f"Removed {before-len(rows)} rows missing or invalid in {c['label']}"
                    elif action=="fill_missing":
                        c=by[field]
                        replacement=to_number(value) if c["kind"]=="number" else to_date(value) if c["kind"]=="date" else str(value)
                        if replacement is None: con.close(); return self.json(400,{"error":"Enter a replacement value matching the column type."})
                        count=0
                        for r in rows:
                            invalid=c["kind"]=="number" and to_number(r.get(field)) is None or c["kind"]=="date" and to_date(r.get(field)) is None
                            if r.get(field) is None or str(r.get(field)).strip()=="" or invalid: r[field]=replacement; count+=1
                        c["invalid_count"]=0; note=f"Filled {count} missing values in {c['label']}"
                    else:
                        if by[field]["kind"] not in ("category","text"): con.close(); return self.json(400,{"error":"Text cleanup needs a text or category column."})
                        for r in rows:
                            if isinstance(r.get(field),str): r[field]=r[field].strip() if action=="trim" else r[field].lower() if action=="lower" else r[field].upper() if action=="upper" else r[field].title()
                        note=f"Applied {action} to {by[field]['label']}"
                elif action=="fill_stat":
                    field=data.get("field"); method=data.get("method")
                    if field not in by or by[field]["kind"]!="number": con.close(); return self.json(400,{"error":"Mean or median replacement needs a numeric column."})
                    if method not in ("mean","median"): con.close(); return self.json(400,{"error":"Choose mean or median replacement."})
                    values=[float(r[field]) for r in rows if isinstance(r.get(field),(int,float)) and math.isfinite(float(r[field]))]
                    if not values: con.close(); return self.json(400,{"error":"This column has no valid numeric values to calculate a replacement."})
                    replacement=sum(values)/len(values) if method=="mean" else (sorted(values)[(len(values)-1)//2]+sorted(values)[len(values)//2])/2
                    count=0
                    for r in rows:
                        value=r.get(field)
                        if value is None or str(value).strip()=="" or to_number(value) is None: r[field]=replacement; count+=1
                    if count==0: con.close(); return self.json(400,{"error":"This column has no missing or invalid values to replace."})
                    by[field]["invalid_count"]=0; note=f"Filled {count} missing or invalid values in {by[field]['label']} with the {method} ({replacement:,.4g})"
                elif action=="derive":
                    a,b=data.get("field_a"),data.get("field_b"); op=data.get("operation"); label=str(data.get("name","")).strip() or "Calculated field"; used={c["name"] for c in cols}; key=norm_name(label,used)
                    if a not in by or b not in by or by[a]["kind"]!="number" or by[b]["kind"]!="number": con.close(); return self.json(400,{"error":"Calculated fields need two numeric columns."})
                    if op not in ("add","subtract","multiply","divide"): con.close(); return self.json(400,{"error":"Choose a supported calculation."})
                    for r in rows:
                        av,bv=r.get(a),r.get(b)
                        if not isinstance(av,(int,float)) or not isinstance(bv,(int,float)) or (op=="divide" and bv==0): r[key]=None
                        else: r[key]=av+bv if op=="add" else av-bv if op=="subtract" else av*bv if op=="multiply" else av/bv
                    cols.append({"name":key,"label":label,"kind":"number","invalid_count":0}); note=f"Added calculated field {label} using {a} {op} {b}"
                elif action=="date_diff":
                    start,end=data.get("date_a"),data.get("date_b"); label=str(data.get("name","Days between dates")).strip() or "Days between dates"; used={c["name"] for c in cols}; key=norm_name(label,used)
                    if start not in by or end not in by or by[start]["kind"]!="date" or by[end]["kind"]!="date": con.close(); return self.json(400,{"error":"Choose two date columns."})
                    for r in rows:
                        a,b=to_date(r.get(start)),to_date(r.get(end))
                        r[key]=(datetime.strptime(b,"%Y-%m-%d")-datetime.strptime(a,"%Y-%m-%d")).days if a and b else None
                    cols.append({"name":key,"label":label,"kind":"number","invalid_count":0}); note=f"Added {label} from {by[start]['label']} to {by[end]['label']}"
                elif action=="join":
                    other_id=int(data.get("dataset_b",0)); other=dataset_row(con,u["id"],other_id)
                    if not other or other["category"]!=ds["category"]: con.close(); return self.json(400,{"error":"Choose another dataset in this workspace."})
                    other_cols=json.loads(other["columns_json"]); other_names={c["name"] for c in other_cols}; key_a=data.get("key_a"); key_b=data.get("key_b")
                    if key_a not in by or key_b not in other_names: con.close(); return self.json(400,{"error":"Choose valid join columns."})
                    match={}; right_rows=read_records(con,other_id)
                    for rr in right_rows:
                        k=rr.get(key_b)
                        if k is not None: match.setdefault(str(k).strip().casefold(),[]).append(rr)
                    join_type=data.get("join_type","left"); used={c["name"] for c in cols}; right_fields=[]
                    for c in other_cols:
                        if c["name"]==key_b: continue
                        new_name=norm_name(f"{Path(other['name']).stem}_{c['name']}",used); right_fields.append((c,new_name))
                    joined=[]; matched=0
                    for left in rows:
                        key=left.get(key_a); candidates=match.get(str(key).strip().casefold(),[]) if key is not None else []
                        if candidates: matched+=1
                        elif join_type=="inner": continue
                        else: candidates=[{}]
                        for right in candidates:
                            merged=dict(left)
                            for c,new_name in right_fields: merged[new_name]=right.get(c["name"])
                            joined.append(merged)
                            if len(joined)>MAX_ROWS: con.close(); return self.json(400,{"error":"This join would exceed 100,000 rows. Filter or deduplicate the second dataset first."})
                    if join_type not in ("left","inner"): con.close(); return self.json(400,{"error":"Choose a left or inner join."})
                    cols.extend([{**c,"name":new_name,"label":f"{Path(other['name']).stem}: {c['label']}","invalid_count":c.get("invalid_count",0)} for c,new_name in right_fields]); rows=joined
                    note=f"{join_type.title()} join with {other['name']} on {key_a} = {key_b}; {matched:,} left-side keys matched" 
                else: con.close(); return self.json(400,{"error":"Choose a supported transformation."})
                stem=Path(ds["name"]); new_name=f"{stem.stem} · cleaned{stem.suffix or '.csv'}"[:180]
                new_id=save_dataset(con,u["id"],new_name,ds["category"],cols,rows,parent_id=did,transform_note=note); con.close(); return self.json(200,{"ok":True,"dataset_id":new_id,"rows":len(rows),"note":note})
            if path=="/api/sample":
                cat=u["category"]
                if not cat: con.close(); return self.json(409,{"error":"Choose a workspace first."})
                name,cols,rows=make_sample(cat); did=save_dataset(con,u["id"],name,cat,cols,rows); con.close(); return self.json(200,{"ok":True,"dataset_id":did,"rows":len(rows)})
            if path=="/api/upload":
                length=int(self.headers.get("Content-Length",0))
                if not length or length>MAX_UPLOAD: con.close(); return self.json(413,{"error":"File is empty or larger than 50 MB."})
                envelope=(f"Content-Type: {self.headers.get('Content-Type')}\r\nMIME-Version: 1.0\r\n\r\n").encode()+self.rfile.read(length)
                msg=BytesParser(policy=default).parsebytes(envelope); part=next((p for p in msg.iter_parts() if p.get_filename()),None)
                if not part: con.close(); return self.json(400,{"error":"Choose a CSV or Excel file to upload."})
                filename=Path(part.get_filename()).name; cols,rows=decode_upload(filename,part.get_payload(decode=True) or b"")
                did=save_dataset(con,u["id"],filename,u["category"] or "general",cols,rows); con.close(); return self.json(200,{"ok":True,"dataset_id":did,"rows":len(rows)})
            if path=="/api/dataset/delete":
                data=self.body_json(); did=int(data.get("dataset",0)); cur=con.execute("DELETE FROM datasets WHERE id=? AND user_id=?",(did,u["id"])); con.commit(); con.close()
                return self.json(200,{"ok":True}) if cur.rowcount else self.json(404,{"error":"Dataset not found."})
            con.close(); return self.send_error(404)
        except (ValueError,StopIteration,json.JSONDecodeError) as e:
            try: con.close()
            except Exception: pass
            return self.json(400,{"error":str(e) or "Could not read that request."})
        except Exception as e:
            try: con.close()
            except Exception: pass
            print(f"Request failed: {e}")
            return self.json(500,{"error":"The request could not be completed. Check the file and try again."})

if __name__=="__main__":
    host=os.environ.get("DATAPRISM_HOST","127.0.0.1"); port=int(os.environ.get("PORT",sys.argv[1] if len(sys.argv)>1 else "8780"))
    connect().close()
    print(f"DataPrism platform running at http://{host}:{port} — database: {DB}")
    ThreadingHTTPServer((host,port),Handler).serve_forever()
