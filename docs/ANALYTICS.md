# Analytics Methodology

This guide describes how DataPrism prepares, summarizes, and visualizes uploaded tabular data. The app is an exploratory analysis aid: it reports patterns in the supplied rows and does not establish causation or guarantee that a dataset represents a wider population.

## 1. Import and field detection

Supported input includes CSV, TSV, XLSX/XLSM, JSON records, and JSONL/NDJSON. Blank lines may be skipped during import. The parser normalizes rows into records and records import diagnostics.

Columns are inferred from their contents as number, date, category, identifier, or text. Numeric and date inference uses a 75% valid-parse threshold over non-empty values. Identifier heuristics consider both column names and value patterns. A detected type is a useful default, not a claim about the field’s business meaning; inspect the preview and adjust column roles where the workspace allows it.

The configured limits are 50 MB per upload and 100,000 rows. When input exceeds the row limit, the server keeps the first configured number of rows and reports the truncation.

## 2. Dataset Health

Dataset Health reports row and column counts, completeness, missing cells, distinct values, and exact duplicate rows. Duplicate detection compares the complete normalized row; it does not decide whether a repeated business event is valid or accidental.

For each field, the report may include sample values, inferred type, missing percentage, and values that could not be parsed according to the inferred type. Invalid values remain visible in the preview but are skipped by calculations that require that type.

Missing values are not treated as zero. Calculations using a field omit rows without a usable value for that calculation. The report distinguishes diagnostics from cleaning actions.

## 3. Numeric Summary

For each eligible quantitative field, DataPrism reports valid and missing counts, arithmetic mean, median, sample standard deviation, sample variance, minimum, first quartile, third quartile, and maximum.

- Mean: sum of valid values divided by their count.
- Median: middle value after sorting; for an even count, the midpoint of the two middle values.
- Sample variance: squared deviations from the mean divided by n − 1. It is shown as zero when fewer than two valid values exist.
- Sample standard deviation: square root of sample variance.
- Quartiles: inclusive quartiles from the Python statistics implementation.
- Interquartile range: Q3 − Q1. The conventional 1.5 × IQR fences are used when describing potential outliers.

Identifier columns are excluded from ordinary numeric summaries. Mean and variance are sensitive to extreme values; compare them with the median and interquartile range. The outlier flag is a prompt to inspect a value, not proof that it is erroneous.

## 4. Category Breakdown

For useful categorical fields, the breakdown reports frequency and share, along with common values and distinct-value counts. Shares are calculated from valid values for the field; missing records are reported separately. Identifier-like or near-unique fields are not useful groupings and may be excluded from category rankings.

A high-frequency category is not necessarily important or causal. Check whether category labels are consistent and whether small groups are meaningful before comparing them.

## 5. Trends Over Time

Time trends require at least one date field and can use a numeric measure or record count. The user selects the date field, measure, grain, and aggregation where available. Supported grains include daily, weekly, monthly, quarterly, and yearly. Aggregations include sum, average, median, count, minimum, and maximum.

Rows without a valid selected date are skipped. When a numeric measure is used, rows without a usable measure are also skipped. The plotted series describes the observed periods; it is not a forecast.

## 6. Relationships

The relationship view compares two numeric fields using paired rows with valid values. Pearson’s r is available when there are at least three paired observations and both fields have non-zero variance. The view may rank candidate numeric pairs by absolute correlation.

Pearson correlation describes linear association. It can be affected by outliers and hidden variables, and it does **not** show that one field caused the other. A weak correlation also does not rule out a non-linear relationship.

## 7. Category-specific analysis

The selected category provides suggested column roles and domain-focused summaries for:

- Sales & E-commerce
- Retail & Inventory
- Education & Academics
- Marketing & Campaigns
- Customer Analytics
- Finance & Accounting
- Product & SaaS
- General / Custom Dataset

Role mappings are editable because column names and meanings vary. Domain metrics are calculated only when the needed fields are available and usable. For example, a profit measure needs a profit or suitable cost field; DataPrism should not infer profit from sales alone.

## 8. Chart Builder

The builder supports:

| Chart | Typical use | Data requirement |
| --- | --- | --- |
| Bar | Compare groups | Category/date on X and a numeric measure or row count |
| Line | Show ordered values or time patterns | Numeric/date X and numeric measure or row count |
| Scatter | Inspect paired values | Two numeric fields |
| Donut | Show part-to-whole composition | Non-negative values across a limited set of categories |
| Histogram | Inspect a numeric distribution | One numeric field; bars show frequency |

The builder supports X and Y selections, aggregation, optional grouping where valid, date grain where applicable, and a custom title. Invalid field/type combinations are rejected with an explanation. Chart exports are available as PNG, SVG, and aggregated chart-data CSV.

Charts summarize the selected rows and settings. Check the axis labels, aggregation, date grain, and excluded values before using an image in a report.

## Interpretation checklist

Before sharing an analysis, ask:

1. Is the detected field type and role correct?
2. How many records were omitted for missing or unparseable values?
3. Are duplicates expected in this dataset?
4. Does the aggregation match the question?
5. Are unusual values errors, legitimate extremes, or a different population?
6. Does the chart show association or a time pattern without implying causation?
7. Does the dataset cover the period and population you want to discuss?