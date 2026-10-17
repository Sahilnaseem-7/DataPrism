FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV DATAPRISM_HOST=0.0.0.0 PORT=8780
EXPOSE 8780
CMD ["python", "platform_server.py"]
