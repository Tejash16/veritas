# ---- Base image ----
FROM python:3.11-slim

# ---- Set working directory ----
WORKDIR /app

# ---- Install system dependencies ----
RUN apt-get update && apt-get install -y build-essential libpq-dev && rm -rf /var/lib/apt/lists/*

# ---- Copy files ----
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# ---- Expose port ----
EXPOSE 8008

# ---- Run FastAPI app ----
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8008"]
