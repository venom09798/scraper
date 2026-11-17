FROM mcr.microsoft.com/playwright/python:v1.65.2-noble

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Playwright ke liye missing system dependencies install karne ke liye
RUN playwright install-deps chromium

COPY . .

EXPOSE 5000

CMD ["python", "leads_scraper.py"]
