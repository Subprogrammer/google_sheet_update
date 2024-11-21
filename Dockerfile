FROM python:3.9-slim

WORKDIR /app

COPY data.xlsx .

COPY . .

RUN pip install -r requirements.txt

CMD ["python", "update_google_sheet.py"]
