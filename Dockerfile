FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir flask==3.0.0 groq==0.9.0 httpx==0.27.0 pypdf2==3.0.1 python-docx==1.1.0

COPY app.py .
COPY templates/ templates/

EXPOSE 5000

CMD ["python", "app.py"]
