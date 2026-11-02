FROM python:3.12-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py .

# Render tự đặt biến PORT; Hugging Face dùng 7860
ENV PORT=7860
EXPOSE 7860
# timeout dài để kịp đọc văn bản dài
CMD gunicorn -b 0.0.0.0:$PORT -w 2 --threads 4 --timeout 300 app:app
