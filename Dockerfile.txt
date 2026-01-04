FROM python:3.11-slim

WORKDIR /app

COPY . .

RUN apt-get update && apt-get install -y \
    build-essential \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --upgrade pip \
    && pip install -r requirements.txt

RUN mkdir -p uploads output

EXPOSE 10000

CMD ["gunicorn", "viewer_app:app", "--bind", "0.0.0.0:10000"]
