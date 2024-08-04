FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt requirements.txt
RUN apt-get update && apt-get install -y \
    build-essential \
    gfortran \
    && rm -rf /var/lib/apt/lists/* \
    && pip install --no-cache-dir -r requirements.txt

COPY . .

CMD ["python", "app.py"]
