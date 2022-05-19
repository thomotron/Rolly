FROM python:alpine

WORKDIR /app/
COPY requirements.txt ./
RUN pip install -r requirements.txt

COPY rolly.py ./

ENTRYPOINT ["python3", "rolly.py"]
