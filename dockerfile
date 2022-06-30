FROM python:3.9

COPY . /root

WORKDIR /root

RUN pip install flask gunicorn flask_wtf requests openpyxl bs4 fake_useragent
