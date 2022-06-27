FROM python:3.7-slim

COPY . /root

WORKDIR /root

RUN pip install flask gunicorn flask_wtf requests openpyxl bs4 undetected-chromedriver
