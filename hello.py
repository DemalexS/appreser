from flask import Flask, request, jsonify, abort, redirect, url_for, render_template, send_file, flash
from bs4 import BeautifulSoup
import requests, statistics
import openpyxl
from flask_wtf import FlaskForm
from wtforms import StringField, FileField
from wtforms.validators import DataRequired
import os
from werkzeug.utils import secure_filename
import auto

app = Flask(__name__)

app.config.update(dict(
    SECRET_KEY="powerful secretkey",
    WTF_CSRF_SECRET_KEY="a csrf secret key"
))


@app.route('/')
def redir_submit():
    return redirect(url_for('form'))

class MyForm(FlaskForm):
    # name = 'name'
    file = FileField(validators=[DataRequired()])
@app.route('/form', methods=['GET', 'POST'])
def form():

    #Причем в начале проверяем наличие авторизации, если флага нет, то кидаем обработку 401 ошибки и не даем работать с прогой
    # if not session.get('logged_in'):
    #     abort(401)

    """
    Тут делаем что-то полезное в случае успешной авторизации
    """

    return render_template('form.html')


@app.route('/submit', methods=('GET', 'POST'))
def submit():
    form = MyForm()
    
    if form.validate_on_submit():
        f = form.file.data
        filename = 'analogi.xlsx'
        f.save(os.path.join(filename))
        auto.autoru_appraiser(filename)
        

        return send_file(filename,
                     mimetype='xlsx',
                     attachment_filename=filename,
                     as_attachment=True)
        
    return render_template('submit.html', form=form)

@app.errorhandler(500)
def page_not_found(e):
    error = 'Произошла ошибка при работе скрипта, вероятно auto.ru опять показывает капчу. Сообщите об этом Алексею прямо сейчас.'
    return render_template('form.html', error=error), 500
 
if __name__ == "__main__":
    app.run(debug=True)