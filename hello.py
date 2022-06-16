from flask import Flask, request, jsonify, abort, redirect, url_for, render_template, send_file, flash, redirect, url_for
from bs4 import BeautifulSoup
import requests, statistics
import openpyxl
import numpy as np
import pandas as pd
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
    return redirect(url_for('submit'))

class MyForm(FlaskForm):
    # name = 'name'
    file = FileField(validators=[DataRequired()])

@app.route('/submit', methods=('GET', 'POST'))
def submit():
    form = MyForm()
    
    if form.validate_on_submit():
        f = form.file.data
        filename = 'Аналоги.xlsx'
        f.save(os.path.join(
            filename
        ))
        auto.autoru_appraiser(filename)
        

        return send_file(filename,
                     mimetype='xlsx',
                     attachment_filename=filename,
                     as_attachment=True)
        
    return render_template('submit.html', form=form)
 
if __name__ == "__main__":
    app.run(debug=True)