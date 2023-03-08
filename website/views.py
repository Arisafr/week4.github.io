from flask import Blueprint, flash, make_response, redirect, render_template, request, url_for, send_file
from werkzeug.utils import secure_filename
import pandas as pd
from flask_wtf import FlaskForm
from wtforms import SubmitField, FileField

views = Blueprint('views',__name__)
filename = None 
# Global variable to store the dataframe
df = pd.DataFrame()


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

class FileUploadForm(FlaskForm):
    file = FileField()
    submit = SubmitField('Upload')

@views.route('/', methods=['GET','POST'])
def index():
    form = FileUploadForm()
    return render_template('home.html',form=form)


@views.route('/contactus')
def contactus():
    return render_template('contactus.html')

@views.route('/tentangprogram')
def tentangprogram():
    return render_template('tentangprogram.html')

@views.route('/index')
def index():
    return render_template('index.html')

