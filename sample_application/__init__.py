from flask import Flask, render_template, flash, session, redirect, url_for
from flask_bootstrap import Bootstrap
from flask_appconfig import AppConfig
from flask_wtf import Form, RecaptchaField
from flask_wtf.file import FileField
from wtforms import TextField, HiddenField, ValidationError, RadioField,\
    BooleanField, SubmitField, IntegerField, FormField, validators
from wtforms.validators import Required
import xlrd
import csv
from os import sys
import os.path
import pandas as pd

XLS_FILE = 'Report.xls'

# straight from the wtforms docs:


def csv_from_excel(excel_file):
 #import ipdb;ipdb.set_trace()
# workbook = xlrd.open_workbook(XLS_FILE)
 xl = pd.ExcelFile(XLS_FILE)
# all_worksheets = workbook.sheet_names()
 for sheet in xl.sheet_names:
     file_name = sheet + '.csv'
     df = pd.read_excel(XLS_FILE, sheet_name=sheet)
     df.to_csv(file_name, index=False)

def create_app(configfile=None):
    app = Flask(__name__)
    #app.wsgi_app = ProxyFix(app.wsgi_app)
    AppConfig(app, configfile)  # Flask-Appconfig is not necessary, but
                                # highly recommend =)
                                # https://github.com/mbr/flask-appconfig
    Bootstrap(app)
    # in a real app, these should be configured through Flask-Appconfig


    @app.errorhandler(404)
    def page_not_found(e):
        return render_template('404.html'), 404


    @app.errorhandler(500)
    def internal_server_error(e):
        return render_template('500.html'), 500


    @app.route('/report_index/<key>', methods=['GET', 'POST'] )
    def report_page(key):
      print ("Reached in {} Report Page\n", key)
      DATA_FILE   = str(key)+'.csv'
      table = []
      # Open the csv file and load it up with TableFu
      #table = table_fu.TableFu(open(DATA_FILE, 'U'))
      with open(DATA_FILE, 'r') as csvfile:
          csvreader = csv.reader(csvfile)
          fields = next(csvreader)
          for row in csvreader:
              table.append(row)
      #Get the template and render it to a string.
      #Pass table in as a var called table.
      return  render_template('report_page.html',table=table)

    @app.route('/report', methods=['GET', 'POST'] )
    def message():
      print ("Reached Here\n")
      if (os.path.isfile(XLS_FILE)):
       csv_from_excel(XLS_FILE)
       workbook = xlrd.open_workbook(XLS_FILE)
       all_worksheets = workbook.sheet_names()
       sheet = {}
       for worksheet_name in all_worksheets:
        sheet[worksheet_name] = str(worksheet_name)
       print ("\n Reach to Sheet name{}", worksheet_name)
       return render_template('report.html',sheet=sheet )
      else:
       return render_template('report_error.html' )

    @app.route('/generate', methods=['GET', 'POST'] )
    def generate():
      print ("Generating Report  Here ..Wait for a sec...\n")
      return redirect(url_for('index'))

    @app.route('/', methods=['GET', 'POST'])
    def index():
        form = {}
        return render_template('index.html', form=form)

    return app

if __name__ == '__main__':
    create_app().run( debug=True)
