import urllib

from flask import Flask, render_template
from src import processing as ps
import os

app = Flask(__name__)

@app.route("/home")
def index():
    cwd = os.getcwd()
    orgNames = ps.getOrgNames()
    return render_template("index.html", data =orgNames)

@app.route("/data")
def generatefile():
    ps.getOrgAppDetails()
    return "File Successfully generated"

@app.route("/")
def getData():
    ps.getOrgAppDetails()
    return "File Successfully Generated"


@app.route("/test")
def json():
    return render_template('test.html')

#background process happening without any refreshing
@app.route('/background_process_test')
def background_process_test():
    print ("Hello")
    return ("nothing")



if (__name__ == "__main__"):
    app.run(debug=True, port=9090)