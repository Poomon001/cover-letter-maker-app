import docx
import os
from flask import Flask, render_template, url_for, redirect, request, session

app = Flask(__name__)

''' save path to this app.py file '''
APP_ROOT = os.path.dirname(os.path.abspath(__file__))

@app.route("/")
def home():
    return render_template("upload.html", text="Please upload")

@app.route("/", methods=['GET', 'POST'])
def upload():
    if request.method == "POST":
        target = os.path.join(APP_ROOT)

        ''' get file from the submitted form '''
        file = request.files.get("file")

        ''' get file name '''
        filename = file.filename

        ''' path + file name = saving location '''
        destination = "./".join([target, filename])
        file.save(destination)
        make()
        return render_template("upload.html", text=filename + " is saved")
    else:
        return render_template("upload.html", text="Upload docx file")

def make():
    print("hello")
    model = ""
    doc = docx.Document("model.docx")
    file = doc.paragraphs

    ''' create style obj (access by document.styles) '''
    style = doc.styles['Normal']

    ''' modify font (style.class.method) '''
    font = style.font
    font.name = 'Arial'
    font.size = docx.shared.Pt(18)

    for line in (file):
        print(line.text)
        model += line.text
        model += ("\n")

    doc.save("output.docx")

if (__name__) == "__main__":
    app.run(debug=True)