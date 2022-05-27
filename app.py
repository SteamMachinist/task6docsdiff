import glob
import os

from flask import Flask, render_template, url_for, request, redirect
import docx2txt as docx
from odf import text, teletype
from odf.opendocument import load
import difflib as dl

from werkzeug.utils import secure_filename

UPLOAD_FOLDER = "docs"
ALLOWED_EXTENSIONS = {'.docx', '.odt'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['ALLOWED_EXTENSIONS'] = ALLOWED_EXTENSIONS

CONTEXT_N = 7

doc1_name, doc2_name, doc1_ext, doc2_ext = '', '', '', ''


def get_document_text(filename: str, extension: str) -> str:
    if extension == ".docx":
        return docx.process(os.path.join(app.config['UPLOAD_FOLDER'], filename))
    else:
        textdoc = load(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        allparas = textdoc.getElementsByType(text.P)
        all_text = ""
        for par in allparas:
            all_text += teletype.extractText(par)
        return allparas


@app.route('/comparison_result')
def comparison_results():
    text1 = get_document_text(doc1_name, doc1_ext).split()
    text2 = get_document_text(doc2_name, doc2_ext).split()
    diffiter = dl.context_diff(text1, text2, doc1_name, doc2_name, n=CONTEXT_N)
    diff = [d.replace(" ", "").replace("\n", "") for d in diffiter]
    return render_template('comparison_results.html', differences=diff)


def clear_previous_docs():
    previous_docs = glob.glob(os.path.join(app.config['UPLOAD_FOLDER'], '*'))
    for doc in previous_docs:
        os.remove(doc)


@app.route('/', methods=['GET', 'POST'])
def upload_documents_index():
    if request.method == 'GET':
        return render_template('upload_files.html', error_message="")

    doc1_file, doc2_file = request.files['doc1_file'], request.files['doc2_file']

    global doc1_name, doc2_name, doc1_ext, doc2_ext
    doc1_name, doc2_name = secure_filename(doc1_file.filename), secure_filename(doc2_file.filename)
    doc1_ext, doc2_ext = os.path.splitext(doc1_name)[1], os.path.splitext(doc2_name)[1]

    if doc1_name == '' or doc2_name == '':
        return render_template('upload_files.html', error_message="No files chosen")

    if doc1_ext not in app.config['ALLOWED_EXTENSIONS'] or doc2_ext not in app.config['ALLOWED_EXTENSIONS']:
        return render_template('upload_files.html', error_message="Choose .docx or .odt files")

    clear_previous_docs()
    doc1_file.save(os.path.join(app.config['UPLOAD_FOLDER'], doc1_name))
    doc2_file.save(os.path.join(app.config['UPLOAD_FOLDER'], doc2_name))
    return redirect(url_for('comparison_results'))


if __name__ == '__main__':
    app.run()
