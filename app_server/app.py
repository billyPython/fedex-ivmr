from flask import Flask, request, flash, redirect, send_file
from werkzeug.utils import secure_filename

from config import ALLOWED_EXTENSIONS

app = Flask(__name__)


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']

        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)

        #TODO: List of tasks:
        # 1. Check `ivmr.py` for todos.
        # 2. Get right mime type(and check if this one is right)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            return send_file(file,
                             attachment_filename=filename,
                             mimetype='application/vnd.ms-excel')
    return '''
    <!doctype html>
    <title>Upload new File</title>
    <h1>Upload new File</h1>
    <form method=post enctype=multipart/form-data>
      <p><input type=file name=file>
         <input type=submit value=Upload>
    </form>
    '''