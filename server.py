# all the imports
import os
from flask import Flask, request, redirect, url_for, abort, \
    render_template
from werkzeug.utils import secure_filename
import BSLC_attendance

ALLOWED_EXTENSIONS = set(['txt','csv', 'xlsx', 'xls'])

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # max 16 MB
app.config.from_object(__name__)
app.config.from_envvar('FLASKR_SETTINGS', silent=True)

# File uploading methods


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

# Index Initialized

@app.route('/')
def index():
    if request.method == 'POST':
        file = request.files['file']
        sender_address = request.form['sender_address']
        password = request.form['password']
        subject = request.form['subject']
        text_body = request.form['text_body']
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            BSLC_attendance.main(filename, sender_address, password, subject, text_body)  # Send the emails from test file
            return redirect(url_for('index'))
    return render_template('file_upload.html')

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5001, debug=True)
