from flask import Flask, render_template, request, send_from_directory, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
from werkzeug.wrappers.response import Response
from convert import Convert
import json
import os
import shutil


app = Flask(__name__)
CORS(app)


@app.route('/')
def home():
    return render_template('home.html')


@app.route('/api/convert', methods=['POST'])
def convert():
    clearFolder('input/')
    clearFolder('output/')

    try:
        file = request.files.get('file[]')
        file.save(f'input/{file.filename}')

        print(file.filename)
        Convert(file.filename)

        outputFile = os.listdir('output/')

        return send_from_directory('output/', path=outputFile[0], as_attachment=True)

    except Exception as e:
        print(e)
        return Response(json.dumps({'message': 'error'}),
                        status=500, mimetype='application/json')


def clearFolder(path):
    folder = path
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))
