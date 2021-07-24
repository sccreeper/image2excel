#__package__ = "image2excel.web"

import sys

sys.path.append("..")

from lib import image2excel as i2e
import os
from random import randint, choice
import time
import magic

import _thread as thread

import json

from flask import Flask, render_template, request, send_file, abort

app = Flask(__name__)

if len(sys.argv) > 1:
    port = sys.argv[1]
else:
    port = 5000

converting_threads = {}
test_image_names = ["../assets/test_image_2.JPG", "../assets/test_image_3.JPG", "../assets/test_image.JPG"]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', defaults={'name':None, 'pid':None}, methods=['POST', 'GET'])
@app.route('/convert/<name>', defaults={'pid':None})
@app.route('/convert/progress/<pid>', defaults={'name':"progress"})
def convert(name,pid):
    global converting_threads

    if request.method == 'POST':

        if name != None:
            return abort(402)

        #choose a name for the file
        while True:
            filename = str(randint(0,100000))

            if os.path.exists(filename):
                continue
            else:
                break


        os.mkdir('temp/' + filename)

        #If the file form is empty use the default test image.
        #print(request.files['file'])

        if request.files['file'].filename == '':
            print("File input empty! Using test image instead.")

            time.sleep(0.5)

            converting_threads[str(filename)] = i2e.ImageConverter("../assets/test_image.JPG", 'temp/' + filename + '/' + filename, i2e.Mode[request.form["filter"]], request.form["scale"], request.form["filter_colour"])
            thread.start_new_thread(converting_threads[str(filename)].convert, ())
        else:
            f = request.files['file']

            temp_file_name = 'temp/' + filename + '/' + filename

            #Save file with custom name
            f.save(temp_file_name)

            if magic.from_file(temp_file_name, mime=True).startswith("image") and not magic.from_file(temp_file_name, mime=True) == "image/gif":
                #Image

                converting_threads[str(filename)] = i2e.ImageConverter(temp_file_name, temp_file_name, i2e.Mode[request.form["filter"]], request.form["scale"], request.form["filter_colour"])
                thread.start_new_thread(converting_threads[str(filename)].convert, ())
            else:
                #Video

                converting_threads[str(filename)] = i2e.VideoConverter(temp_file_name, temp_file_name, i2e.Mode[request.form["filter"]], request.form["scale"], request.form["filter_colour"], int(request.form["frameskip"]), videocut=int(request.form["videocut"])/100)
                thread.start_new_thread(converting_threads[str(filename)].convert, ())

        return str(filename)
    else:
        if name == 'progress':
            progress_report = {}

            progress_report['progress'] = converting_threads[pid].progress
            progress_report['finished'] = str(converting_threads[pid].finished)
            progress_report['status'] = converting_threads[pid].status

            return json.dumps(progress_report)
        else:
            #serve the file
            print('temp/' + name + '/' + name + '.xlsx')

            return send_file('temp' + '/' + name + '/' + name + '.xlsx',   as_attachment=True, download_name=f'Conversion {name}.xlsx')

#Stop browser caching
@app.after_request
def add_header(r):
    """
    Add headers to both force latest IE rendering engine or Chrome Frame,
    and also to cache the rendered page for 10 minutes.
    """
    r.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    r.headers["Pragma"] = "no-cache"
    r.headers["Expires"] = "0"
    r.headers['Cache-Control'] = 'public, max-age=0'
    return r


if os.path.exists("temp") != True:
    os.mkdir("temp")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=port, threaded=True)