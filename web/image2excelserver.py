from PIL import Image
import xlsxwriter
from datetime import datetime
import sys
import os, shutil
import random
import time

import _thread, threading
import json

from flask import Flask, render_template, request, send_from_directory, send_file, abort, redirect
from werkzeug.utils import secure_filename

app = Flask(__name__)

print('|------------------------------------------------------------------------------------|')
print('| DO NOT USE THIS WEBSERVER IN A PRODUCTION ENVIROMENT. IT IS NOT DESIGNED FOR THAT! |')
print('|                       THERE IS ONLY SOME ERROR HANDLING                            |')
print('|------------------------------------------------------------------------------------|')
#Modes:
#Greyscale (GS)
#RGB
#FILTER

#Convert normal coordinates (1,1) into excel coordiantes A1
def excel_coords(x):
    string = ""
    while x > 0:
        x, remainder = divmod(x - 1, 26)
        string = chr(65 + remainder) + string

    return str(string)

def get_time():
    return str(datetime.now().hour) + str(datetime.now().minute)
#Convert the RGB colour values to hex
def rgb_hex(r, g, b):
    '#%02x%02x%02x' % (r, g, b)

#Generate a loading bar
def gen_bar(percent):
    bar = ''
    for i in range(percent):
        bar += '■'
    for i in range(100-len(bar)):
        bar += '□'

    return bar

converting_threads = {}

#Thread for converting images.
class ConvertingThread(threading.Thread):
    def __init__(self, display_name, image_name, output, mode, filters=None):
        
        #Define all the self variables
        self.progress = 0
        self.finished = False
        self.display_name = display_name
        self.image_name = image_name
        self.output = output
        self.mode = mode
        self.filter = filters
        self.status = 'Converting...'
        
        super().__init__()

    def run(self):

        #Open the image.
        im = Image.open(self.image_name)
        width, height = im.size

        rgb_im = im.convert('RGB')

        #print(r,g,b)

        #Create the spreadsheet
        workbook = xlsxwriter.Workbook('{}.xlsx'.format(self.output))

        worksheet = workbook.add_worksheet("Image")

        i=0

        #Iterate through every pixel row by row.
        while i < height:

            #print("Converting. {}% {}".format(round((i/height)*100)+1, gen_bar(round((i/height)*100)+1)), end="\r")

            #Iterate through each pixel in every column and write their RGB values to seperate lines in a spreadsheet.
            for pixel in range(width):
                r, g, b = rgb_im.getpixel((pixel, i))

                worksheet.write(i+2, pixel, r)
                worksheet.write(i+3, pixel, g)
                worksheet.write(i+4, pixel, b)

            #Get the locations on the spreadsheet needed for conditional formatting
            red_value = 'A' + str(i) + ':' + excel_coords(width+5)+ str(1)
            green_value = 'A' + str(i+1) + ':' + excel_coords(width+5) + str(2)
            blue_value = 'A' + str(i+2) + ':' + excel_coords(width+5) + str(3)

            #Write 0 and 255 at the end of every row so conditional formatting works
            worksheet.write(i, width, 0)
            worksheet.write(i, width+1, 255)

            worksheet.write(i+1, width, 0)
            worksheet.write(i+1, width+1, 255)

            worksheet.write(i+2, width, 0)
            worksheet.write(i+2, width+1, 255)

            #print(red_value)
            #print(green_value)
            #print(blue_value)

            #worksheet.conditional_format(red_value , {'type': '2_color_scale'})
            #worksheet.conditional_format(green_value , {'type': '2_color_scale'})
            #worksheet.conditional_format(blue_value , {'type': '2_color_scale'})

            #RGB self.mode
            if self.mode == 'RGB':
                worksheet.conditional_format(red_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#FF0000"})
                worksheet.conditional_format(green_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#00FF00"})

                worksheet.conditional_format(blue_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#0000FF"})

            #Greyscale self.mode
            if self.mode == 'GS':
                worksheet.conditional_format(red_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#FFFFFF"})
                worksheet.conditional_format(green_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#FFFFFF"})

                worksheet.conditional_format(blue_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#FFFFFF"})
            #self.filter self.mode
            if self.mode == 'FILTER':
                #print(self.filter)
                worksheet.conditional_format(red_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': self.filter})
                worksheet.conditional_format(green_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': self.filter})

                worksheet.conditional_format(blue_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': self.filter})

            i+=3
            #print(i)
            self.progress = round((i/height) * 100) + 1
        print("\nFinishing off...")
        self.status += '\nFinishing off...'
        #Make the cells nice and square
        worksheet.set_column(0, width, 2.14)

        worksheet.write('A1', 'Image produced by the Image to Excel converter. Zoom out to view the full image. Converter made by Oscar Peace')
        worksheet.write('A2', 'Original dimensions: {}px X {}px ({} pixels). Spreadsheet dimensions: {}cells X {}cells. ({} cells)'.format(width,height,width*height, width, height*3, width*(height*3)))
        worksheet.write_url('B1', 'https://github.com/sccreeper/image2excel/', string='View the source code on GitHub')

        #Insert the original image into another sheet
        self.status += '\nAdding orginal image to sheet...'

        worksheet1 = workbook.add_worksheet("Original Image")

        worksheet1.write('A1', "Original image: '{}'".format(os.path.basename(self.display_name)))
        worksheet1.insert_image('A2', self.image_name)

        #Remove unused
        print("Removing unused rows and columns...")
        self.status += '\nRemoving unused rows and columns...'
        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.set_column('{}:XFD'.format(excel_coords(width+3)), None, None, {'hidden': True})

        #Save and close
        print("Saving...")
        self.status += '\nGetting ready for download...'
        workbook.close()
        print("Image saved to {}".format(self.output + ".xlsx"))

        self.finished = True

#Flask web server code
#Check folders to see if age is over ten minutes if so then delete them
def check_folders():
    while True:
        time.sleep(1)

        folders = os.listdir('temp/')

        for i in range(len(folders)):
            print('checking')
            if (time.time() - os.path.getmtime('temp/{}/{}.xlsx'.format(folders[i],folders[i]))) > 600:
                shutil.rmtree('temp/{}'.format(folders[i]))
            else:
                continue 
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
            filename = str(random.randint(0,100000))

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

            converting_threads[str(filename)] = ConvertingThread("test_image.JPG" ,'static/test_image.JPG', 'temp/' + filename + '/' + filename, request.form['filter'], request.form['colour'])
            converting_threads[str(filename)].start()
        else:
            f = request.files['file']
            #Save file with custom name
            f.save('temp/' + filename + '/' + filename)

            converting_threads[str(filename)] = ConvertingThread(secure_filename(f.filename) ,'temp/' + filename + '/' + secure_filename(filename), 'temp/' + filename + '/' + filename, request.form['filter'], request.form['colour'])
            converting_threads[str(filename)].start()

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
            
            return send_file('temp' + '/' + name + '/' + name + '.xlsx',   as_attachment=True, attachment_filename='conversion.xlsx')

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

_thread.start_new_thread(check_folders, ())


if os.path.exists("temp") != True:
    os.mkdir("temp")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=80, threaded=True)
