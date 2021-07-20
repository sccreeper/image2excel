__name__ = "image2excel"

from lib import image2excel
import sys
import _thread as thread
from random import choice

args = sys.argv

#Generate a loading bar
def gen_bar(percent):
    bar = ''
    for i in range(percent):
        bar += '■'
    for i in range(100-len(bar)):
        bar += '□'

    return bar

#See if help has been selected
if args[1] == 'help':
    print(f"""
Image to Excel converter. Written by Oscar Peace AKA sccreeper.

Command syntax:
    python {args[0]} <media type> <filename> <file output> <scale> <mode>
Example usage:
    python {args[0]} image test_image test_output 1 RGB
    python {args[0]} video video.mp4 video_output 1 RGB
    python {args[0]} image test_image test_output 1 FILTER #FF00FF

    <filename> can be 'test_image' if you just want to test the program.
Modes:
    GREYSCALE
    RGB
    FILTER <HEX COLOUR>


Scale: controls the scale of the image. For large images (over 1000px) a smaller scale such as 0.1 or 0.05 is recommended.
Media type: The type of media. video, image.
""")
    input('Press enter to continue...')
    exit()

test_images = ['assets/test_image.JPG', 'assets/test_image_2.JPG', 'assets/test_image_3.JPG']

if args[2] == 'test_image':
    image_name = choice(test_images)
else:
    image_name = args[2]

#Set the variables fot cmd args
output_file = args[3]
scale = args[4]

if args[5] == "FILTER":
    filter = image2excel.Mode.FILTER
    filter_colour = args[6]
else:
    filter = image2excel.Mode[args[5]]

if sys.argv[1] == "image":
    try:
        converter = image2excel.ImageConverter(image_name, output_file, filter, scale, filter_colour)
    except NameError:
        converter = image2excel.ImageConverter(image_name, output_file, filter, scale)
else:
    try:
        converter = image2excel.VideoConverter(image_name, output_file, filter, scale, filter_colour)
    except NameError:
        converter = image2excel.VideoConverter(image_name, output_file, filter, scale)

conversion_status = ""

#print("Converting...")

#do on different thread
thread.start_new_thread(converter.convert, ())

while not converter.finished:
    if not converter.status == conversion_status:
        conversion_status = converter.status
        print(converter.status + "\n")
    if converter.status == "Converting...":
        print(f"Progess: {converter.progress}% {gen_bar(converter.progress)}", end="\r")
