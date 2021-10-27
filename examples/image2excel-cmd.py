#Demonstrates basic functionality of the lib
__name__ = "image2excel"

import sys
sys.path.append("..")

from lib import image2excel
import _thread as thread
from random import choice
import argparse
import multiprocessing as mp

#Generate a loading bar
def gen_bar(percent):
    bar = ''
    for i in range(percent):
        bar += '■'
    for i in range(100-len(bar)):
        bar += '□'

    return bar

parser = argparse.ArgumentParser(description='Takes a image/video and converts them into Excel spreadsheets. Made by sccreeper.')

parser.add_argument("type", help="The type of media. image or video", type=str)
parser.add_argument("file_path", help="Input file path", type=str)
parser.add_argument("output_path", help="Output path", type=str)
parser.add_argument("--mode", type=str, help="Mode: GREYSCALE, RGB (default), FILTER", default="RGB")
parser.add_argument("--scale", type=float, help="Scale of the media", default=0.5)
parser.add_argument("--filter", type=str, help="Hex colour of the filter")
parser.add_argument("--frameskip", type=int, help="The amount of frames to skip (gap between each spreadsheet), defaults to 25.", default=25)
parser.add_argument("--forceframeskip", action="store_true", help="Force the frame skip")
parser.add_argument("--videocut", type=float, help="Cuts the video down by a certain percentage. E.G 0.5 would be half.", default=0.5)
parser.add_argument("--workbooksplit", type=int, help="Split the video into seperate files, every x frames. <= 10 recommended.", default=None)
parser.add_argument("--threads", type=int, help="How many threads you want to use for video conversion. 0 will disable multithreading.", default=1)

args = parser.parse_args()


test_images = ['assets/test_image.JPG', 'assets/test_image_2.JPG', 'assets/test_image_3.JPG']

#print(args.scale)
#print(args.file_path)

if args.file_path == 'test_image':
    image_name = choice(test_images)
else:
    image_name = args.file_path

#Check whether threadcount is valid
if args.threads > mp.cpu_count():
    print(f"Maximum number of threads allowed is {mp.cpu_count()}!")
    exit()
elif args.threads == 0:
    use_multithreading = False
else:
    use_multithreading = True
    print(args.threads)

#Set the variables for cmd args
output_file = args.output_path
scale = args.scale

if args.mode == "FILTER":
    filter = image2excel.Mode.FILTER
    filter_colour = args.filter
else:
    filter = image2excel.Mode[args.mode]

if args.type == "image" and not ".gif" in args.file_path:
    try:
        converter = image2excel.ImageConverter(image_name, output_file, filter, scale, filter_colour)
    except NameError:
        converter = image2excel.ImageConverter(image_name, output_file, filter, scale)
else:
    try:
        converter = image2excel.VideoConverter(image_name, output_file, filter, scale, filter_colour, frame_skip=args.frameskip, force_frame_skip=args.forceframeskip, videocut=args.videocut, workbooksplit=args.workbooksplit, threadcount=args.threads, use_multithreading=True)
    except NameError:
        converter = image2excel.VideoConverter(image_name, output_file, filter, scale, frame_skip=args.frameskip, force_frame_skip=args.forceframeskip, videocut=args.videocut, workbooksplit=args.workbooksplit)

conversion_status = ""

#print("Converting...")

#do on different thread
thread.start_new_thread(converter.convert, ())

while not converter.finished:
    if not converter.status == conversion_status:
        conversion_status = converter.status
        print("\n" + converter.status)
    if converter.status.startswith("Converting"):
        print(f"Progess: {converter.progress}% {gen_bar(converter.progress)}", end="\r")
