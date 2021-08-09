#This example shows how to use the workbooksplit functionality found in the VideoConverter class
import sys
sys.path.append("..")
from lib import image2excel as i2e
import _thread as thread

file_path = input("Full file path: ")
output_path = input("Full output path: ")
use_zip = True if input("Use zip (y/n): ").lower() == "y" else False

converter = i2e.VideoConverter(file_path, output_path, scale=0.1, frame_skip=100) #workbooksplit=1, use_zip=use_zip)

#do on different thread
thread.start_new_thread(converter.convert, ())

conversion_status = ""

while not converter.finished:
    if not converter.status == conversion_status:
        conversion_status = converter.status
        print("\n" + converter.status)
    if converter.status.startswith("Converting"):
        print(f"Progess: {converter.progress}%", end="\r")
