"""
image2excel is software that converts images and video to Excel spreadsheets.
    Copyright (C) 2021  Oscar Peace AKA sccreeper

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

__name__ = "image2excel.lib"

from typing import Union
from PIL import Image
import cv2
import math
import numpy as np
import tempfile
import shutil
import xlsxwriter
from datetime import datetime
import time
from enum import Enum
import zipfile
import os
from distutils.dir_util import copy_tree


#Convert normal coordinates (1,1) into excel coordiantes A1
def to_excel_coords(x):
    string = ""
    while x > 0:
        x, remainder = divmod(x - 1, 26)
        string = chr(65 + remainder) + string

    return str(string)

class Mode(Enum):
    RGB = 1
    GREYSCALE = 2
    FILTER = 3

#The image converter, will convert an image into different coloured cells on a spreadsheet.
#Also works with 2D arrays of RGB data.
#Array format:
#[ [ [r, g, b] ] ]

class ImageConverter:
    def __init__(self, file_name: Union[list, str], output_path: str, mode: int=Mode.RGB, scale: float=1.0, filter_colour: str="#FFFFFF"):
        """The main image class.

        Args:
            file_name (str, list): The path of the image you want to convert or image array data. Array example: [ [ [r, g, b] ] ]
            output_path (str): The output path of the Excel spreadsheet.
            mode ([type], optional): The filter mode. Use the `Mode` enum. Defaults to Mode.RGB.
            scale (int, optional): Scale of the image. For images over 1000px, a scale of 0.25 or lower is recommended. Defaults to 1.
            filter_colour (str, optional): The filter colour. Only required if using `Mode.FILTER`. Defaults to "#FFFFFF".
        """

        if isinstance(file_name, list):
            self.image_array = file_name
            self.is_array = True
        else:
            self.is_array = False
        
        self.mode = mode
        self.scale = float(scale)
        self.filter_colour = filter_colour
        self.output_path = output_path

        self.progress = 0
        self.status = ""
        self.finished = False

        self.temp_dir = ""

    def __get_time(self):
        return str(datetime.now().hour) + str(datetime.now().minute)
    #Convert the RGB colour values to hex
    def __rgb_hex(r, g, b):
        '#%02x%02x%02x' % (r, g, b)
    
    def convert(self):
        """Begins the conversion of the image.
        """
        
        start_time = time.time()

        #Open the image and save if array
        if self.is_array:
            data = np.array(self.image_array).astype('uint8')

            im = Image.fromarray(data, "RGB")

            self.temp_dir = tempfile.mkdtemp()

            self.file_name = f"{self.temp_dir}/array {data.size}.png"
            self.image_name = f"{self.temp_dir}/array {data.size}.png"

            im.save(f"{self.temp_dir}/array {data.size}.png")
        else:
            im = Image.open(self.image_name)

        #Scale
        im = im.resize((round(im.size[0] * self.scale), round(im.size[1] * self.scale)))

        width, height = im.size

        rgb_im = im.convert('RGB')

        #print(r,g,b)

        #Create the spreadsheet
        workbook = xlsxwriter.Workbook('{}.xlsx'.format(self.output_path))

        worksheet = workbook.add_worksheet("Image")

        write_line=0
        pixel_line = 0

        #Iterate through every pixel row by row.
        while pixel_line < height:
            
            self.status = "Converting..."
            
            #Iterate through each pixel in the row and write values to spreadsheet.
            for pixel in range(width):
                r, g, b = rgb_im.getpixel((pixel, pixel_line))

                worksheet.write(write_line+2, pixel, r)
                worksheet.write(write_line+3, pixel, g)
                worksheet.write(write_line+4, pixel, b)

            #Get the locations on the spreadsheet needed for conditional formatting
            red_value = 'A' + str(write_line) + ':' + to_excel_coords(width+5)+ str(1)
            green_value = 'A' + str(write_line+1) + ':' + to_excel_coords(width+5) + str(2)
            blue_value = 'A' + str(write_line+2) + ':' + to_excel_coords(width+5) + str(3)

            #Write 0 and 255 at the end of every row so conditional formatting works
            worksheet.write(write_line, width, 0)
            worksheet.write(write_line, width+1, 255)

            worksheet.write(write_line+1, width, 0)
            worksheet.write(write_line+1, width+1, 255)

            worksheet.write(write_line+2, width, 0)
            worksheet.write(write_line+2, width+1, 255)

            #Apply conditional formatting
            #Lots of inline if statements probably bad, but looks smart            

            worksheet.conditional_format(red_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': ("#FF0000" if self.mode == Mode.RGB else ("#FFFFFF" if self.mode == Mode.GREYSCALE else self.filter_colour))})
            worksheet.conditional_format(green_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': ("#00FF00" if self.mode == Mode.RGB else ("#FFFFFF" if self.mode == Mode.GREYSCALE else self.filter_colour))})

            worksheet.conditional_format(blue_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': ("#0000FF" if self.mode == Mode.RGB else ("#FFFFFF" if self.mode == Mode.GREYSCALE else self.filter_colour))})

            pixel_line+=1
            write_line+=3

            #print(i)

            self.progress = round((pixel_line/height)*100)+1

        self.progress = 100

        self.status = "Finishing off..."
        #Make the cells nice and square
        worksheet.set_column(0, width, 2.14)

        #Insert the original image into another sheet
        worksheet1 = workbook.add_worksheet("Info")

        worksheet1.write('A1', f'Original dimensions: {width}px X {height}px ({width*height} pixels). Spreadsheet dimensions: {width}cells X {height}cells. ({width*height} cells)')

        worksheet1.write_url('A1', 'https://github.com/sccreeper/image2excel/', string='View the source code on GitHub')
        worksheet1.write('A3', "Original image: '{}'".format(self.file_name))
        worksheet1.write('A4', f"Time taken: {time.time()-start_time} second{'' if time.time() == 1 else 's'}")
        worksheet1.insert_image('A5', self.file_name)

        #Remove unused
        self.status  = "Removing unused rows and columns..."
        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.set_column('{}:XFD'.format(to_excel_coords(width+3)), None, None, {'hidden': True})

        #Save and close
        self.status = "Saving..."        
        workbook.close()
        self.status = "Image saved!"

        if self.is_array:
            shutil.rmtree(self.temp_dir)

        self.finished = True


#The video converter, basically a "why not" thing. Each every 25th frame is skipped by default in order to save space. 

class VideoConverter:

        def __init__(self, file_name: str, output_path: str, mode: int=Mode.RGB, scale: float=1, filter_colour: str="#FFFFFF", frame_skip: int=25, force_frame_skip: bool=False, videocut: float=1.0, workbooksplit: int=None, use_zip: bool=True):
            """The main image class.

            Args:
                file_name (str): The path of the image you want to convert
                output_path (str): The output path of the Excel spreadsheet.
                mode ([type], optional): The filter mode. Use the `Mode` enum. Defaults to Mode.RGB.
                scale (int, optional): Scale of the image. For images over 1000px, a scale of 0.25 or lower is recommended. Defaults to 1.
                filter_colour (str, optional): The filter colour. Only required if using `Mode.FILTER`. Defaults to "#FFFFFF".
                frame_skip (int, optional): Frame skip, cannot be lower than 25. Higher values are recommended for videos over 10sec.
                frame_skip_force (bool, optional): Forces frame skip, not recommended.
                videocut (float, optional): Percentage as float to cut the video by.
                workbooksplit (int, optional): Split the video into seperate files, every x frames. <= 10 recommended.
            """

            self.image_name = file_name
            self.file_name = file_name
            self.mode = mode
            self.scale = float(scale)
            self.filter_colour = filter_colour

            if frame_skip < 25 and not force_frame_skip:
                raise ValueError(f"frame_skip ({frame_skip}) cannot be less than 25!")

            #Check whether videocut is valid
            if videocut > 1:
                raise ValueError(f"videocut ({videocut}) cannot be greater than 1!")
            else:
                self.videocut = videocut

            #Check whether the sheetsplit is valid
            if workbooksplit != None and workbooksplit > 1:
                raise ValueError(f"Sheetsplit ({workbooksplit}) must be greater than 1!") 
            elif not isinstance(workbooksplit, int) and workbooksplit != None:
                raise TypeError(f"Sheetsplit type: {type(workbooksplit)} must be an integer!")
            elif workbooksplit == None:
                self.__splitworkbooks = False
                self.workbooksplit = 0
            else:
                self.workbooksplit = workbooksplit
                self.__splitworkbooks = True

            self.use_zip = use_zip
            
            self.frame_skip = frame_skip

            self.output_path = output_path

            self.temp_directory = tempfile.mkdtemp()
            
            self.progress = 0
            self.status = ""
            self.finished = False

        #https://stackoverflow.com/a/63763138
        def __rescale_frame(self, frame_input, percent=100):
                width = int(frame_input.shape[1] * percent / 100)
                height = int(frame_input.shape[0] * percent / 100)
                dim = (width, height)
                return cv2.resize(frame_input, dim, interpolation=cv2.INTER_AREA)

        #https://stackoverflow.com/a/1855118
        def __zipdir(self, path, ziph):
        # ziph is zipfile handle
            for root, dirs, files in os.walk(path):
                for file in files:
                    ziph.write(os.path.join(root, file), 
                            os.path.relpath(os.path.join(root, file), 
                                            os.path.join(path, '..')))


        def convert(self):
            """Begins conversion of the video
            """

            start_time = time.time()

            video = cv2.VideoCapture(self.file_name)

            frame_count = math.floor(self.videocut * video.get(cv2.CAP_PROP_FRAME_COUNT))
            print(frame_count)
            print(frame_count * self.videocut)

            width = math.floor(video.get(cv2.CAP_PROP_FRAME_WIDTH))
            height = math.floor(video.get(cv2.CAP_PROP_FRAME_HEIGHT))

            width = math.floor(width * self.scale)
            height = math.floor(height * self.scale)

            framerate = video.get(cv2.CAP_PROP_FPS)
            duration = frame_count * framerate

            #If there was no workbooksplit and we just want it to be standard
            #workbook_total = the total amount of files to be made
            if not self.__splitworkbooks:
                workbook_total = 1
                workbook_count = 1
            else:
                workbook_total = math.floor((frame_count / self.frame_skip) / self.workbooksplit)
                workbook_count = 1

                print(workbook_total)
                print(workbook_count)
                frame_iterator = 0

            workbook_files = []

            #Current frame's index in frame_worksheets
            frame_list_index = 0
            #Actual location of frame in video.
            frame_video_index = 0

            frame_worksheets = []

            for file in range(workbook_total):

                new = False

                if not self.__splitworkbooks:
                    workbook_files.append(xlsxwriter.Workbook('{}.xlsx'.format(self.output_path)))
                    new = True
                else:
                    if frame_iterator >= self.workbooksplit:
                        frame_iterator = 0
                        workbook_count += 1

                        self.status = f"Saving {file}/{workbook_total}"

                        workbook_files[file-1].close()
                        workbook_files.append(xlsxwriter.Workbook(f"{self.temp_directory}/{workbook_count} out of {workbook_total}.xlsx"))

                        new = True
                    else:
                        #Assume start
                        frame_iterator = 0
                        workbook_files.append(xlsxwriter.Workbook(f"{self.temp_directory}/{workbook_count} out of {workbook_total}.xlsx"))
                        new = True

                if new:
                    #Begin by adding info worksheets
                    info_worksheet = workbook_files[file].add_worksheet("Info")

                    f_bold_text = workbook_files[file].add_format({"bold": True})

                    info_worksheet.write('A1', 'Image produced by the image2excel converter. Zoom out to view the full image. Converter made by sccreeper')
                    info_worksheet.write('A2', 'Original dimensions: {}px X {}px ({} pixels). Spreadsheet dimensions: {}cells X {}cells. ({} cells)'.format(width,height,width*height, width, height*3, width*(height*3)))
                    info_worksheet.write('A3', f"Time taken: {time.time()-start_time} second{'' if time.time() == 1 else 's'}")
                    info_worksheet.write('A4', f'Video length: {duration} seconds')
                    info_worksheet.write('A5', f'Original framecount: {video.get(cv2.CAP_PROP_FRAME_COUNT)} frames')
                    info_worksheet.write('A6', f'Shortened framecount: {round(frame_count/self.frame_skip)}')
                    info_worksheet.write('A7', f'Workbook {workbook_count}/{workbook_total}')
                    info_worksheet.write_url('A9', 'https://github.com/sccreeper/image2excel/', string='View the source code on GitHub', cell_format=f_bold_text)

                while frame_video_index < math.floor(frame_count):

                    frame_worksheets.append(workbook_files[file].add_worksheet(f"Frame {frame_video_index}"))

                    #print(frame_worksheets)
                    #print(frame_list_index)
                    #print(width)
                    #print(height)

                    video.set(1, frame_video_index+1)

                    res, frame_data = video.read()

                    frame_data = self.__rescale_frame(frame_data, percent=self.scale*100)
                    
                    pixel_line = 0
                    write_line = 0
                    
                    self.status = f"Converting frame {frame_list_index+1}/{math.floor(frame_count/self.frame_skip)+1}..."
                    self.progress = round((frame_list_index/(frame_count/self.frame_skip))*100)
                    
                    while pixel_line < height:
                        
                        #Iterate through each pixel in every column and write their RGB values to seperate lines in a spreadsheet.
                        for pixel in range(width):
                            r = frame_data[pixel, pixel_line, 2]
                            g = frame_data[pixel, pixel_line, 1]
                            b = frame_data[pixel, pixel_line, 0]
                            
                            frame_worksheets[frame_list_index].write(write_line+2, pixel, r)
                            frame_worksheets[frame_list_index].write(write_line+3, pixel, g)
                            frame_worksheets[frame_list_index].write(write_line+4, pixel, b)                

                        #Get the locations on the spreadsheet needed for conditional formatting
                        red_value = 'A' + str(write_line) + ':' + to_excel_coords(width+5)+ str(1)
                        green_value = 'A' + str(write_line+1) + ':' + to_excel_coords(width+5) + str(2)
                        blue_value = 'A' + str(write_line+2) + ':' + to_excel_coords(width+5) + str(3)

                        #Write 0 and 255 at the end of every row so conditional formatting works
                        frame_worksheets[frame_list_index].write(write_line, width, 0)
                        frame_worksheets[frame_list_index].write(write_line, width+1, 255)

                        frame_worksheets[frame_list_index].write(write_line+1, width, 0)
                        frame_worksheets[frame_list_index].write(write_line+1, width+1, 255)

                        frame_worksheets[frame_list_index].write(write_line+2, width, 0)
                        frame_worksheets[frame_list_index].write(write_line+2, width+1, 255)
                        
                        #Apply conditional formatting
                        #Lots of inline if statements probably bad, but looks smart            

                        frame_worksheets[frame_list_index].conditional_format(red_value, {'type': '2_color_scale',
                                                                'min_color': "#000000",
                                                                'max_color': ("#FF0000" if self.mode == Mode.RGB else ("#FFFFFF" if self.mode == Mode.GREYSCALE else self.filter_colour))})
                        frame_worksheets[frame_list_index].conditional_format(green_value, {'type': '2_color_scale',
                                                                'min_color': "#000000",
                                                                'max_color': ("#00FF00" if self.mode == Mode.RGB else ("#FFFFFF" if self.mode == Mode.GREYSCALE else self.filter_colour))})

                        frame_worksheets[frame_list_index].conditional_format(blue_value, {'type': '2_color_scale',
                                                                'min_color': "#000000",
                                                                'max_color': ("#0000FF" if self.mode == Mode.RGB else ("#FFFFFF" if self.mode == Mode.GREYSCALE else self.filter_colour))})
                        
                        #Increment values and move to next line
                        write_line += 3
                        pixel_line += 1

                    #Make columns square
                    frame_worksheets[frame_list_index].set_column(0, width, 2.14)
                    
                    #Remove unused rows and columns
                    frame_worksheets[frame_list_index].set_default_row(hide_unused_rows=True)
                    frame_worksheets[frame_list_index].set_column('{}:XFD'.format(to_excel_coords(width+3)), None, None, {'hidden': True})
                    
                    frame_video_index += self.frame_skip
                    frame_list_index += 1

                    if self.__splitworkbooks:
                        frame_iterator += 1
                    if frame_iterator >= self.workbooksplit:
                        break                    

            if not self.__splitworkbooks:    
                #Save worksheet
                self.status = f"Saving {file+1}/{workbook_total}"
                workbook_files[file].close()

            
            if self.__splitworkbooks:
            #Save all files into a zip if use_zip is true, if not copy to output path
                if self.use_zip:
                    zipf = zipfile.ZipFile(f"{self.output_path}.zip", "w", zipfile.ZIP_DEFLATED)
                    self.__zipdir(self.temp_directory, zipf)
                    zipf.close()
                else:
                    copy_tree(self.temp_directory, self.output_path)
                    
            #When finished converting, save the workbook
            self.progress = 100
            self.status = "Saving..."
            self.status = "Video saved!"

            self.finished = True
