"""
image2excel is a program that converts images and video to Excel spreadsheets.
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

from PIL import Image
import cv2
import math
import xlsxwriter
from datetime import datetime
from enum import Enum


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
class ImageConverter:
    def __init__(self, file_name: str, output_path: str, mode=Mode.RGB, scale=1, filter_colour="#FFFFFF"):
        """The main image class.

        Args:
            file_name (str): The path of the image you want to convert
            output_path (str): The output path of the Excel spreadsheet.
            mode ([type], optional): The filter mode. Use the `Mode` enum. Defaults to Mode.RGB.
            scale (int, optional): Scale of the image. For images over 1000px, a scale of 0.25 or lower is recommended. Defaults to 1.
            filter_colour (str, optional): The filter colour. Only required if using `Mode.FILTER`. Defaults to "#FFFFFF".
        """

        self.image_name = file_name
        self.file_name = file_name
        self.mode = mode
        self.scale = float(scale)
        self.filter_colour = filter_colour
        self.output_path = output_path

        self.progress = 0
        self.status = ""
        self.finished = False

    def __get_time(self):
        return str(datetime.now().hour) + str(datetime.now().minute)
    #Convert the RGB colour values to hex
    def __rgb_hex(r, g, b):
        '#%02x%02x%02x' % (r, g, b)
    
    def convert(self):
        """Begins the conversion of the image.
        """

        #Open the image.
        im = Image.open(self.image_name)

        width, height = im.size

        #Scale
        im = im.resize((round(width * self.scale), round(height * self.scale)))

        width, height = im.size

        rgb_im = im.convert('RGB')

        #print(r,g,b)

        #Create the spreadsheet
        workbook = xlsxwriter.Workbook('{}.xlsx'.format(self.output_path))

        worksheet = workbook.add_worksheet("Image")

        i=0

        #Iterate through every pixel row by row.
        while i < height:
            
            self.status = "Converting..."
            
            #Iterate through each pixel in the row and write values to spreadsheet.
            for pixel in range(width):
                r, g, b = rgb_im.getpixel((pixel, i))

                worksheet.write(i+2, pixel, r)
                worksheet.write(i+3, pixel, g)
                worksheet.write(i+4, pixel, b)

            #Get the locations on the spreadsheet needed for conditional formatting
            red_value = 'A' + str(i) + ':' + to_excel_coords(width+5)+ str(1)
            green_value = 'A' + str(i+1) + ':' + to_excel_coords(width+5) + str(2)
            blue_value = 'A' + str(i+2) + ':' + to_excel_coords(width+5) + str(3)

            #Write 0 and 255 at the end of every row so conditional formatting works
            worksheet.write(i, width, 0)
            worksheet.write(i, width+1, 255)

            worksheet.write(i+1, width, 0)
            worksheet.write(i+1, width+1, 255)

            worksheet.write(i+2, width, 0)
            worksheet.write(i+2, width+1, 255)

            #Apply conditional formatting
            
            #RGB Mode
            if self.mode == Mode.RGB:
                worksheet.conditional_format(red_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#FF0000"})
                worksheet.conditional_format(green_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#00FF00"})

                worksheet.conditional_format(blue_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#0000FF"})
            
            #Greyscale mode                                       
            if self.mode == Mode.GREYSCALE:
                worksheet.conditional_format(red_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#FFFFFF"})
                worksheet.conditional_format(green_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#FFFFFF"})

                worksheet.conditional_format(blue_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': "#FFFFFF"})
            #Filter mode
            if self.mode == Mode.FILTER:
                worksheet.conditional_format(red_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': self.filter_colour})
                worksheet.conditional_format(green_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': self.filter_colour})

                worksheet.conditional_format(blue_value, {'type': '2_color_scale',
                                                    'min_color': "#000000",
                                                    'max_color': self.filter_colour})

            i+=3 
            #print(i)

            self.progress = round((i/height)*100)+1

        self.progress = 100

        self.status = "Finishing off..."
        #Make the cells nice and square
        worksheet.set_column(0, width, 2.14)

        worksheet.write('A1', 'Image produced by the image2excel converter. Zoom out to view the full image. Converter made by sccreeper')
        worksheet.write('A2', 'Original dimensions: {}px X {}px ({} pixels). Spreadsheet dimensions: {}cells X {}cells. ({} cells)'.format(width,height,width*height, width, height*3, width*(height*3)))
        worksheet.write_url('B1', 'https://github.com/sccreeper/image2excel/', string='View the source code on GitHub')

        #Insert the original image into another sheet
        worksheet1 = workbook.add_worksheet("Original Image")

        worksheet1.write_url('A1', 'https://github.com/sccreeper/image2excel/', string='View the source code on GitHub')
        worksheet1.write('A2', "Original image: '{}'".format(self.file_name))
        worksheet1.insert_image('A3', self.file_name)

        #Remove unused
        self.status  = "Removing unused rows and columns..."
        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.set_column('{}:XFD'.format(to_excel_coords(width+3)), None, None, {'hidden': True})

        #Save and close
        self.status = "Saving..."        
        workbook.close()
        self.status = "Image saved!"

        self.finished = True


#The video converter, basically a "why not" thing. Each every 25th frame is skipped by default in order to save space. 

class VideoConverter:

        def __init__(self, file_name: str, output_path: str, mode=Mode.RGB, scale=1, filter_colour="#FFFFFF", frame_skip=25, force_frame_skip=False):
            """The main image class.

            Args:
                file_name (str): The path of the image you want to convert
                output_path (str): The output path of the Excel spreadsheet.
                mode ([type], optional): The filter mode. Use the `Mode` enum. Defaults to Mode.RGB.
                scale (int, optional): Scale of the image. For images over 1000px, a scale of 0.25 or lower is recommended. Defaults to 1.
                filter_colour (str, optional): The filter colour. Only required if using `Mode.FILTER`. Defaults to "#FFFFFF".
                frame_skip (int, optional): Frame skip, cannot be lower than 25. Higher values are recommended for videos over 10sec.
                frame_skip_force (bool, optional): Forces frame skip, not recommended.
            """

            self.image_name = file_name
            self.file_name = file_name
            self.mode = mode
            self.scale = float(scale)
            self.filter_colour = filter_colour

            if frame_skip < 25 and not force_frame_skip:
                raise ValueError(f"frame_skip ({frame_skip}) cannot be less than 25!")

            self.frame_skip = frame_skip

            self.output_path = output_path

            self.progress = 0
            self.status = ""
            self.finished = False

        
        def convert(self):
            """Begins conversion of the video
            """

            video = cv2.VideoCapture(self.file_name)

            frame_count = video.get(cv2.CAP_PROP_FRAME_COUNT)
            width = math.floor(video.get(cv2.CAP_PROP_FRAME_WIDTH))
            height = math.floor(video.get(cv2.CAP_PROP_FRAME_HEIGHT))
            framerate = video.get(cv2.CAP_PROP_FPS)
            duration = frame_count * framerate

            workbook = xlsxwriter.Workbook('{}.xlsx'.format(self.output_path))

            #Begin by adding info worksheets
            info_worksheet = workbook.add_worksheet("Info")

            info_worksheet.write('A1', 'Image produced by the image2excel converter. Zoom out to view the full image. Converter made by sccreeper')
            info_worksheet.write('A2', 'Original dimensions: {}px X {}px ({} pixels). Spreadsheet dimensions: {}cells X {}cells. ({} cells)'.format(width,height,width*height, width, height*3, width*(height*3)))
            info_worksheet.write('A3', f'Video length: {duration} seconds', )
            info_worksheet.write_url('A3', 'https://github.com/sccreeper/image2excel/', string='View the source code on GitHub')

            #Current frame's index in frame_worksheets
            frame_list_index = 0
            #Actual location of frame in video.
            frame_video_index = 0

            frame_worksheets = []

            while frame_video_index < math.floor(frame_count):

                frame_worksheets.append(workbook.add_worksheet(f"Frame {frame_video_index}"))

                #print(frame_worksheets)
                #print(frame_list_index)
                #print(width)
                #print(height)

                video.set(1, frame_video_index+1)

                res, frame_data = video.read()

                line = 0
                
                self.status = f"Converting frame {frame_list_index}/{math.floor(frame_count/self.frame_skip)}..."
                self.progress = round((frame_list_index/(frame_count/self.frame_skip))*100)
                
                while line < height:
                    
                    #Iterate through each pixel in every column and write their RGB values to seperate lines in a spreadsheet.
                    for pixel in range(width):
                        r = frame_data[pixel, line, 2]
                        g = frame_data[pixel, line, 1]
                        b = frame_data[pixel, line, 0]
                        
                        frame_worksheets[frame_list_index].write(line+1, pixel, r)
                        frame_worksheets[frame_list_index].write(line+2, pixel, g)
                        frame_worksheets[frame_list_index].write(line+3, pixel, b)                

                    #Get the locations on the spreadsheet needed for conditional formatting
                    red_value = 'A' + str(line) + ':' + to_excel_coords(width+5)+ str(1)
                    green_value = 'A' + str(line+1) + ':' + to_excel_coords(width+5) + str(2)
                    blue_value = 'A' + str(line+2) + ':' + to_excel_coords(width+5) + str(3)

                    #Write 0 and 255 at the end of every row so conditional formatting works
                    frame_worksheets[frame_list_index].write(line, width, 0)
                    frame_worksheets[frame_list_index].write(line, width+1, 255)

                    frame_worksheets[frame_list_index].write(line+1, width, 0)
                    frame_worksheets[frame_list_index].write(line+1, width+1, 255)

                    frame_worksheets[frame_list_index].write(line+2, width, 0)
                    frame_worksheets[frame_list_index].write(line+2, width+1, 255)

                    #Apply conditional formatting
                    
                    #RGB Mode
                    if self.mode == Mode.RGB:
                        frame_worksheets[frame_list_index].conditional_format(red_value, {'type': '2_color_scale',
                                                            'min_color': "#000000",
                                                            'max_color': "#FF0000"})
                        frame_worksheets[frame_list_index].conditional_format(green_value, {'type': '2_color_scale',
                                                            'min_color': "#000000",
                                                            'max_color': "#00FF00"})

                        frame_worksheets[frame_list_index].conditional_format(blue_value, {'type': '2_color_scale',
                                                            'min_color': "#000000",
                                                            'max_color': "#0000FF"})
                    
                    #Greyscale mode                                       
                    if self.mode == Mode.GREYSCALE:
                        frame_worksheets[frame_list_index].conditional_format(red_value, {'type': '2_color_scale',
                                                            'min_color': "#000000",
                                                            'max_color': "#FFFFFF"})
                        frame_worksheets[frame_list_index].conditional_format(green_value, {'type': '2_color_scale',
                                                            'min_color': "#000000",
                                                            'max_color': "#FFFFFF"})

                        frame_worksheets[frame_list_index].conditional_format(blue_value, {'type': '2_color_scale',
                                                            'min_color': "#000000",
                                                            'max_color': "#FFFFFF"})
                    #Filter mode
                    if self.mode == Mode.FILTER:
                        frame_worksheets[frame_list_index].conditional_format(red_value, {'type': '2_color_scale',
                                                            'min_color': "#000000",
                                                            'max_color': self.filter_colour})
                        frame_worksheets[frame_list_index].conditional_format(green_value, {'type': '2_color_scale',
                                                            'min_color': "#000000",
                                                            'max_color': self.filter_colour})

                        frame_worksheets[frame_list_index].conditional_format(blue_value, {'type': '2_color_scale',
                                                            'min_color': "#000000",
                                                            'max_color': self.filter_colour})
                    #Remove unused rows and columns
                    frame_worksheets[frame_list_index].set_default_row(hide_unused_rows=True)
                    frame_worksheets[frame_list_index].set_column('{}:XFD'.format(to_excel_coords(width+3)), None, None, {'hidden': True})
                    
                    #Increment values and move to next line
                    line += 3
                
                frame_video_index += self.frame_skip
                frame_list_index += 1
                
        
            #When finished converting, save the workbook
            self.status = "Saving..."
            self.progress = 100

            #Save and close
            self.status = "Saving..."        
            workbook.close()
            self.status = "Image saved!"

            self.finished = True
