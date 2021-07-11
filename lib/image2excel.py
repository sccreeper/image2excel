__name__ = "image2excel.lib"

from PIL import Image
import xlsxwriter
from datetime import datetime
from enum import Enum

class Mode(Enum):
    RGB = 1
    GREYSCALE = 2
    FILTER = 3

class Converter:
    def __init__(self, file_name: str, output_path: str, mode=Mode.RGB, scale=1, filter_colour="#FFFFFF"):
        """The main converter class.

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
    
    #Convert normal coordinates (1,1) into excel coordiantes A1
    def __excel_coords(self, x):
        string = ""
        while x > 0:
            x, remainder = divmod(x - 1, 26)
            string = chr(65 + remainder) + string

        return str(string)

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
            
            #Iterate through each pixel in every column and write their RGB values to seperate lines in a spreadsheet.
            for pixel in range(width):
                r, g, b = rgb_im.getpixel((pixel, i))

                worksheet.write(i+2, pixel, r)
                worksheet.write(i+3, pixel, g)
                worksheet.write(i+4, pixel, b)

            #Get the locations on the spreadsheet needed for conditional formatting
            red_value = 'A' + str(i) + ':' + self.__excel_coords(width+5)+ str(1)
            green_value = 'A' + str(i+1) + ':' + self.__excel_coords(width+5) + str(2)
            blue_value = 'A' + str(i+2) + ':' + self.__excel_coords(width+5) + str(3)

            #Write 0 and 255 at the end of every row so conditional formatting works
            worksheet.write(i, width, 0)
            worksheet.write(i, width+1, 255)

            worksheet.write(i+1, width, 0)
            worksheet.write(i+1, width+1, 255)

            worksheet.write(i+2, width, 0)
            worksheet.write(i+2, width+1, 255)

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
        worksheet.set_column('{}:XFD'.format(self.__excel_coords(width+3)), None, None, {'hidden': True})

        #Save and close
        self.status = "Saving..."        
        workbook.close()
        self.status = "Image saved!"

        self.finished = True
