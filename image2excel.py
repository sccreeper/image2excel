from PIL import Image
import xlsxwriter
from datetime import datetime
import sys
import os

image_name = sys.argv[1]

#See if help has been selected
if image_name == 'help':
    print("""
Image to Excel converter. Written by Oscar Peace AKA sccreeper.

Command syntax:
    excel converter.py <filename> <file output> <mode>

    <filename> can be 'test_image' if you just want to test the program.
Modes:
    GS (Greyscale)
    RGB
    FILTER <HEX COLOUR>
""")
    pause = input('')
    exit()

if image_name == 'test_image':
    image_name = 'assets/test_image.JPG'

#Set the variables fot cmd args
output = sys.argv[2]
mode = sys.argv[3]

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

#Open the image.
im = Image.open(image_name)
width, height = im.size

rgb_im = im.convert('RGB')

#print(r,g,b)

#Create the spreadsheet
workbook = xlsxwriter.Workbook('{}.xlsx'.format(output))

worksheet = workbook.add_worksheet("Image")

i=0

#Iterate through every pixel row by row.
while i < height:
    
    print("Converting. {}% {}".format(round((i/height)*100)+1, gen_bar(round((i/height)*100)+1)), end="\r")
    
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

    #RGB Mode
    if mode == 'RGB':
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
    if mode == 'GS':
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
    if mode == 'FILTER':
        worksheet.conditional_format(red_value, {'type': '2_color_scale',
                                             'min_color': "#000000",
                                             'max_color': sys.argv[4]})
        worksheet.conditional_format(green_value, {'type': '2_color_scale',
                                             'min_color': "#000000",
                                             'max_color': sys.argv[4]})

        worksheet.conditional_format(blue_value, {'type': '2_color_scale',
                                             'min_color': "#000000",
                                             'max_color': sys.argv[4]})

    i+=3 
    #print(i)

print("\nFinishing off...")
#Make the cells nice and square
worksheet.set_column(0, width, 2.14)

worksheet.write('A1', 'Image produced by the Image to Excel converter. Zoom out to view the full image. Converter made by Oscar Peace AKA thecodedevourer')
worksheet.write('A2', 'Original dimensions: {}px X {}px ({} pixels). Spreadsheet dimensions: {}cells X {}cells. ({} cells)'.format(width,height,width*height, width, height*3, width*(height*3)))
worksheet.write_url('B1', 'https://github.com/sccreeper/image2excel/', string='View the source code on GitHub')

#Insert the original image into another sheet
worksheet1 = workbook.add_worksheet("Original Image")

worksheet1.write('A1', "Original image: '{}'".format(os.path.basename(sys.argv[1])))
worksheet1.insert_image('A2', sys.argv[1])

#Remove unused
print("Removing unused rows and columns...")
worksheet.set_default_row(hide_unused_rows=True)
worksheet.set_column('{}:XFD'.format(excel_coords(width+3)), None, None, {'hidden': True})

#Save and close
print("Saving...")        
workbook.close()
print("Image saved to {}".format(sys.argv[2] + ".xlsx"))

#Ask if user wants to open spreadsheet or not
if input("Open image (y/n)").lower() == "y":
    os.system('start {}'.format(sys.argv[2] + '.xlsx'))
    print("Opening...")
else:
    exit()
