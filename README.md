# image2excel
A program which converts images into excel spreadsheets.
### Prerequisites
Libraries required for the program to run.

**Command line program:**

A simple program that can be executed from the command line.

`pip install xlsxwriter` 

 **Webserver**

 A web based interface for the program which runs on [localhost:80](http://localhost:80) by default.

`pip install flask`

`pip install xlsxwriter`

### Command syntax

**Command line program.**
     `python image2excel.py <image to convert> <output path> <mode> <filter>`

Modes:

 - GS *(greyscale)*
 - RGB *(red, green and blue)*
 - FILTER *(applies a filter to the image)*

**Note:** The filter argument (a hexidecimal colour) is only required if you set the mode to `FILTER`

**Webserver**

    python image2excelserver.py
	
