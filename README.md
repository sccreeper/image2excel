# image2excel
A program which converts images and video into Excel spreadsheets.

Video only works for cmd utility atm. Videos also take a long time to save.

---

### Example conversion
Original image:

![Original image](assets/test_image_readme.JPG)
---
Result as viewed in Excel:

![Conversion](assets/test_conversion.png)

---
### Prerequisites

Libraries required can be installed with:

`pip install -r requirements.txt`

---

### Command syntax

`python image2excel-cmd.py <media type> <source path> <output path> <scale> <mode> <filter colour (optional)>`

**Example:** `python image2excel-cmd.py image test_image test_ouput 0.1 RGB`

**Example with filter:** `python image2excel-cmd.py image test_image test_ouput 0.1 FILTER #35C391`

**Media type:** `video` or `image`

**Source path:** The location of the original iamge. `test_image` can be for one of the test images, no test videos (scale of 0.1 is recommended).

**Output path:** The output path of the Excel spreadsheet

**Scale:** The scale of the image. For images over 1000px in size, a scale of less than 0.25 is recommended.

**Mode**:

 - `GREYSCALE` *(applies a greyscale effect to the image)* 
 - `RGB` *(no filter applied)*
 - `FILTER` *(applies a colour filter to the image)*

---

### Webserver

Make sure you run this in the `web` directory.

`python image2excel-web.py <port>`

**Note:** The port will default to `5000`

---

### TODO

 - Implement video conversion for web server
 - Implement scale for video
 - Frame skip config
 - Cut down video length as option