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

Video is the same except each frame is it's own spreadsheet.

---
### Prerequisites

Libraries required can be installed with:

`pip install -r requirements.txt`

---

### Command syntax

`image2excel-cmd.py [-h] [--scale SCALE] [--filter FILTER] [--frameskip FRAMESKIP] [--forceframeskip] [--videocut VIDEOCUT] type file_path output_path mode`

Use `image2excel-cmd.py -h` for help. Information about arguments is displayed below.

**Media type:** `video` or `image`

**File path:** The location of the original iamge. `test_image` can be for one of the test images, no test videos (scale of 0.1 is recommended).

**Output path:** The output path of the Excel spreadsheet

**Mode**:

 - `GREYSCALE` *(applies a greyscale effect to the image)* 
 - `RGB` *(no filter applied)*
 - `FILTER` *(applies a colour filter to the image)*

#### **Optional arguments**

**Scale:** The scale of the image. For images over 1000px in size, a scale of less than 0.25 is recommended.

**Filter:** The colour of the filter to be applied to the image.

##### Video only

**Frameskip:** How many frames to skip between each spreadsheet. 50 is recommended.

**Force frame skip:** Force the frame skip to be lower than 25

**Videocut:** How much to shorten the video by

---

### Webserver

Make sure you run this in the `web` directory.

`python image2excel-web.py <port>`

**Note:** The port will default to `5000`

---

### TODO

 - Implement video conversion for web server
 - Implement scale for video :white_check_mark:
 - Frame skip config :white_check_mark:
 - Cut down video length as option :white_check_mark: