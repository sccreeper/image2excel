package image2excel

import (
	"bytes"
	"errors"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"log"
	"path"
	"path/filepath"
	"runtime"
	"strings"

	"github.com/xuri/excelize/v2"
	"golang.org/x/image/draw"
)

type Pixel struct {
	R uint8
	G uint8
	B uint8
	A uint8
}

type PixelType uint8

const (
	Red   PixelType = 0
	Green PixelType = 1
	Blue  PixelType = 2
	Alpha PixelType = 3
)

var conditional_format_red excelize.ConditionalFormatOptions = excelize.ConditionalFormatOptions{
	Type:     "2_color_scale",
	Criteria: "=",
	MinType:  "min",
	MaxType:  "max",
	MinColor: "#000000",
	MaxColor: "#FF0000",
}

var conditional_format_green excelize.ConditionalFormatOptions = excelize.ConditionalFormatOptions{
	Type:     "2_color_scale",
	Criteria: "=",
	MinType:  "min",
	MaxType:  "max",
	MinColor: "#000000",
	MaxColor: "#00FF00",
}

var conditional_format_blue excelize.ConditionalFormatOptions = excelize.ConditionalFormatOptions{
	Type:     "2_color_scale",
	Criteria: "=",
	MinType:  "min",
	MaxType:  "max",
	MinColor: "#000000",
	MaxColor: "#0000FF",
}

const (
	infoSheetName string = "Info"
)

var creditsDisplay = "image2excel GitHub repo"

// Converts image bytes into bytes
func Convert(imageData *bytes.Buffer, file_name string, doOutput bool, scaleFactor float64, width int, height int) (*bytes.Buffer, error) {

	var spreadsheetBytes bytes.Buffer
	var mainSheetName string

	var imageHeight int
	var imageWidth int

	var imageBytes []byte

	mainSheetName = path.Base(file_name)

	imageBytes = imageData.Bytes()

	// Decode image
	img, _, err := image.Decode(imageData)
	checkError(err)

	// Set scaling parameters

	if width == 0 && height == 0 {
		imageWidth = img.Bounds().Dx()
		imageHeight = img.Bounds().Dy()
	} else if width != 0 {
		imageWidth = width
		imageHeight = int((float64(img.Bounds().Dx()) / float64(imageWidth)) * float64(imageHeight))
	} else if height != 0 {
		imageHeight = height
		imageWidth = int((float64(img.Bounds().Dy()) / float64(imageHeight)) * float64(imageWidth))
	}

	if scaleFactor > 1.0 {
		return nil, errors.New("scale factor too large")
	} else {
		imageWidth = int(float64(imageWidth) * scaleFactor)
		imageHeight = int(float64(imageHeight) * scaleFactor)
	}

	// Resize image

	var resized draw.Image

	// Don't resize if no parameters have been changed.
	if !(scaleFactor == 1.0 && imageWidth == img.Bounds().Dx() && imageHeight == img.Bounds().Dy()) {
		log.Println("Resizing...")

		resized = image.NewNRGBA(image.Rect(0, 0, imageWidth, imageHeight))

		draw.NearestNeighbor.Scale(resized, image.Rect(0, 0, imageWidth, imageHeight), img, img.Bounds(), draw.Over, nil)
	} else {
		log.Println("Skipping resize...")

		resized = image.NewRGBA(img.Bounds())

		for y := 0; y < imageHeight; y++ {
			for x := 0; x < imageWidth; x++ {
				resized.Set(x, y, img.At(x, y))
			}
		}
	}

	if doOutput {

		log.Printf("Image res: %dx%d", resized.Bounds().Dx(), resized.Bounds().Dy())

	}

	// Create Excel file

	f := excelize.NewFile()
	checkError(err)

	f.DeleteSheet("Sheet1")

	// Add "info" sheet
	_, err = f.NewSheet(infoSheetName)
	checkError(err)

	// Credits link

	f.SetCellHyperLink(infoSheetName, "A1", "https://github.com/sccreeper/image2excel", "External", excelize.HyperlinkOpts{
		Display: &(creditsDisplay),
	})
	f.SetCellValue(infoSheetName, "A1", creditsDisplay)

	link_style, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "#22934B", Underline: "single"},
	})
	checkError(err)

	err = f.SetCellStyle(infoSheetName, "A1", "A1", link_style)
	checkError(err)

	// Add 'metadata'

	f.SetCellValue(infoSheetName, "A3", fmt.Sprintf("Original image: %s", mainSheetName))
	f.SetCellValue(infoSheetName, "A2", fmt.Sprintf("Original resolution: %dx%d Resized resolution: %dx%d Spreadsheet resolution: %dx%d", img.Bounds().Dx(), img.Bounds().Dy(), resized.Bounds().Dx(), resized.Bounds().Dy(), resized.Bounds().Dx(), resized.Bounds().Dy()*3))

	if doOutput {

		log.Printf("Image basename: %s", strings.TrimSuffix(mainSheetName, filepath.Ext(mainSheetName)))
		log.Printf("Image extension: %s", strings.ToLower(filepath.Ext(mainSheetName)))

	}

	f.AddPictureFromBytes(
		infoSheetName,
		"A5",
		strings.TrimSuffix(mainSheetName, filepath.Ext(mainSheetName)),
		strings.ToLower(filepath.Ext(mainSheetName)),
		imageBytes,
		&excelize.GraphicOptions{
			Positioning: "oneCell",
		},
	)

	// Add main image sheet

	index, err := f.NewSheet(mainSheetName)
	checkError(err)
	f.SetActiveSheet(index)

	s, err := f.NewStreamWriter(mainSheetName)
	checkError(err)

	//Set column widths

	err = s.SetColWidth(1, resized.Bounds().Dx(), 2.5)
	checkError(err)

	// Loop through at set pixels

	var writeY int = 0

	var m runtime.MemStats

	for y := 0; y < imageHeight; y++ {
		if doOutput {
			runtime.ReadMemStats(&m)

			fmt.Printf("\rProgress: %d%% Mem: %d...", int((float64(y)/float64(imageHeight))*100.0), m.Sys)
		}

		rowR, _ := excelize.CoordinatesToCellName(1, writeY+1)
		rowG, _ := excelize.CoordinatesToCellName(1, writeY+2)
		rowB, _ := excelize.CoordinatesToCellName(1, writeY+3)

		rowEndR, _ := excelize.CoordinatesToCellName(resized.Bounds().Dx()+2, writeY+1)
		rowEndG, _ := excelize.CoordinatesToCellName(resized.Bounds().Dx()+2, writeY+2)
		rowEndB, _ := excelize.CoordinatesToCellName(resized.Bounds().Dx()+2, writeY+3)

		s.SetRow(rowR, append(getImageRow(resized, y, Red), 0, 255))
		f.SetConditionalFormat(mainSheetName, fmt.Sprintf("%s:%s", rowR, rowEndR), []excelize.ConditionalFormatOptions{conditional_format_red})

		s.SetRow(rowG, append(getImageRow(resized, y, Green), 0, 255))
		f.SetConditionalFormat(mainSheetName, fmt.Sprintf("%s:%s", rowG, rowEndG), []excelize.ConditionalFormatOptions{conditional_format_green})

		s.SetRow(rowB, append(getImageRow(resized, y, Blue), 0, 255))
		f.SetConditionalFormat(mainSheetName, fmt.Sprintf("%s:%s", rowB, rowEndB), []excelize.ConditionalFormatOptions{conditional_format_blue})

		writeY += 3

	}

	s.Flush()

	if doOutput {
		var m runtime.MemStats
		runtime.ReadMemStats(&m)

		fmt.Printf("\rProgress: 100%%... Mem: %d...\n", m.Sys)
		log.Println("Writing to buffer...")
	}

	// Finish up and "save" file

	f.Write(&spreadsheetBytes)
	f.Close()

	if doOutput {
		log.Println("Returning...")
	}

	return &spreadsheetBytes, nil

}

// Gets a row of pixels from an image
func getImageRow(i image.Image, row int, colour PixelType) []interface{} {

	imgInterface := make([]interface{}, i.Bounds().Dx())

	for x := 0; x < i.Bounds().Dx(); x++ {
		r, g, b, _ := i.At(x, row).RGBA()

		switch colour {
		case Red:
			imgInterface[x] = uint8(r / 255)
		case Green:
			imgInterface[x] = uint8(g / 255)
		case Blue:
			imgInterface[x] = uint8(b / 255)
		}

	}

	return imgInterface

}

// Converts image to 2D array
func imgToArray(i image.Image) [][]Pixel {

	imgArray := make([][]Pixel, i.Bounds().Dy())

	for y := 0; y < i.Bounds().Dy(); y++ {

		imgArray[y] = make([]Pixel, i.Bounds().Dx())

		for x := 0; x < i.Bounds().Dx(); x++ {

			r, g, b, a := i.At(x, y).RGBA()

			imgArray[y][x] = Pixel{
				R: uint8(r / 255),
				G: uint8(g / 255),
				B: uint8(b / 255),
				A: uint8(a / 255),
			}

		}

	}

	return imgArray

}
