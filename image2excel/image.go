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
	"strings"

	"github.com/xuri/excelize/v2"
	"golang.org/x/image/draw"
)

const (
	infoSheetName string = "Info"
)

var creditsDisplay = "image2excel GitHub repo"

type ConvertImageOptions struct {
	ScaleFactor float64
	Width       int
	Height      int
	DoLogging   bool
	SheetName   string
}

func setImageOptionDefaults(opts *ConvertImageOptions) {

	if opts.ScaleFactor == 0 {
		opts.ScaleFactor = 0.5
	} else if opts.SheetName == "" {
		opts.SheetName = "ConvertedImage"
	}

}

// Converts image bytes into bytes
func ConvertImage(imageData *bytes.Buffer, opts *ConvertImageOptions) (*bytes.Buffer, error) {
	setImageOptionDefaults(opts)

	var spreadsheetBytes bytes.Buffer
	var mainSheetName string

	var imageHeight int
	var imageWidth int

	var imageBytes []byte

	mainSheetName = path.Base(opts.SheetName)

	imageBytes = imageData.Bytes()

	// Decode image
	img, _, err := image.Decode(imageData)
	checkError(err)

	// Set scaling parameters

	if opts.Width == 0 && opts.Height == 0 {
		imageWidth = img.Bounds().Dx()
		imageHeight = img.Bounds().Dy()
	} else if opts.Width != 0 {
		imageWidth = opts.Width
		imageHeight = int((float64(img.Bounds().Dx()) / float64(imageWidth)) * float64(imageHeight))
	} else if opts.Height != 0 {
		imageHeight = opts.Height
		imageWidth = int((float64(img.Bounds().Dy()) / float64(imageHeight)) * float64(imageWidth))
	}

	if opts.ScaleFactor > 1.0 {
		return nil, errors.New("scale factor too large")
	} else {
		imageWidth = int(float64(imageWidth) * opts.ScaleFactor)
		imageHeight = int(float64(imageHeight) * opts.ScaleFactor)
	}

	// Resize image

	var resized draw.Image

	// Don't resize if no parameters have been changed.
	if !(opts.ScaleFactor == 1.0 && imageWidth == img.Bounds().Dx() && imageHeight == img.Bounds().Dy()) {
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

	if opts.DoLogging {

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

	if opts.DoLogging {

		log.Printf("Image basename: %s", strings.TrimSuffix(mainSheetName, filepath.Ext(mainSheetName)))
		log.Printf("Image extension: %s", strings.ToLower(filepath.Ext(mainSheetName)))

	}

	f.AddPictureFromBytes(
		infoSheetName,
		"A5",
		&excelize.Picture{
			Extension: strings.ToLower(filepath.Ext(mainSheetName)),
			File:      imageBytes,
			Format: &excelize.GraphicOptions{
				Positioning: "oneCell",
			},
		},
	)

	// Add main image sheet

	addImageToSheet(f, mainSheetName, resized, opts.DoLogging)

	// Finish up and "save" file

	f.Write(&spreadsheetBytes)
	f.Close()

	if opts.DoLogging {
		log.Println("Returning...")
	}

	return &spreadsheetBytes, nil

}
