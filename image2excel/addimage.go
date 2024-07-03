package image2excel

import (
	"fmt"
	"image"
	"log"
	"runtime"

	"github.com/xuri/excelize/v2"
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

var conditionalFormatRed excelize.ConditionalFormatOptions = excelize.ConditionalFormatOptions{
	Type:     "2_color_scale",
	Criteria: "=",
	MinType:  "min",
	MaxType:  "max",
	MinColor: "#000000",
	MaxColor: "#FF0000",
}

var conditionalFormatGreen excelize.ConditionalFormatOptions = excelize.ConditionalFormatOptions{
	Type:     "2_color_scale",
	Criteria: "=",
	MinType:  "min",
	MaxType:  "max",
	MinColor: "#000000",
	MaxColor: "#00FF00",
}

var conditionalFormatBlue excelize.ConditionalFormatOptions = excelize.ConditionalFormatOptions{
	Type:     "2_color_scale",
	Criteria: "=",
	MinType:  "min",
	MaxType:  "max",
	MinColor: "#000000",
	MaxColor: "#0000FF",
}

// Used by both ConvertImage and ConvertVideo
func addImageToSheet(f *excelize.File, name string, _image image.Image, doLogging bool) error {

	index, err := f.NewSheet(name)
	if err != nil {
		return err
	}

	f.SetActiveSheet(index)

	s, err := f.NewStreamWriter(name)
	checkError(err)

	//Set column widths

	err = s.SetColWidth(1, _image.Bounds().Dx(), 2.5)
	checkError(err)

	// Loop through at set pixels

	var writeY int = 0

	var m runtime.MemStats

	for y := 0; y < _image.Bounds().Dy(); y++ {
		if doLogging {
			runtime.ReadMemStats(&m)

			fmt.Printf("\rProgress: %d%% Mem: %d...", int((float64(y)/float64(_image.Bounds().Dy()))*100.0), m.Sys)
		}

		rowR, _ := excelize.CoordinatesToCellName(1, writeY+1)
		rowG, _ := excelize.CoordinatesToCellName(1, writeY+2)
		rowB, _ := excelize.CoordinatesToCellName(1, writeY+3)

		rowEndR, _ := excelize.CoordinatesToCellName(_image.Bounds().Dx()+2, writeY+1)
		rowEndG, _ := excelize.CoordinatesToCellName(_image.Bounds().Dx()+2, writeY+2)
		rowEndB, _ := excelize.CoordinatesToCellName(_image.Bounds().Dx()+2, writeY+3)

		s.SetRow(rowR, append(getImageRow(_image, y, Red), 0, 255))
		f.SetConditionalFormat(name, fmt.Sprintf("%s:%s", rowR, rowEndR), []excelize.ConditionalFormatOptions{conditionalFormatRed})

		s.SetRow(rowG, append(getImageRow(_image, y, Green), 0, 255))
		f.SetConditionalFormat(name, fmt.Sprintf("%s:%s", rowG, rowEndG), []excelize.ConditionalFormatOptions{conditionalFormatGreen})

		s.SetRow(rowB, append(getImageRow(_image, y, Blue), 0, 255))
		f.SetConditionalFormat(name, fmt.Sprintf("%s:%s", rowB, rowEndB), []excelize.ConditionalFormatOptions{conditionalFormatBlue})

		writeY += 3

	}

	s.Flush()

	if doLogging {
		var m runtime.MemStats
		runtime.ReadMemStats(&m)

		fmt.Printf("\rProgress: 100%%... Mem: %d...\n", m.Sys)
		log.Println("Writing to buffer...")
	}

	return nil

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
