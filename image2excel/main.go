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
	"runtime"

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

// Converts image bytes into bytes
func Convert(image_data *bytes.Buffer, file_name string, do_output bool, scale_factor float64, width int, height int) (*bytes.Buffer, error) {

	var spreadsheet_bytes bytes.Buffer
	var main_sheet_name string

	var image_height int
	var image_width int

	main_sheet_name = path.Base(file_name)

	// Decode image
	img, _, err := image.Decode(image_data)
	check_error(err)

	// Set scaling parameters

	if width == 0 && height == 0 {
		image_width = img.Bounds().Dx()
		image_height = img.Bounds().Dy()
	} else if width != 0 {
		image_width = width
		image_height = int((float64(img.Bounds().Dx()) / float64(image_width)) * float64(image_height))
	} else if height != 0 {
		image_height = height
		image_width = int((float64(img.Bounds().Dy()) / float64(image_height)) * float64(image_width))
	}

	if scale_factor > 1.0 {
		return nil, errors.New("scale factor too large")
	} else {
		image_width = int(float64(image_width) * scale_factor)
		image_height = int(float64(image_height) * scale_factor)
	}

	// Resize image

	var resized draw.Image

	// Don't resize if no parameters have been changed.
	if !(scale_factor == 1.0 && image_width == img.Bounds().Dx() && image_height == img.Bounds().Dy()) {
		log.Println("Resizing...")

		resized = image.NewNRGBA(image.Rect(0, 0, image_width, image_height))

		draw.NearestNeighbor.Scale(resized, image.Rect(0, 0, image_width, image_height), img, img.Bounds(), draw.Over, nil)
	} else {
		log.Println("Skipping resize...")

		resized = image.NewRGBA(img.Bounds())

		for y := 0; y < image_height; y++ {
			for x := 0; x < image_width; x++ {
				resized.Set(x, y, img.At(x, y))
			}
		}
	}

	// Create Excel file

	f := excelize.NewFile()
	check_error(err)

	index, err := f.NewSheet(main_sheet_name)
	check_error(err)
	f.SetActiveSheet(index)

	s, err := f.NewStreamWriter(main_sheet_name)
	check_error(err)

	// Loop through at set pixels

	for y := 0; y < image_height; y++ {
		if do_output {
			var m runtime.MemStats
			runtime.ReadMemStats(&m)

			fmt.Printf("\rProgress: %d%% Mem: %d...", int((float64(y)/float64(image_height))*100.0), m.Sys)
		}

		cell, _ := excelize.CoordinatesToCellName(1, y)

		s.SetRow(cell, get_image_row(resized, y, Red))

	}

	s.Flush()

	if do_output {
		var m runtime.MemStats
		runtime.ReadMemStats(&m)

		fmt.Printf("\rProgress: 100%%... Mem: %d...\n", m.Sys)
		log.Println("Writing to buffer...")
	}

	// Finish up and "save" file

	f.Write(&spreadsheet_bytes)
	f.Close()

	if do_output {
		log.Println("Returning...")
	}

	return &spreadsheet_bytes, nil

}

// Gets a row of pixels from an image
func get_image_row(i image.Image, row int, colour PixelType) []interface{} {

	img_interface := make([]interface{}, i.Bounds().Dx())

	for x := 0; x < i.Bounds().Dx(); x++ {
		r, _, _, _ := i.At(x, row).RGBA()

		img_interface[x] = uint8(r / 255)
	}

	return img_interface

}

// Converts image to 2D array
func img_to_array(i image.Image) [][]Pixel {

	img_array := make([][]Pixel, i.Bounds().Dy())

	for y := 0; y < i.Bounds().Dy(); y++ {

		img_array[y] = make([]Pixel, i.Bounds().Dx())

		for x := 0; x < i.Bounds().Dx(); x++ {

			r, g, b, a := i.At(x, y).RGBA()

			img_array[y][x] = Pixel{
				R: uint8(r / 255),
				G: uint8(g / 255),
				B: uint8(b / 255),
				A: uint8(a / 255),
			}

		}

	}

	return img_array

}
