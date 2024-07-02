package main

import (
	"bytes"
	"fmt"
	"log"
	"os"
	"sccreeper/image2excel/image2excel"

	"github.com/urfave/cli/v2"
)

var imageWidth int = 0
var imageHeight int = 0
var scaleFactor float64

func main() {

	app := &cli.App{
		Name:        "image2excel",
		Description: "Convert images to Excel spreadsheets",
		Action:      _convert,
		Flags: []cli.Flag{

			&cli.IntFlag{
				Name:        "width",
				Usage:       "Width of the image to be scaled down to (cannot be used with scale factor)",
				Destination: &imageWidth,
				Required:    false,
				Value:       0,
			},
			&cli.IntFlag{
				Name:        "height",
				Usage:       "Height of the image to be scaled down to (cannot be used with scale factor)",
				Destination: &imageHeight,
				Required:    false,
				Value:       0,
			},
			&cli.Float64Flag{
				Name:        "scalefactor",
				Aliases:     []string{"s"},
				Usage:       "Scale factor of the image",
				Destination: &scaleFactor,
				Required:    true,
				Value:       1.0,
			},
		},
	}

	if err := app.Run(os.Args); err != nil {
		log.Fatal(err)
	}

}

func _convert(ctx *cli.Context) error {

	var fileData bytes.Buffer

	if ctx.Args().First() == "" {
		log.Println("No file!")
		log.Println("Exiting...")
		os.Exit(1)
	} else {

		log.Println("Reading file...")

		b, err := os.ReadFile(ctx.Args().First())
		if err != nil {
			panic(err)
		}

		fileData.Write(b)

	}

	log.Println("Converting file...")

	conv, err := image2excel.ConvertImage(&fileData, ctx.Args().First(), true, scaleFactor, imageHeight, imageWidth)

	if err != nil {
		fmt.Printf("Error: %s\n", err.Error())
		os.Exit(1)
	}

	log.Println("Saving file...")

	os.WriteFile("out.xlsx", conv.Bytes(), os.ModePerm)

	log.Println("Converted successfully!")

	return nil

}
