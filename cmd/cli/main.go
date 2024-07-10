package main

import (
	"bytes"
	"fmt"
	"log"
	"os"
	i2e "sccreeper/image2excel/image2excel"

	"github.com/urfave/cli/v2"
	"github.com/xuri/excelize/v2"
)

var scale float64
var outputPath string

var imageWidth int = 0
var imageHeight int = 0

var perFile int
var interval int

var experimental bool

func main() {

	app := &cli.App{
		Name:        "image2excel",
		Usage:       "Images and videos to Excel spreadsheets",
		Description: "CLI for image2excel. Converts images and video files to Excel spreadsheets. Can accept png, jpg, and gif images. Can accept video formats supported by OpenCV.",
		Action:      convertImage,
		Commands: []*cli.Command{
			{
				Name:    "image",
				Aliases: []string{"i"},
				Usage:   "./i2e image <path>",
				Action:  convertImage,
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
						Name:        "scale",
						Aliases:     []string{"s"},
						Usage:       "Scale factor of the image",
						Destination: &scale,
						Required:    true,
						Value:       0.5,
					},
				},
			},
			{
				Name:    "video",
				Aliases: []string{"v"},
				Usage:   "./i2e video <path>",
				Action:  convertVideo,
				Flags: []cli.Flag{
					&cli.Float64Flag{
						Name:        "scale",
						Aliases:     []string{"s"},
						Usage:       "Proportionally scale the size of each frame. Recommended is < 0.5. Cannot be > 1.",
						Destination: &scale,
						Required:    true,
						Value:       0.5,
					},
					&cli.IntFlag{
						Name:        "perfile",
						Aliases:     []string{"p"},
						Usage:       "How many frames to add to each spreadsheet file.",
						Destination: &perFile,
						Value:       15,
					},
					&cli.IntFlag{
						Name:        "interval",
						Aliases:     []string{"i"},
						Usage:       "Gap to leave between every nth frame.",
						Destination: &interval,
						Value:       5,
					},
					&cli.StringFlag{
						Name:        "output",
						Aliases:     []string{"o"},
						Usage:       "Output directory. Can't already exist.",
						Destination: &outputPath,
						Value:       "out",
					},
					&cli.BoolFlag{
						Name:        "experimental",
						Aliases:     []string{"e"},
						Usage:       "Use the experimental ConvertVideoCustom method. Outputted files may not work in all programs.",
						Destination: &experimental,
						Value:       false,
					},
				},
			},
			{
				Name:        "validate",
				Description: "Used for debugging. Opens the files using Excelize.",
				Action:      validate,
			},
		},
	}

	if err := app.Run(os.Args); err != nil {
		log.Fatal(err)
	}

}

func convertImage(ctx *cli.Context) error {

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

	conv, err := i2e.ConvertImage(&fileData, &i2e.ConvertImageOptions{
		SheetName:   ctx.Args().First(),
		ScaleFactor: scale,
		Height:      imageHeight,
		Width:       imageWidth,
		DoLogging:   true,
	})

	if err != nil {
		fmt.Printf("Error: %s\n", err.Error())
		os.Exit(1)
	}

	log.Println("Saving file...")

	os.WriteFile("out.xlsx", conv.Bytes(), os.ModePerm)

	log.Println("Converted successfully!")

	return nil

}

func convertVideo(ctx *cli.Context) error {

	if experimental {

		i2e.ConvertVideoCustom(ctx.Args().First(), outputPath, &i2e.ConvertVideoOptions{
			Interval:  interval,
			PerFile:   perFile,
			Scale:     scale,
			DoLogging: true,
		})

	} else {

		i2e.ConvertVideo(ctx.Args().First(), outputPath, &i2e.ConvertVideoOptions{
			Interval:  interval,
			PerFile:   perFile,
			Scale:     scale,
			DoLogging: true,
		})

	}

	return nil

}

func validate(ctx *cli.Context) error {

	fileName := ctx.Args().First()

	workbook, err := excelize.OpenFile(fileName)
	if err != nil {
		return err
	}

	for _, v := range workbook.GetSheetMap() {
		fmt.Println(v)
	}

	workbook.Close()

	fmt.Println("Opened successfully.")

	return nil

}
