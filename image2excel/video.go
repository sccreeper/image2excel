package image2excel

import (
	"fmt"
	"image"
	"os"
	"runtime"
	"sccreeper/image2excel/image2excel/sheets"
	"sync"
	"sync/atomic"

	"github.com/xuri/excelize/v2"
	"gocv.io/x/gocv"
)

// videoUri - The uri of the source video to be converted.
//
// interval - gap to leave between each frame.
//
// perFile - How many frames to have perFile. If zero, all frames in one file.
//
// scale - How much to scale each frame up or down by. Values >1 are not recommended for resolutions greater than 480p
//
// doLogging - Enable logging
//
// outputDir - Output location for file
type ConvertVideoOptions struct {
	Interval  int
	PerFile   int
	Scale     float64
	DoLogging bool
}

func fillDefaults(opts *ConvertVideoOptions) {
	if opts.Interval == 0 {
		opts.Interval = 5
	} else if opts.PerFile == 0 {
		opts.PerFile = 10
	} else if opts.Scale == 0.0 {
		opts.Scale = 0.25
	}
}

// Outputs frames to a folder containing spreadsheets numbered 0-n. Where n is the number of frames/frame interval as specified in opts.
func ConvertVideo(videoUri string, outputDir string, opts *ConvertVideoOptions) {
	fillDefaults(opts)

	// Figure out frames in file

	vc, err := gocv.VideoCaptureFile(videoUri)
	checkError(err)

	frameCount := int(vc.Get(gocv.VideoCaptureFrameCount))

	vc.Close()

	//
	// Setup
	//

	os.Mkdir(outputDir, 0755)

	// Pre-generate the frame numbers in order to avoid horrible for loop stuff.

	frameNumbers := make([]int, 0, int(frameCount/opts.Interval))

	for i := 0; i < frameCount; i += opts.Interval {
		frameNumbers = append(frameNumbers, i)
	}

	maxGoroutines := runtime.NumCPU()
	var numGoroutines atomic.Int32
	var progress atomic.Int32

	semaphoreChannel := make(chan struct{}, maxGoroutines)
	defer close(semaphoreChannel)
	var wg sync.WaitGroup

	//
	// Main loop
	//

	for i := 0; i < len(frameNumbers); i += opts.PerFile {

		wg.Add(1)
		semaphoreChannel <- struct{}{}

		numGoroutines.Add(1)

		//
		// Main goroutine
		//

		go func(done chan struct{}, _wg *sync.WaitGroup, sheetNumber int, startIndex int) {

			vc, err := gocv.VideoCaptureFile(videoUri)
			checkError(err)

			// Cleanup
			defer func() {
				vc.Close()
				<-semaphoreChannel
				numGoroutines.Add(-1)
				_wg.Done()
			}()

			mats := getImageMats(vc, frameNumbers[startIndex:startIndex+opts.PerFile]...)
			frames := make([]image.Image, 0)

			for _, mat := range mats {

				// Convert mat

				convertedMat := gocv.NewMat()

				mat.ConvertTo(&convertedMat, gocv.MatTypeCV8UC1)

				gocv.Resize(convertedMat, &convertedMat, image.Pt(0, 0), opts.Scale, opts.Scale, gocv.InterpolationNearestNeighbor)

				frame, err := convertedMat.ToImage()
				if err != nil {
					panic(err)
				}

				mat.Close()
				convertedMat.Close()

				frames = append(frames, frame)

			}

			f := excelize.NewFile()

			for i, frame := range frames {
				addImageToSheet(f, fmt.Sprintf("%d", frameNumbers[startIndex+i]), frame, false)

				if opts.DoLogging {
					fmt.Printf(
						"\rProgress: %d%% Frame: %3d/%3d Active goroutines: %2d/%2d...",
						percentage(int(progress.Load()), len(frameNumbers)),
						frameNumbers[progress.Load()],
						frameNumbers[len(frameNumbers)-1],
						numGoroutines.Load(),
						runtime.NumCPU(),
					)
				}

				progress.Add(1)
			}

			f.DeleteSheet("Sheet1")
			f.SaveAs(fmt.Sprintf("%s/%d.xlsx", outputDir, sheetNumber))
			f.Close()

		}(semaphoreChannel, &wg, (i+opts.PerFile)/opts.PerFile, i)
	}

	wg.Wait()

	if opts.DoLogging {
		fmt.Println("\nCleaning up...")
	}

}

func ConvertVideoCustom(videoUri string, outputDir string, opts *ConvertVideoOptions) {
	fillDefaults(opts)

	// Figure out frames in file

	vc, err := gocv.VideoCaptureFile(videoUri)
	checkError(err)

	frameCount := int(vc.Get(gocv.VideoCaptureFrameCount))

	vc.Close()

	//
	// Setup
	//

	os.Mkdir(outputDir, 0755)

	// Pre-generate the frame numbers in order to avoid horrible for loop stuff.

	frameNumbers := make([]int, 0, int(frameCount/opts.Interval))

	for i := 0; i < frameCount; i += opts.Interval {
		frameNumbers = append(frameNumbers, i)
	}

	maxGoroutines := runtime.NumCPU()
	var numGoroutines atomic.Int32
	var progress atomic.Int32

	semaphoreChannel := make(chan struct{}, maxGoroutines)
	defer close(semaphoreChannel)
	var wg sync.WaitGroup

	//
	// Main loop
	//

	for i := 0; i < len(frameNumbers); i += opts.PerFile {

		wg.Add(1)
		semaphoreChannel <- struct{}{}

		numGoroutines.Add(1)

		//
		// Main goroutine
		//

		go func(done chan struct{}, _wg *sync.WaitGroup, sheetNumber int, startIndex int) {

			vc, err := gocv.VideoCaptureFile(videoUri)
			checkError(err)

			// Cleanup
			defer func() {
				vc.Close()
				<-semaphoreChannel
				numGoroutines.Add(-1)
				_wg.Done()
			}()

			mats := getImageMats(vc, frameNumbers[startIndex:startIndex+opts.PerFile]...)
			frames := make([]image.Image, 0)

			for _, mat := range mats {

				// Convert mat

				convertedMat := gocv.NewMat()

				mat.ConvertTo(&convertedMat, gocv.MatTypeCV8UC1)

				gocv.Resize(convertedMat, &convertedMat, image.Pt(0, 0), opts.Scale, opts.Scale, gocv.InterpolationNearestNeighbor)

				frame, err := convertedMat.ToImage()
				if err != nil {
					panic(err)
				}

				mat.Close()
				convertedMat.Close()

				frames = append(frames, frame)

			}

			f := sheets.NewWorkbook(fmt.Sprintf("%s/%d.xlsx", outputDir, sheetNumber))

			for i, frame := range frames {
				s := f.AddSheet(fmt.Sprintf("%d", frameNumbers[startIndex+i]))
				s.AddImage(frame)

				if opts.DoLogging {
					fmt.Printf(
						"\rProgress: %d%% Frame: %3d/%3d Active goroutines: %2d/%2d...",
						percentage(int(progress.Load()), len(frameNumbers)),
						frameNumbers[progress.Load()],
						frameNumbers[len(frameNumbers)-1],
						numGoroutines.Load(),
						runtime.NumCPU(),
					)
				}

				progress.Add(1)
			}

			f.Save()

		}(semaphoreChannel, &wg, (i+opts.PerFile)/opts.PerFile, i)
	}

	wg.Wait()

	if opts.DoLogging {
		fmt.Println("\nCleaning up...")
	}

}

func getImageMats(vc *gocv.VideoCapture, frames ...int) (mats []gocv.Mat) {

	mats = make([]gocv.Mat, len(frames))

	for i, v := range frames {

		mats[i] = gocv.NewMat()
		vc.Set(gocv.VideoCapturePosFrames, float64(v))
		vc.Read(&mats[i])

	}

	return
}
