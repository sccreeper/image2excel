package main

import (
	"bytes"
	"fmt"
	"io/ioutil"
	"net/http"
	i2e "sccreeper/image2excel/image2excel"
	"strconv"

	"github.com/gin-gonic/gin"

	_ "embed"
)

//go:embed index.html
var mainPage string

const (
	maxFileSize = 8000000 //8 megabytes
)

func main() {

	r := gin.Default()

	r.Static("/public", "public")

	r.GET("/", func(ctx *gin.Context) {
		ctx.Data(http.StatusOK, "text/html; charset=utf-8", []byte(mainPage))
	})

	r.POST("/convert", func(ctx *gin.Context) {

		var fileBytes []byte
		var fileBuffer bytes.Buffer

		scaleForm, _ := ctx.GetPostForm("scale")
		scale, err := strconv.ParseFloat(scaleForm, 64)
		if err != nil {
			ctx.AbortWithStatus(http.StatusInternalServerError)
		}

		fmt.Println(scale)

		file, header, err := ctx.Request.FormFile("file")
		if err != nil {
			ctx.AbortWithStatus(http.StatusInternalServerError)
		}

		if header.Size >= maxFileSize {
			ctx.AbortWithStatus(http.StatusUnprocessableEntity)
		}

		fileBytes, err = ioutil.ReadAll(file)
		if err != nil {
			ctx.AbortWithStatus(http.StatusInternalServerError)
		}

		fileBuffer.Write(fileBytes)

		resultBuffer, err := i2e.Convert(&fileBuffer, header.Filename, true, scale, 0, 0)
		if err != nil {
			ctx.AbortWithStatus(http.StatusInternalServerError)
		}

		ctx.Data(http.StatusOK, "application/octet-stream", resultBuffer.Bytes())

	})

	r.Run()

}
