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
var main_page string

const (
	max_file_size = 8000000 //8 megabytes
)

func main() {

	r := gin.Default()

	r.Static("/public", "public")

	r.GET("/", func(ctx *gin.Context) {
		ctx.Data(http.StatusOK, "text/html; charset=utf-8", []byte(main_page))
	})

	r.POST("/convert", func(ctx *gin.Context) {

		var file_bytes []byte
		var file_buffer bytes.Buffer

		scale_form, _ := ctx.GetPostForm("scale")
		scale, err := strconv.ParseFloat(scale_form, 64)
		if err != nil {
			ctx.AbortWithStatus(http.StatusInternalServerError)
		}

		fmt.Println(scale)

		file, header, err := ctx.Request.FormFile("file")
		if err != nil {
			ctx.AbortWithStatus(http.StatusInternalServerError)
		}

		if header.Size >= max_file_size {
			ctx.AbortWithStatus(http.StatusUnprocessableEntity)
		}

		file_bytes, err = ioutil.ReadAll(file)
		if err != nil {
			ctx.AbortWithStatus(http.StatusInternalServerError)
		}

		file_buffer.Write(file_bytes)

		result_buffer, err := i2e.Convert(&file_buffer, header.Filename, true, scale, 0, 0)
		if err != nil {
			ctx.AbortWithStatus(http.StatusInternalServerError)
		}

		ctx.Data(http.StatusOK, "application/octet-stream", result_buffer.Bytes())

	})

	r.Run()

}
