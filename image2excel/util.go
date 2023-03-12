package image2excel

func check_error(e error) {
	if e != nil {
		panic(e)
	}
}
