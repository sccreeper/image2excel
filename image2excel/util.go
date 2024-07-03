package image2excel

func checkError(e error) {
	if e != nil {
		panic(e)
	}
}

// This only exists because doing it inline is ugly
func percentage(numerator int, denominator int) int {

	return int((float64(numerator) / float64(denominator)) * 100.0)

}

func clamp(val int, min int, max int) int {

	if val < min {
		return min
	} else if val > max {
		return max
	} else {
		return val
	}

}
