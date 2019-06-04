package protoexcel

// Euro represents a float64, that will be presented as Euro in excel
type Euro float64

// Formula represents a string, that will be presented as a Formula in excel
type Formula string

func formatting(format interface{}) (isFormatted bool, formatStr string) {
	switch format.(type) {
	case Euro:
		return true, "_-* #.##0,00 €_-;-* #.##0,00 €_-;_-* \"-\"?? €_-;_-@_-"
	case Formula:
		return true, ""
	default:
		return false, ""
	}
}
