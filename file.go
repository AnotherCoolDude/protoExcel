package protoexcel

import (
	"fmt"
	"time"

	"github.com/unidoc/unioffice/color"
	"github.com/unidoc/unioffice/schema/soo/sml"
	"github.com/unidoc/unioffice/spreadsheet"
)

func createFile(prototypes []*SheetPrototype, path string) {
	ss := spreadsheet.New()

	for _, prototype := range prototypes {
		sheet := ss.AddSheet()
		sheet.SetName(prototype.name)
		prototype.sortDraft(func(row *RowProtoype) {
			for _, cell := range *row {
				cell.addToExcelSheet(prototype.name, ss)
			}
		})
	}

	err := ss.SaveToFile(path)
	if err != nil {
		fmt.Println(err)
	}
}

// addValue adds the value of cell to given speadsheet.CellPrototype
func (cell *CellPrototype) addValue(excelCell *spreadsheet.Cell) {
	switch v := cell.Value.(type) {
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64:
		(*excelCell).SetNumber(v.(float64))
	case float32:
		(*excelCell).SetNumber(float64(v))
	case float64:
		(*excelCell).SetNumber(v)
	case string:
		(*excelCell).SetString(v)
	case time.Time:
		(*excelCell).SetDateWithStyle(v)
	case bool:
		(*excelCell).SetBool(v)
	case nil:
		(*excelCell).SetString("")
	case Euro:
		(*excelCell).SetNumber(float64(v))
	case Formula:
		(*excelCell).SetFormulaRaw(string(v))
	default:
		(*excelCell).SetString(v.(string))
	}
}

// addStyle adds the style of cell to given spreadsheet.CellPrototype using spreadsheet.Stylesheet
func (cell *CellPrototype) addStyle(excelCell *spreadsheet.Cell, styleSheet *spreadsheet.StyleSheet) {
	var cellStyle spreadsheet.CellStyle
	var borderStyle spreadsheet.Border

	//formats
	_, f := formatting(cell.Value)
	if f == "" && cell.Border == None {
		return
	}

	cellStyle = styleSheet.AddCellStyle()
	cellStyle.SetNumberFormat(f)

	//border
	if cell.Border == None {
		(*excelCell).SetStyle(cellStyle)
		return
	}

	borderStyle = styleSheet.AddBorder()
	cellStyle.SetBorder(borderStyle)

	switch cell.Border {
	case Top:
		borderStyle.SetTop(sml.ST_BorderStyleThin, color.Black)
	case Bottom:
		borderStyle.SetBottom(sml.ST_BorderStyleThin, color.Black)
	case DoubleBottom:
		borderStyle.SetBottom(sml.ST_BorderStyleDouble, color.Black)
	case Left:
		borderStyle.SetLeft(sml.ST_BorderStyleThin, color.Black)
	case Right:
		borderStyle.SetRight(sml.ST_BorderStyleThin, color.Black)
	default:
	}

	(*excelCell).SetStyle(cellStyle)
}

// addToExcelSheet creates a spreadsheet.CellPrototype based on cell and adds it to sheet in woekbook
func (cell *CellPrototype) addToExcelSheet(sheet string, workbook *spreadsheet.Workbook) {
	if *cell.coords == (coordinates{}) {
		fmt.Printf("[file.go/addToSheet] coordinates of cell %+v are empty\n", *cell)
		return
	}
	excelSheet, err := workbook.GetSheet(sheet)
	if err != nil {
		fmt.Printf("[file.go/addToSheet] sheetname %s invalid\n", sheet)
		return
	}
	excelCell := excelSheet.Cell(cell.CoordsToString())
	cell.addValue(&excelCell)
	cell.addStyle(&excelCell, &workbook.StyleSheet)

}
