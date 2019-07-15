package protoexcel

import (
	"fmt"

	pb "gopkg.in/cheggaaa/pb.v2"

	"github.com/unidoc/unioffice/spreadsheet"
)

// Read wraps a spreadsheet.Workbook
type Read struct {
	workbook   *spreadsheet.Workbook
	cachedRows map[string][]*RowProtoype
	verbose    bool
}

// ReadExcel opens a excelfile at path for Read-Only access
func ReadExcel(path string, verbose bool) *Read {
	wb := openFile(path)
	return &Read{
		workbook:   wb,
		cachedRows: map[string][]*RowProtoype{},
		verbose:    verbose,
	}
}

// Sheets returns the names of all sheets
func (r *Read) Sheets() []string {
	names := make([]string, len(r.workbook.Sheets()))
	for _, sh := range r.workbook.Sheets() {
		names = append(names, sh.Name())
	}
	return names
}

// Rows returns all Rows of Sheet with sheetname
func (r *Read) Rows(sheetname string) []*RowProtoype {
	if cached, ok := r.cachedRows[sheetname]; ok {
		return cached
	}

	sheet, err := r.workbook.GetSheet(sheetname)
	rows := []*RowProtoype{}
	if err != nil {
		fmt.Printf("couldn't get sheet with name %s: %s\n", sheetname, err)
		return rows
	}

	var count int
	var bar *pb.ProgressBar

	if r.verbose {
		count = len(sheet.Rows())
		fmt.Printf("amount of rows in sheet %s: %d\n", sheetname, count)
		bar = pb.StartNew(count)
	}

	for _, row := range sheet.Rows() {

		values := []interface{}{}

		lastCellIdx := len(row.Cells())
		lastCell := row.Cells()[lastCellIdx-1]
		lastCellCol, _ := lastCell.Column()
		colNum, _ := columnNameToNumber(lastCellCol)
		cells := row.Cells()

		found := false
		for i := 1; i <= colNum; i++ {

			for _, cell := range cells {
				col, _ := cell.Column()
				num, _ := columnNameToNumber(col)
				if num == i {
					switch {
					case cell.IsBool():
						v, _ := cell.GetValueAsBool()
						values = append(values, v)
					case cell.IsNumber():
						v, _ := cell.GetValueAsNumber()
						values = append(values, v)
					case cell.IsEmpty():
						values = append(values, "")
					default:
						v, _ := cell.GetRawValue()
						values = append(values, v)
					}
					found = true
				}
			}
			if !found {
				values = append(values, " ")
			}
			found = false

		}
		if r.verbose {
			bar.Increment()
		}
		rows = append(rows, R(values))

		// for _, cell := range row.Cells() {
		// 	switch {
		// 	case cell.IsBool():
		// 		v, _ := cell.GetValueAsBool()
		// 		values = append(values, v)
		// 	case cell.IsNumber():
		// 		v, _ := cell.GetValueAsNumber()
		// 		values = append(values, v)
		// 	case cell.IsEmpty():
		// 		values = append(values, "")
		// 	default:
		// 		v, _ := cell.GetRawValue()
		// 		values = append(values, v)
		// 	}
		// }
		// if r.verbose {
		// 	bar.Increment()
		// }
		// rows = append(rows, R(values))
	}
	if r.verbose {
		bar.Finish()
	}
	r.cachedRows[sheetname] = rows
	return rows
}

// Column returns all values from sheet sheetname and column col
func (r *Read) Column(sheetname string, col int) []*CellPrototype {
	if col < 1 {
		panic("col has to be greater than 0")
	}
	cells := []*CellPrototype{}
	rows := []*RowProtoype{}
	if cached, ok := r.cachedRows[sheetname]; ok {
		rows = cached
	} else {
		rows = r.Rows(sheetname)
	}

	for _, row := range rows {
		if len(*row) < col {
			fmt.Printf("no cell at col %d found, inserting empty cell \n", col)
			cells = append(cells, &CellPrototype{Value: "", Border: None})
		}
		for idx, c := range *row {
			if idx == col-1 {
				cells = append(cells, c)
			}
		}
	}
	return cells
}

// helper

// openFile opens the file at path and returns a *spreadsheet.Workbook
func openFile(path string) *spreadsheet.Workbook {
	wb, err := spreadsheet.Open(path)
	if err != nil {
		fmt.Printf("error opening file %s: %s\n", path, err)
		panic(err)
	}
	return wb
}
