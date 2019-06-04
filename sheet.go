package protoexcel

import (
	"fmt"
	"sort"
)

// func main() {
// 	proto := NewSheetPrototype("new Sheet", "Col 1", "Col 2", "Col 3", "Col 4", "Col 5")
// 	proto.AppendEmptyRow()
// 	proto.AppendRow(R("Value 1", "Value 2", 1224, 11.33, F("C3+D3")).Euro().AddBorderToRange(DoubleBottom, 2, 4))
// 	proto.printDraft(true)
// 	createFile([]*SheetPrototype{proto}, "/Users/christianhovenbitzer/Desktop/testExcel.xlsx")
// }

// SheetPrototype is a prototype for a excel file
type SheetPrototype struct {
	name  string
	draft map[int]*RowProtoype
}

// NewSheetPrototype creates and returns a new ExcelProtoype instance
func NewSheetPrototype(name string, headerNames ...interface{}) *SheetPrototype {
	proto := &SheetPrototype{
		name:  name,
		draft: map[int]*RowProtoype{},
	}
	headervalues := make([]interface{}, len(headerNames))
	for i := range headerNames {
		headervalues[i] = headerNames[i]
	}
	row := R(headerNames)
	row.addCoordinates(1)
	proto.draft[1] = row
	return proto
}

// WritePrototypeToFile writes the provided sheets to path
func WritePrototypeToFile(sheetPrototypes []*SheetPrototype, path string) {
	createFile(sheetPrototypes, path)
}

// AppendRow appends a Row to draft
func (prototype *SheetPrototype) AppendRow(row *RowProtoype) {
	cr := prototype.currentRow() + 1
	row.addCoordinates(cr)
	prototype.draft[cr] = row
}

// AppendEmptyRow appends a empty Row to draft
func (prototype *SheetPrototype) AppendEmptyRow() {
	emptyRow := R("")
	prototype.AppendRow(emptyRow)
}

// HeaderColumns returns the header columns of draft
func (prototype *SheetPrototype) HeaderColumns() []string {
	firstRow := prototype.draft[1]
	values := []string{}
	for _, cell := range *firstRow {
		values = append(values, cell.Value.(string))
	}
	return values
}

// Helper functions

// currentRow returns the index of the most recent added row
func (prototype *SheetPrototype) currentRow() int {
	currentRow := 1
	for rowIndex := range prototype.draft {
		if rowIndex > currentRow {
			currentRow = rowIndex
		}
	}
	return currentRow
}

// PrintDraft print values of draft
func (prototype *SheetPrototype) PrintDraft(verbose bool) {
	prototype.sortDraft(func(row *RowProtoype) {
		for _, cell := range *row {
			if verbose {
				fmt.Printf("[v: %s, r: %d c: %d b: %d]\t", cell.Value, cell.coords.row, cell.coords.column, int(cell.Border))
			} else {
				fmt.Printf("%s\t", cell.Value)
			}
		}
		fmt.Println()
	})
}

func (prototype *SheetPrototype) sortDraft(fn func(*RowProtoype)) {
	keys := []int{}
	for k := range prototype.draft {
		keys = append(keys, k)
	}
	sort.Ints(keys)
	for _, k := range keys {
		fn(prototype.draft[k])
	}
}
