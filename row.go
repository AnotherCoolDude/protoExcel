package protoexcel

import (
	"fmt"
	"reflect"
)

// RowProtoype represents a row in an Excel file
type RowProtoype []*CellPrototype

// R creates a new RowProtoype
func R(values ...interface{}) *RowProtoype {
	newRow := RowProtoype{}

	if len(values) == 0 {
		return &newRow
	}

	vv := values

	if reflect.TypeOf(values[0]).Kind() == reflect.Slice {
		//fmt.Println("values is a slice at values[0]")
		vv = values[0].([]interface{})
	}

	for _, value := range vv {
		newRow = append(newRow, &CellPrototype{Value: value})
	}
	return &newRow
}

// Helper functions

// addCoordinates adds coordinates to each Cell in RowProtoype
func (row *RowProtoype) addCoordinates(rowIndex int) {
	for columnIndex, cell := range *row {
		(*cell).coords = &coordinates{row: rowIndex, column: columnIndex + 1}
	}
}

// Euro formates values in row as euro
func (row *RowProtoype) Euro() *RowProtoype {
	for _, cell := range *row {
		formatted, _ := formatting(cell.Value)
		if !formatted {
			switch cell.Value.(type) {
			case int:
				(*cell).Value = Euro(float64(cell.Value.(int)))
			case float64:
				(*cell).Value = Euro(cell.Value.(float64))
			case float32:
				(*cell).Value = Euro(float64(cell.Value.(float32)))
			}
		}
	}
	return row
}

// AddBorderToRange adds a border to cells in the given range
func (row *RowProtoype) AddBorderToRange(border Border, colStart, colEnd int) *RowProtoype {
	if (colStart < 1 && colEnd < 1) || colStart > colEnd {
		fmt.Println("[row.go/AddBorderToRange: range invalid]")
	}
	for col, cell := range *row {
		if col+1 >= colStart && col+1 <= colEnd {
			(*cell).Border = border
		}
	}
	return row
}

// AddBorder adds a border to every cell in row
func (row *RowProtoype) AddBorder(border Border) *RowProtoype {
	for _, cell := range *row {
		(*cell).Border = border
	}
	return row
}
