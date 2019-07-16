package protoexcel

import (
	"errors"
	"fmt"
	"strconv"
	"strings"
)

// CellPrototype represents a cell in an Excel file
type CellPrototype struct {
	Value  interface{}
	Border Border
	coords *coordinates
}

// Create Cells

// F returns a Cell with formula
func F(formula string) Formula {
	return Formula(formula)
}

// EUR returns a Cell with amount formatted as Euro
func EUR(amount float64) Euro {
	return Euro(amount)
}

// Handle coordinates

// Coordiantes represents the current coordinates of CellPrototype
type coordinates struct {
	row, column int
}

// CoordsToString returns the coords of cell as excel-like string
func (cell *CellPrototype) CoordsToString() string {
	return cell.coords.coordsToString()
}

func (coords *coordinates) coordsToString() string {
	colStr, err := ColumnNumberToName(coords.column)
	if err != nil {
		fmt.Println(err)
	}
	return fmt.Sprintf("%s%d", colStr, coords.row)
}

// coordinatesFromString returns a coordinates struct from coordString
func coordinatesFromString(coordString string) *coordinates {
	colStr, row, err := splitCellName(coordString)
	if err != nil {
		fmt.Println(err)
	}
	col, err := ColumnNameToNumber(colStr)
	if err != nil {
		fmt.Println(err)
	}
	return &coordinates{
		column: col,
		row:    row,
	}
}

//helper

// ColumnNameToNumber provides a function to convert Excel sheet column name
// to int. Column name case insensitive. The function returns an error if
// column name incorrect.
func ColumnNameToNumber(name string) (int, error) {
	if len(name) == 0 {
		return -1, errors.New("[cell.go/columnNameToNumber] received empty string")
	}
	col := 0
	multi := 1
	for i := len(name) - 1; i >= 0; i-- {
		r := name[i]
		if r >= 'A' && r <= 'Z' {
			col += int(r-'A'+1) * multi
		} else if r >= 'a' && r <= 'z' {
			col += int(r-'a'+1) * multi
		} else {
			return -1, errors.New("[cell.go/columnNameToNumber] invalid string")
		}
		multi *= 26
	}
	return col, nil
}

// ColumnNumberToName provides a function to convert the integer to Excel
// sheet column title.
func ColumnNumberToName(num int) (string, error) {
	if num < 1 {
		return "", fmt.Errorf("[cell.go/columnNumberToName] incorrect column number %d", num)
	}
	var col string
	for num > 0 {
		col = string((num-1)%26+65) + col
		num = (num - 1) / 26
	}
	return col, nil
}

// SplitCellName splits cell name to column name and row number.

func splitCellName(cell string) (string, int, error) {
	alpha := func(r rune) bool {
		return ('A' <= r && r <= 'Z') || ('a' <= r && r <= 'z')
	}

	if strings.IndexFunc(cell, alpha) == 0 {
		i := strings.LastIndexFunc(cell, alpha)
		if i >= 0 && i < len(cell)-1 {
			col, rowstr := cell[:i+1], cell[i+1:]
			if row, err := strconv.Atoi(rowstr); err == nil && row > 0 {
				return col, row, nil
			}
		}
	}
	return "", -1, errors.New("[cell.go/splitCellName] could not split string")
}
