package excel

import (
	"github.com/AnotherCoolDude/protoexcel"
)

// Open opens the excel file at path for read only access
func Open(path string) *protoexcel.Read {
	return protoexcel.ReadExcel(path, true)
}

// FilterColumns extracts cols and returns a map with values per col
func FilterColumns(cols []int, read *protoexcel.Read) map[int][]interface{} {
	colMap := map[int][]interface{}{}
	for _, col := range cols {
		extractedCol := read.Column(read.Sheets()[1], col)
		values := []interface{}{}
		for _, cell := range extractedCol {
			values = append(values, cell.Value)
		}
		colMap[col] = values
	}
	return colMap
}
