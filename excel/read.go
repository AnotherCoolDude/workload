package excel

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
	pe "github.com/AnotherCoolDude/protoexcel"
)

// Open opens the excel file at path for read only access
func Open(path string) *pe.Read {
	return pe.ReadExcel(path, true)
}

// FilterColumns extracts cols and returns a map with values per col
func FilterColumns(cols []int, read *pe.Read) map[int][]interface{} {
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

// TimetrackingFile wraps the time tracking excel file from proad in a struct
type TimetrackingFile struct {
	file      *excelize.File
	length    int
	sheetname string
}

// ReadProadExcel opens a proad export, that has been saved as excel
func ReadProadExcel(path string) *TimetrackingFile {
	file, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		return &TimetrackingFile{}
	}
	sheetname := file.GetSheetName(file.GetActiveSheetIndex())
	rows, _ := file.GetRows(sheetname)
	length := len(rows)
	return &TimetrackingFile{
		file:      file,
		length:    length,
		sheetname: sheetname,
	}
}

// GetColumns returns a map with columns as key and cellValues as value array
func (ttf *TimetrackingFile) GetColumns(columns []int) map[int][]string {
	filteredValues := map[int][]string{}
	for _, col := range columns {
		colName, err := excelize.ColumnNumberToName(col)
		if err != nil {
			fmt.Println(err)
			continue
		}
		// row = 2 to evade first row with column titles
		for row := 2; row <= ttf.length; row++ {
			value, err := ttf.file.GetCellValue(ttf.sheetname, fmt.Sprintf("%s%d", colName, row))
			if err != nil {
				fmt.Println(err)
				continue
			}
			filteredValues[col] = append(filteredValues[col], value)
		}
	}
	return filteredValues
}
