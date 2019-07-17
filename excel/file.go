package excel

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/AnotherCoolDude/protoexcel"
	"github.com/unidoc/unioffice/spreadsheet"
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

func writeToSheet(sheetname string, excelfile *spreadsheet.Workbook) {
	sheet, err := excelfile.GetSheet(sheetname)
	if err != nil {
		fmt.Println(err)
	}

}

// WorkloadFile represents a workload file
type WorkloadFile struct {
	workbook      *spreadsheet.Workbook
	sheets        []string
	latestColumns map[string]string
}

// OpenWorkloadFile opens and returns a workloadfile
func OpenWorkloadFile(path string) *WorkloadFile {
	wb, err := spreadsheet.Open(path)

	if err != nil {
		fmt.Println(err)
	}

	sheets := wb.Sheets()
	sheetnames := []string{}
	latestColumns := map[string]string{}

	regex := regexp.MustCompile("[a-zA-Z]*")

	for _, sh := range sheets {
		sheetnames = append(sheetnames, sh.Name())
		latestCol := strings.Split(sh.Extents(), ":")[1]
		latestColumns[sh.Name()] = regex.FindStringSubmatch(latestCol)[0]

	}

	return &WorkloadFile{
		workbook:      wb,
		sheets:        sheetnames,
		latestColumns: latestColumns,
	}

}

// CurrentColumn returns the current column of the sheet with sheetname, e.g. "B"
func (wf *WorkloadFile) CurrentColumn(sheetname string) int {
	extend := wf.latestColumns[sheetname]
	coords := strings.Split(extend, ":")
	num, err := protoexcel.ColumnNameToNumber(string(coords[1][0]))
	if err != nil {
		fmt.Println(err)
	}
	return num
}

// Sheetnames returns all sheetnames of the workloadfile
func (wf *WorkloadFile) Sheetnames() []string {
	return wf.sheets
}

func (wf *WorkloadFile) AddValuesToNextColumn(values map[string]float32, sheetname string) {
	nextCol := wf.CurrentColumn(sheetname) + 1

}
