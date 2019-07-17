package excel

import (
	"fmt"
	"regexp"
	"strings"
	"time"

	pe "github.com/AnotherCoolDude/protoexcel"
	"github.com/unidoc/unioffice/spreadsheet"
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

	for _, sh := range sheets[:5] {

		sheetnames = append(sheetnames, sh.Name())
		latestCell := strings.Split(sh.Extents(), ":")[1]
		latestCol := regex.FindStringSubmatch(latestCell)[0]
		colNum, _ := pe.ColumnNameToNumber(latestCol)

		for col := 2; col < colNum; col++ {
			colName, _ := pe.ColumnNumberToName(col)
			cell := sh.Cell(fmt.Sprintf("%s%d", colName, 2))
			if value, err := cell.GetRawValue(); err == nil {
				if value == "" {
					lastUsedColName, _ := pe.ColumnNumberToName(col - 1)
					latestColumns[sh.Name()] = lastUsedColName
				}
			}
		}

	}

	return &WorkloadFile{
		workbook:      wb,
		sheets:        sheetnames,
		latestColumns: latestColumns,
	}

}

// Sheetnames returns all sheetnames of the workloadfile
func (wf *WorkloadFile) Sheetnames() []string {
	return wf.sheets
}

// AddValueToEmployee adds a value to employee in the last used column of sheet
func (wf *WorkloadFile) AddValueToEmployee(employee string, value float64, sheetname string) {
	lastUsedCol := wf.latestColumns[sheetname]
	sheet, err := wf.workbook.GetSheet(sheetname)
	if err != nil {
		fmt.Println(err)
		return
	}
	employeeRow := 0
	for row := 1; row < 100; row++ {
		cell := sheet.Cell(fmt.Sprintf("%s%d", "A", row))
		if name, err := cell.GetRawValue(); err == nil {
			if name == employee {
				employeeRow = row
			}
		}
	}
	if employeeRow == 0 {
		fmt.Printf("couldn't find employee %s\n", employee)
		return
	}
	sheet.Cell(fmt.Sprintf("%s%d", lastUsedCol, employeeRow)).SetNumber(value)
}

// DeclareNewColumnForPeriod adds a new period into the next free column of sheetname
func (wf *WorkloadFile) DeclareNewColumnForPeriod(period string, sheetname string) {
	lastUsedCol := wf.latestColumns[sheetname]
	sheet, err := wf.workbook.GetSheet(sheetname)
	if err != nil {
		fmt.Println(err)
		return
	}
	colNum, _ := pe.ColumnNameToNumber(lastUsedCol)
	nextColName, _ := pe.ColumnNumberToName(colNum + 1)
	sheet.Cell(fmt.Sprintf("%s%d", nextColName, 2)).SetString(period)
	wf.latestColumns[sheetname] = nextColName
}

// DeclareNewColumnWithNextPeriod adds a new column to sheetname with a week more based on the last week
func (wf *WorkloadFile) DeclareNewColumnWithNextPeriod(sheetname string) {
	lastUsedCol := wf.latestColumns[sheetname]
	sheet, err := wf.workbook.GetSheet(sheetname)
	if err != nil {
		fmt.Sprintln(err)
		return
	}
	lastPeriod := sheet.Cell(fmt.Sprintf("%s%d", lastUsedCol, 2)).GetString()
	dates := strings.Split(lastPeriod, "-")
	lastDate, err := time.Parse("02.01.06", dates[1])
	if err != nil {
		fmt.Println(err)
		return
	}
	newStartDate := lastDate.Add(time.Hour * 24)
	newEndDate := lastDate.Add(time.Hour * 24 * 7)
	newPeriod := fmt.Sprintf("%s-%s", newStartDate.Format("02.01"), newEndDate.Format("02.01.06"))
	wf.DeclareNewColumnForPeriod(newPeriod, sheetname)

}
