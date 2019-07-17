package excel

import (
	"fmt"
	"regexp"
	"sort"
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

// Department describes the departments employees are categorised in
type Department string

const (
	managingDirectors Department = "Geschäftsführung"
	consulting        Department = "Beratung"
	creation          Department = "Kreative"
	production        Department = "Produktion"
	text              Department = "Text"
	administration    Department = "Verwaltung"
	training          Department = "Auszubildende/Trainee"
	pr                Department = "PR"
)

func departments() []string {
	return []string{
		string(managingDirectors), string(consulting), string(creation), string(production), string(text), string(administration), string(training), string(pr),
	}
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

// InsertEmployee inserts a new Employee into the workloadfile
func (wf *WorkloadFile) InsertEmployee(name string, department Department) {
	for _, sheetname := range wf.sheets {
		sh, err := wf.workbook.GetSheet(sheetname)
		if err != nil {
			fmt.Printf("couldn't get sheet with name: %s\n", sheetname)
			continue
		}
		existingEmployees := map[int]string{}
	rowLoop:
		for row := 1; row < 100; row++ {
			if sh.Cell(fmt.Sprintf("%s%d", "A", row)).GetString() == string(department) {
				for existingRows := row; existingRows > 0; existingRows-- {
					str := sh.Cell(fmt.Sprintf("%s%d", "A", existingRows)).GetString()
					if strings.TrimSpace(str) != "" {
						existingEmployees[existingRows] = str
					} else {
						break rowLoop
					}
				}
			}
		}
		newEmployeeRow := calcNewRow(name, existingEmployees)
		newRow := sh.InsertRow(newEmployeeRow)
		newRow.Cell("A").SetString(name)
	}
}

func calcNewRow(newEmployee string, existingEmployeesMap map[int]string) int {
	names := []string{}
	lowest := 0
	highest := 0
	for key, value := range existingEmployeesMap {
		if lowest == 0 {
			lowest = key
		}
		if key > highest {
			highest = key
		}
		if key < lowest {
			lowest = key
		}
		names = append(names, value)
	}
	names = append(names, newEmployee)
	sort.Strings(names)
	newRow := 0
	count := 0
	for l := lowest; l < highest+1; l++ {
		if names[count] == newEmployee {
			newRow = l
		}
		count++
	}
	return newRow
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

func (wf *WorkloadFile) removePeriodAtColumn(col string, sheetname string) {
	sheet, err := wf.workbook.GetSheet(sheetname)
	if err != nil {
		fmt.Println(err)
		return
	}
	for row := 2; row < 100; row++ {
		if !sheet.Cell(fmt.Sprintf("%s%d", col, row)).HasFormula() {
			sheet.Cell(fmt.Sprintf("%s%d", col, row)).SetString("")
		}
	}
}

// RemoveLastPeriod removes the values of the most recent added column
func (wf *WorkloadFile) RemoveLastPeriod(sheetname string) {
	wf.removePeriodAtColumn(wf.latestColumns[sheetname], sheetname)
}
