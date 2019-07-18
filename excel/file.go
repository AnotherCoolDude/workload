package excel

import (
	"fmt"
	"regexp"
	"sort"
	"strconv"
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
	workbook    *spreadsheet.Workbook
	sheets      []string
	nextColumn  map[string]string
	finalColumn map[string]string
	finalRow    map[string]int
}

// Department describes the departments employees are categorised in
type Department string

const (
	// ManagingDirectors Department
	ManagingDirectors Department = "Geschäftsführung"
	// Consulting Department
	Consulting Department = "Beratung"
	//Creation Department
	Creation Department = "Kreative"
	//Production Department
	Production Department = "Produktion"
	//Text Department
	Text Department = "Text"
	//Administration Department
	Administration Department = "Verwaltung"
	//Training Department
	Training Department = "Auszubildende/Trainee"
	//PR Department
	PR Department = "PR"
)

func departments() []string {
	return []string{
		string(ManagingDirectors), string(Consulting), string(Creation), string(Production), string(Text), string(Administration), string(Training), string(PR),
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
	nextColumn := map[string]string{}
	finalColumn := map[string]string{}
	finalRow := map[string]int{}
	regexChars := regexp.MustCompile("[a-zA-Z]*")
	regexNums := regexp.MustCompile("[0-9]*")
	for _, sh := range sheets {

		sheetnames = append(sheetnames, sh.Name())
		latestCell := strings.Split(sh.Extents(), ":")[1]
		latestCol := regexChars.FindStringSubmatch(latestCell)[0]
		numsresult := regexNums.FindAllString(latestCell, 10)
		for _, result := range numsresult {
			if strings.TrimSpace(result) != "" {
				row, err := strconv.Atoi(result)
				if err != nil {
					fmt.Println(err)
				}
				finalRow[sh.Name()] = row
			}
		}
		finalColumn[sh.Name()] = latestCol

		colNum, _ := pe.ColumnNameToNumber(latestCol)
		for col := 2; col < colNum; col++ {
			colName, _ := pe.ColumnNumberToName(col)
			cell := sh.Cell(fmt.Sprintf("%s%d", colName, 2))
			if cell.IsEmpty() {
				nextColumn[sh.Name()] = colName
				break
			}
		}
	}

	return &WorkloadFile{
		workbook:    wb,
		sheets:      sheetnames,
		nextColumn:  nextColumn,
		finalColumn: finalColumn,
		finalRow:    finalRow,
	}

}

func (wf *WorkloadFile) nextAndFinalColNums(sheetname string) (next, final int) {
	next, err := pe.ColumnNameToNumber(wf.nextColumn[sheetname])
	if err != nil {
		fmt.Println(err)
	}
	final, err = pe.ColumnNameToNumber(wf.finalColumn[sheetname])
	if err != nil {
		fmt.Println(err)
	}
	return
}

// Sheetnames returns all sheetnames of the workloadfile
func (wf *WorkloadFile) Sheetnames() []string {
	return wf.sheets
}

// AddValueToEmployee adds a value to employee in the last used column of sheet
func (wf *WorkloadFile) AddValueToEmployee(employee string, value float64, sheetname, column string) {
	sheet, err := wf.workbook.GetSheet(sheetname)
	if err != nil {
		fmt.Println(err)
		return
	}
	employeeRow := 0
	for row := 1; row < wf.finalRow[sheetname]; row++ {
		cell := sheet.Cell(fmt.Sprintf("%s%d", "A", row))
		if cell.GetString() == employee {
			employeeRow = row
		}
	}
	if employeeRow == 0 {
		fmt.Printf("couldn't find employee %s\n", employee)
		return
	}
	sheet.Cell(fmt.Sprintf("%s%d", column, employeeRow)).SetNumber(value)
}

// InsertEmployee inserts a new Employee into the workloadfile
func (wf *WorkloadFile) InsertEmployee(name string, department Department) {
	for _, sheetname := range wf.sheets {
		sh, err := wf.workbook.GetSheet(sheetname)
		if err != nil {
			fmt.Printf("couldn't get sheet with name: %s\n", sheetname)
			continue
		}
		departmentrows := []int{}
	rowLoop:
		for row := 1; row < wf.finalRow[sheetname]; row++ {
			cellValue := sh.Cell(fmt.Sprintf("%s%d", "A", row)).GetString()
			if cellValue == name {
				fmt.Printf("employee %s allready exists\n", name)
				return
			}
			for _, dep := range departments() {
				if dep == cellValue {
					departmentrows = append(departmentrows, row)
					if cellValue == string(department) {
						break rowLoop
					}
				}
			}
		}
		depStart := 0
		depEnd := departmentrows[len(departmentrows)-1] - 1
		names := []string{}
		if len(departmentrows) == 1 {
			depStart = 8
		} else {
			depStart = departmentrows[len(departmentrows)-2] + 2
		}
		for r := depStart; r <= depEnd; r++ {
			cellValue := sh.Cell(fmt.Sprintf("%s%d", "A", r)).GetString()
			fmt.Print(cellValue)
			if cellValue != "" {
				names = append(names, cellValue)
			}
		}
		newEmpIndex := 0
		namesExt := names
		namesExt = append(namesExt, name)
		sort.Strings(names)
		sort.Strings(namesExt)
		for i, n := range names {
			if n != namesExt[i] {
				newEmpIndex = i
			}
		}
		if newEmpIndex == 0 {
			newEmpIndex = len(namesExt) - 1
		}
		fmt.Println(newEmpIndex)
		newRow := sh.InsertRow(depStart + newEmpIndex)
		newRow.Cell("A").SetString(name)

		// rowLoop:
		// 	for row := 1; row < wf.finalRow[sheetname]; row++ {
		// 		if sh.Cell(fmt.Sprintf("%s%d", "A", row)).GetString() == string(department) {
		// 			for existingRows := row; existingRows > 0; existingRows-- {
		// 				str := sh.Cell(fmt.Sprintf("%s%d", "A", existingRows)).GetString()
		// 				if strings.TrimSpace(str) != "" {
		// 					existingEmployees[existingRows] = str
		// 				} else {
		// 					break rowLoop
		// 				}
		// 			}
		// 		}
		// 	}
		// 	newEmployeeRow := calcNewRow(name, existingEmployees)
		// 	newRow := sh.InsertRow(newEmployeeRow)
		// 	newRow.Cell("A").SetString(name)

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
	next, final := wf.nextAndFinalColNums(sheetname)
	sheet, err := wf.workbook.GetSheet(sheetname)
	if err != nil {
		fmt.Println(err)
		return
	} else if next >= final {
		fmt.Println("reached end of file")
		return
	}
	nextName, _ := pe.ColumnNumberToName(next)
	newNextName, _ := pe.ColumnNumberToName(next + 1)
	sheet.Cell(fmt.Sprintf("%s%d", nextName, 2)).SetString(period)
	wf.nextColumn[sheetname] = newNextName
}

// DeclareNewColumnWithNextPeriod adds a new column to sheetname with a week more based on the last week
func (wf *WorkloadFile) DeclareNewColumnWithNextPeriod(sheetname string) {
	next, _ := wf.nextAndFinalColNums(sheetname)
	sheet, err := wf.workbook.GetSheet(sheetname)
	if err != nil {
		fmt.Println(err)
		return
	}
	lastPeriodColName, _ := pe.ColumnNumberToName(next - 1)
	lastPeriod := sheet.Cell(fmt.Sprintf("%s%d", lastPeriodColName, 2)).GetString()
	if !strings.Contains(lastPeriod, "-") {
		currentDate := time.Now()
		wf.DeclareNewColumnForPeriod(fmt.Sprintf("01.01-06.01.%s", currentDate.Format("06")), sheetname)
		return
	}
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
	for row := 2; row < wf.finalRow[sheetname]; row++ {
		if !sheet.Cell(fmt.Sprintf("%s%d", col, row)).HasFormula() {
			sheet.Cell(fmt.Sprintf("%s%d", col, row)).Clear()
		}
	}
}

// RemoveLastPeriod removes the values of the most recent added column
func (wf *WorkloadFile) RemoveLastPeriod(sheetname string) {
	nextCol, _ := wf.nextAndFinalColNums(sheetname)
	if nextCol < 2 {
		fmt.Println("no periods left to remove")
		return
	}
	removeColName, _ := pe.ColumnNumberToName(nextCol - 1)
	wf.removePeriodAtColumn(removeColName, sheetname)
}

// Save saves the workloadfile to path
func (wf *WorkloadFile) Save(path string) {
	wf.workbook.RemoveCalcChain()
	err := wf.workbook.SaveToFile(path)
	if err != nil {
		fmt.Println(err)
	}
}

// Validate validates the workloadfile
func (wf *WorkloadFile) Validate() {
	err := wf.workbook.Validate()
	if err != nil {
		fmt.Println(err)
	}
}
