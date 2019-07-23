package excel

import (
	"bufio"
	"fmt"
	"os"
	"sort"
	"strconv"

	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// WorkloadFile represents a workload file
type WorkloadFile struct {
	workbook    *excelize.File //*spreadsheet.Workbook
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

// Departments returns a list of all avaiable departments
func Departments() []string {
	return []string{
		string(ManagingDirectors), string(Consulting), string(Creation), string(Production), string(Text), string(Administration), string(Training), string(PR),
	}
}

// OpenWorkloadFile opens and returns a workloadfile
func OpenWorkloadFile(path string) *WorkloadFile {
	wb, err := excelize.OpenFile(path)

	if err != nil {
		fmt.Println(err)
	}

	sheetnames := []string{}
	finalRow := map[string]int{}
	finalColumn := map[string]string{}
	nextColumn := map[string]string{}
	for _, sheetname := range wb.GetSheetMap() {
		sheetnames = append(sheetnames, sheetname)
		rows, err := wb.GetRows(sheetname)
		if err != nil {
			fmt.Println(err)
			continue
		}
		finalRow[sheetname] = len(rows) - 1
		maxCol := 1
		for _, r := range rows {
			if len(r) > maxCol {
				maxCol = len(r)
			}
		}
		maxColName, _ := excelize.ColumnNumberToName(maxCol - 1)
		finalColumn[sheetname] = maxColName
		for idx, value := range rows[1] {
			if strings.TrimSpace(value) == "" {
				freeCol, _ := excelize.ColumnNumberToName(idx + 1)
				nextColumn[sheetname] = freeCol
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

// ModifiableSheetnames returns all sheetnames, that need to be modified
func (wf *WorkloadFile) ModifiableSheetnames() []string {
	return []string{
		"Kundenjobs", "Pitch_Neugeschäft", "Keine Arbeit", "Interne Jobs", "Urlaub", "Krankheit", "Feiertage", "Überstundenabbau",
	}
}
func (wf *WorkloadFile) nextAndFinalColNums(sheetname string) (next, final int) {
	next, err := excelize.ColumnNameToNumber(wf.nextColumn[sheetname])
	if err != nil {
		fmt.Println(err)
	}
	final, err = excelize.ColumnNameToNumber(wf.finalColumn[sheetname])
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
	names := strings.Split(employee, " ")
	employeeCoords, _ := wf.workbook.SearchSheet(sheetname, fmt.Sprintf("(%s).*(%s)|(%s).*(%s)", names[0], names[1], names[1], names[0]), true)
	if len(employeeCoords) < 1 {
		fmt.Printf("employee %s not found\n", employee)
		return
	} else if len(employeeCoords) > 1 {
		fmt.Printf("employee %s found multiple times\n", employee)
		return
	}
	_, row, _ := excelize.CellNameToCoordinates(employeeCoords[0])
	hoursCoords := fmt.Sprintf("%s%d", column, row)
	oldValue, err := wf.workbook.GetCellValue(sheetname, hoursCoords)
	hours := 0.0
	if err == nil {
		hours, _ = strconv.ParseFloat(oldValue, 64)
	}
	wf.workbook.SetCellFloat(sheetname, hoursCoords, value+hours, 2, 64)
}

// DeclareNewColumnForPeriod adds a new period into the next free column of sheetname
func (wf *WorkloadFile) DeclareNewColumnForPeriod(period string, sheetname string) string {
	next, final := wf.nextAndFinalColNums(sheetname)
	if next >= final {
		fmt.Println("reached end of file")
		return ""
	}
	nextName, _ := excelize.ColumnNumberToName(next)
	newNextName, _ := excelize.ColumnNumberToName(next + 1)
	wf.workbook.SetCellStr(sheetname, fmt.Sprintf("%s%d", nextName, 2), period)
	wf.nextColumn[sheetname] = newNextName
	return nextName
}

// DeclareNewColumnWithNextPeriod adds a new column to sheetname with a week more based on the last week
func (wf *WorkloadFile) DeclareNewColumnWithNextPeriod(sheetname string) string {
	next, _ := wf.nextAndFinalColNums(sheetname)

	lastPeriodColName, _ := excelize.ColumnNumberToName(next - 1)
	lastPeriod, err := wf.workbook.GetCellValue(sheetname, fmt.Sprintf("%s%d", lastPeriodColName, 2))
	if err != nil {
		fmt.Println(err)
		return ""
	}
	if !strings.Contains(lastPeriod, "-") {
		currentDate := time.Now()
		wf.DeclareNewColumnForPeriod(fmt.Sprintf("01.01-06.01.%s", currentDate.Format("06")), sheetname)
		return ""
	}
	dates := strings.Split(lastPeriod, "-")
	lastDate, err := time.Parse("02.01.06", dates[1])
	if err != nil {
		fmt.Println(err)
		return ""
	}
	newStartDate := lastDate.Add(time.Hour * 24)
	newEndDate := lastDate.Add(time.Hour * 24 * 7)
	newPeriod := fmt.Sprintf("%s-%s", newStartDate.Format("02.01"), newEndDate.Format("02.01.06"))
	return wf.DeclareNewColumnForPeriod(newPeriod, sheetname)

}

func (wf *WorkloadFile) removePeriodAtColumn(col string, sheetname string, rowExeptions []int) {
	for row := 2; row < wf.finalRow[sheetname]; row++ {
		isExepetion := false
		for _, exeption := range rowExeptions {
			if row == exeption {
				isExepetion = true
			}
		}
		formula, _ := wf.workbook.GetCellFormula(sheetname, fmt.Sprintf("%s%d", col, row))
		if strings.TrimSpace(formula) == "" && !isExepetion {
			wf.workbook.SetCellValue(sheetname, fmt.Sprintf("%s%d", col, row), "")
		}
	}
}

// RemoveLastPeriod removes the values of the most recent added column
func (wf *WorkloadFile) RemoveLastPeriod(sheetname string, rowExeptions []int) {
	nextCol, _ := wf.nextAndFinalColNums(sheetname)
	if nextCol < 2 {
		fmt.Println("no periods left to remove")
		return
	}
	removeColName, _ := excelize.ColumnNumberToName(nextCol - 1)
	wf.removePeriodAtColumn(removeColName, sheetname, rowExeptions)
	wf.nextColumn[sheetname] = removeColName
}

// AddEmployee adds a new Employee alphabetically and in the provided department
func (wf *WorkloadFile) AddEmployee(name string, department Department) {
	depCoords, err := wf.workbook.SearchSheet(wf.ModifiableSheetnames()[0], string(department))
	if err != nil {
		fmt.Println(err)
		return
	}
	depCol, depRow, err := excelize.CellNameToCoordinates(depCoords[0])
	if err != nil {
		fmt.Println(err)
		return
	}
	employees := []string{}
	for dr := depRow; dr > 2; dr-- {
		coords, _ := excelize.CoordinatesToCellName(depCol, dr)
		employee, err := wf.workbook.GetCellValue(wf.ModifiableSheetnames()[0], coords)
		if err != nil {
			fmt.Println(err)
			return
		}
		if strings.TrimSpace(employee) == "" {
			break
		}
		employees = append(employees, employee)
	}
	employees = append(employees, name)
	sort.Strings(employees)
	newEmployeeIndex := sort.SearchStrings(employees, name)
	newEmployeeRow := depRow - (len(employees) - newEmployeeIndex)
	updatedFormulas := wf.UpdateFormulas(wf.ModifiableSheetnames()[0], newEmployeeRow)
	err = wf.workbook.InsertRow(wf.ModifiableSheetnames()[0], newEmployeeRow)
	if err != nil {
		fmt.Println(err)
		return
	}
	err = wf.workbook.SetCellStr(wf.ModifiableSheetnames()[0], fmt.Sprintf("%s%d", "A", newEmployeeRow), name)
	if err != nil {
		fmt.Println(err)
		return
	}

	wf.Save("newEmployee.xlsx")
	fmt.Println("open the workloadfile, click 'delete' when asked and save the file again. Then press Enter")
	reader := bufio.NewReader(os.Stdin)
	reader.ReadString('\n')

	wf = OpenWorkloadFile("newEmployee.xlsx")
	for coords, formula := range updatedFormulas {
		err := wf.workbook.SetCellFormula(wf.ModifiableSheetnames()[0], coords, formula)
		if err != nil {
			fmt.Println(err)
			continue
		}
	}

}

// UpdateFormulas corrects formulas, that became incorrect by inserting a row
func (wf *WorkloadFile) UpdateFormulas(sheetname string, belowRow int) map[string]string {
	formulaMap := map[string]string{}
	formulaCoords := []string{}
	formulaParts := [][]string{}
	_, endCol := wf.nextAndFinalColNums(sheetname)
	for row := belowRow; row <= wf.finalRow[sheetname]; row++ {
		for col := 1; col <= endCol; col++ {
			coords, _ := excelize.CoordinatesToCellName(col, row)
			formula, err := wf.workbook.GetCellFormula(sheetname, coords)
			if err != nil {
				fmt.Println(err)
				continue
			}
			if formula != "" {
				parts := strings.FieldsFunc(formula, func(c rune) bool {
					return c == '(' || c == ';' || c == ':' || c == ')' || c == '+'
				})
				if len(parts) < 2 {
					continue
				}
				formulaCoords = append(formulaCoords, coords)
				formulaParts = append(formulaParts, parts)
			}
		}
	}

	for idx, parts := range formulaParts {

		switch parts[0] {
		case "SUM":
			// SUM formula e.g. SUM(B1:B3)
			startCoordsCol, startCoordsRow, _ := excelize.CellNameToCoordinates(parts[1])
			endCoordsCol, endCoordsRow, _ := excelize.CellNameToCoordinates(parts[2])

			// same col, different rows
			if startCoordsCol == endCoordsCol {
				name, err := wf.workbook.GetCellValue(sheetname, fmt.Sprintf("%s%d", "A", belowRow-1))
				if err != nil {
					fmt.Println(err)
					continue
				}
				if strings.TrimSpace(name) != "" {
					endCoordsRow++
				} else {
					startCoordsRow--
				}
				// different Cols, same row
			} else {
				startCoordsRow++
				endCoordsRow++
			}
			startColName, _ := excelize.ColumnNumberToName(startCoordsCol)
			endColName, _ := excelize.ColumnNumberToName(endCoordsCol)
			//err := wf.workbook.SetCellFormula(sheetname, formulaCoords[idx], fmt.Sprintf("SUM(%s%d:%s:%d)", startColName, startCoordsRow, endColName, endCoordsRow))
			formulaMap[formulaCoords[idx]] = fmt.Sprintf("SUM(%s%d:%s:%d)", startColName, startCoordsRow, endColName, endCoordsRow)
			// if err != nil {
			// 	fmt.Println(err)
			// 	continue
			// }
		default:
			// adding e.g. B1+B2+B3
			coords := []string{}
			for _, p := range parts {
				col, row, _ := excelize.CellNameToCoordinates(p)
				updatedCoords, _ := excelize.CoordinatesToCellName(col, row+1)
				coords = append(coords, updatedCoords)
			}
			formulaMap[formulaCoords[idx]] = strings.Join(coords, "+")
			// err := wf.workbook.SetCellFormula(sheetname, formulaCoords[idx], strings.Join(coords, "+"))
			// if err != nil {
			// 	fmt.Println(err)
			// 	continue
			// }
		}
	}
	return formulaMap
}

// Save saves the workloadfile to path
func (wf *WorkloadFile) Save(path string) {
	err := wf.workbook.SaveAs(path)
	if err != nil {
		fmt.Println(err)
	}
}
