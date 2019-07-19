package excel

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"

	"strings"
	"time"
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

func departments() []string {
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

func (wf *WorkloadFile) removePeriodAtColumn(col string, sheetname string) {
	for row := 2; row < wf.finalRow[sheetname]; row++ {
		formula, _ := wf.workbook.GetCellFormula(sheetname, fmt.Sprintf("%s%d", col, row))
		if strings.TrimSpace(formula) == "" {
			wf.workbook.SetCellValue(sheetname, fmt.Sprintf("%s%d", col, row), "")
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
	removeColName, _ := excelize.ColumnNumberToName(nextCol - 1)
	wf.removePeriodAtColumn(removeColName, sheetname)
	wf.nextColumn[sheetname] = removeColName
}

// Save saves the workloadfile to path
func (wf *WorkloadFile) Save(path string) {
	err := wf.workbook.SaveAs(path)
	if err != nil {
		fmt.Println(err)
	}
}
