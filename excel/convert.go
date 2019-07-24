package excel

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// ConvertCSV converts a csv file to an excel file
func ConvertCSV(path string, verbose bool) *excelize.File {
	csvFile, err := os.Open(path)
	if err != nil {
		fmt.Println(err)
		os.Exit(0)
	}
	defer csvFile.Close()

	r := bufio.NewReader(csvFile)
	var str string
	var correctedCSV string
	lines := []string{}

	for {
		str, err = r.ReadString('\r')
		if err != nil {
			break
		}
		str = strings.TrimLeft(str, "\n")

		if str[0] != '"' && len(lines) > 0 {
			previousStr := lines[len(lines)-1]
			str = previousStr[:len(previousStr)-1] + " " + str
			lines = lines[:len(lines)-1]
		}
		splitted := strings.Split(str, ";")
		parts := []string{}

		for _, s := range splitted {
			trimmed := strings.TrimFunc(s, func(r rune) bool {
				if r == '\r' || r == '\n' {
					return true
				}
				return false
			})
			correctedQuotes := strings.ReplaceAll(trimmed[1:len(trimmed)-1], "\"", "\"\"")
			if trimmed[len(trimmed)-1] != '"' {
				correctedQuotes = correctedQuotes + string(trimmed[len(trimmed)-1])
			}
			parts = append(parts, "\""+correctedQuotes+"\"")
		}
		lines = append(lines, strings.Join(parts, ";"))
	}
	if verbose {
		for idx, line := range lines {
			sep := strings.Split(line, ";")
			fmt.Printf("%d [%d]: %s\n", idx, len(sep), line)
		}
	}

	correctedCSV = strings.Join(lines, "\n")
	if err != io.EOF {
		fmt.Println(err)
	}

	reader := csv.NewReader(strings.NewReader(correctedCSV))
	reader.Comma = rune(';')

	excelFile := excelize.NewFile()

	fields, err := reader.Read()
	rows := 1
	for err == nil {
		for col, value := range fields {
			coords, _ := excelize.CoordinatesToCellName(col+1, rows)
			excelFile.SetCellValue(excelFile.GetSheetName(excelFile.GetActiveSheetIndex()), coords, value)
		}
		fields, err = reader.Read()
		rows++
	}
	fmt.Printf("%d rows extracted from csv file\n", rows-1)
	if err != nil && err != io.EOF {
		fmt.Println(err)
	}
	if verbose {
		excelFile.SaveAs("convertedCSV.xlsx")
	}
	return excelFile
}
