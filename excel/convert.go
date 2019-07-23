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
	c := 0
	for {
		str, err = r.ReadString('\r')
		if err != nil {
			break
		}
		if str == "" {
			continue
		}
		if c < 2 {
			fmt.Printf("%q\n", str)
			c++
		}

		// strTrimmed := strings.TrimFunc(str, func(r rune) bool {
		// 	if r == '\n' || r == '\r' {
		// 		return true
		// 	}
		// 	return false
		// })
		if !strings.ContainsAny(str, "\"") {
			continue
		}
		if str[0] != '"' {
			previousStr := lines[len(lines)-1]
			str = previousStr[:len(previousStr)-1] + " " + str
			lines = lines[:len(lines)-1]
		}
		splitted := strings.Split(str, ";")
		parts := []string{}
		//fmt.Println(splitted[len(splitted)-1])

		for _, s := range splitted {
			trimmed := strings.TrimFunc(s, func(r rune) bool {
				if r == '\r' || r == '\n' {
					return true
				}
				return false
			})
			correctedQuotes := strings.ReplaceAll(trimmed[1:len(trimmed)-1], "\"", "\"\"")
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

	reader := csv.NewReader(strings.NewReader(correctedCSV)) //csv.NewReader(ReplaceSoloCarriageReturns(csvFile))
	reader.Comma = rune(';')
	reader.LazyQuotes = true

	excelFile := excelize.NewFile()

	fields, err := reader.Read()
	rows := 1
	for err == nil {
		for col, value := range fields {
			coords, _ := excelize.CoordinatesToCellName(col+1, rows)
			excelFile.SetCellValue(excelFile.GetSheetName(excelFile.GetActiveSheetIndex()), coords, value)
			//fmt.Printf("%s\t", value)
		}
		fields, err = reader.Read()
		rows++
	}
	fmt.Printf("Rows in excel file: %d\n", rows)
	if err != nil {
		fmt.Println(err)
	}
	return excelFile
}

// ReplaceSoloCarriageReturns wraps an io.Reader, on every call of Read it
// for instances of lonely \r replacing them with \r\n before returning to the end customer
// lots of files in the wild will come without "proper" line breaks, which irritates go's
// standard csv package. This'll fix by wrapping the reader passed to csv.NewReader:
//    rdr, err := csv.NewReader(ReplaceSoloCarriageReturns(r))
//
func ReplaceSoloCarriageReturns(data io.Reader) io.Reader {
	return crlfReplaceReader{
		rdr: bufio.NewReader(data),
	}
}

// crlfReplaceReader wraps a reader
type crlfReplaceReader struct {
	rdr *bufio.Reader
}

// Read implements io.Reader for crlfReplaceReader
func (c crlfReplaceReader) Read(p []byte) (n int, err error) {
	if len(p) == 0 {
		return
	}

	for {
		if n == len(p) {
			return
		}

		p[n], err = c.rdr.ReadByte()
		if err != nil {
			return
		}

		// any time we encounter \r & still have space, check to see if \n follows
		// if next char is not \n, add it in manually
		if p[n] == '\r' && n < len(p) {
			if pk, err := c.rdr.Peek(1); (err == nil && pk[0] != '\n') || (err != nil && err.Error() == io.EOF.Error()) {
				n++
				p[n] = '\n'
			}
		}

		n++
	}
}
