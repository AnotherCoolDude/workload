package excel

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// ConvertCSV converts a csv file to an excel file
func ConvertCSV(path string) *excelize.File {
	csvFile, err := os.Open(path)
	if err != nil {
		fmt.Println(err)
		os.Exit(0)
	}
	defer csvFile.Close()

	f, _ := ioutil.ReadFile(path)
	corrected := strings.Replace(string(f), "\r", "\n", -1)
	// scanner := bufio.NewScanner(csvFile)
	// for scanner.Scan() {
	// 	fmt.Println(scanner.Text())
	// }

	reader := csv.NewReader() //csv.NewReader(ReplaceSoloCarriageReturns(csvFile))
	reader.Comma = rune(';')
	reader.LazyQuotes = true

	excelFile := excelize.NewFile()

	fields, err := reader.Read()
	rows := 1
	for err == nil {
		for col, value := range fields {
			coords, _ := excelize.CoordinatesToCellName(col+1, rows)
			excelFile.SetCellValue(excelFile.GetSheetName(excelFile.GetActiveSheetIndex()), coords, value)
			fmt.Printf("%s\t", value)
			fields, err = reader.Read()
		}
		fmt.Println()
		rows++
	}

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
