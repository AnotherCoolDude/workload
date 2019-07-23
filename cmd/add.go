/*
Copyright © 2019 NAME HERE <EMAIL ADDRESS>

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
package cmd

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/AnotherCoolDude/workload/excel"

	"github.com/spf13/cobra"
)

const (
	jobNrOvertime = "SEIN-0001-0167"
	jobNrNoWork   = "SEIN-0001-0169"
	jobNrSick     = "SEIN-0001-0015"
	jobNrVacation = "SEIN-0001-0012"
)

var (
	freelancer = []string{"Tina Botz", "Jörg Tacke"}
	tempPath   = ".tempxlsx.xlsx"
)

// addCmd represents the add command
var addCmd = &cobra.Command{
	Use:   "add",
	Short: "add adds a csv file from proad to the employee workload file",
	Long: `proad exports the workload of each employee as a csv file.
	add sorts the csv file and extracts its content. 
	The content is then added to the emplyee workload file.`,
	Run: func(cmd *cobra.Command, args []string) {
		if len(args) != 1 {
			fmt.Println("requires only one path argument")
			return
		}
		var ttfilePath string
		if strings.HasSuffix(args[0], "csv") {
			converted := excel.ConvertCSV(args[0], false)
			converted.SaveAs(tempPath)
			ttfilePath = tempPath
		} else if strings.HasSuffix(args[0], "xlsx") {
			ttfilePath = args[0]
		}

		wf := excel.OpenWorkloadFile(WorkloadFileName)
		read := excel.ReadProadExcel(ttfilePath)
		colmap := read.GetColumns([]int{1, 2, 4, 7, 8, 9})

		currentPeriodColumn := ""
		for _, sheetname := range wf.ModifiableSheetnames() {
			currentPeriodColumn = wf.DeclareNewColumnWithNextPeriod(sheetname)
		}

		for i := 1; i < len(colmap[1]); i++ {
			employeeName := fmt.Sprintf("%s", colmap[2][i])
			workhours := parseFloat(colmap[9][i])
			jobnr := colmap[8][i]

			if isFreelancer(employeeName) {
				continue
			}

			switch jobnr {
			case jobNrNoWork:
				wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[2], currentPeriodColumn)
			case jobNrOvertime:
				wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[7], currentPeriodColumn)
			case jobNrSick:
				wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[5], currentPeriodColumn)
			case jobNrVacation:
				wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[4], currentPeriodColumn)
			default:
				if caseInsensitiveContains(fmt.Sprintf("%s", colmap[7][i]), "pitch") {
					wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[1], currentPeriodColumn)
				} else if strings.Contains(jobnr, "SEIN") {
					wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[3], currentPeriodColumn)
				} else {
					wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[0], currentPeriodColumn)
				}
			}
		}
		wf.Save(WorkloadFileName)
		// delete temp file
		wd, err := os.Getwd()
		if err != nil {
			fmt.Println(err)
		}
		_, err = os.Stat(wd + "/" + tempPath)
		if os.IsExist(err) {
			err = os.Remove(wd + "/" + tempPath)
			if err != nil {
				fmt.Println(err)
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(addCmd)

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// addCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// addCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
}

func caseInsensitiveContains(s, substr string) bool {
	s, substr = strings.ToUpper(s), strings.ToUpper(substr)
	return strings.Contains(s, substr)
}

func convertToWorkloadFileName(name string) string {
	separatedNames := strings.Split(name, " ")
	return strings.TrimSpace(fmt.Sprintf("%s, %s", separatedNames[1], separatedNames[0]))
}

func isFreelancer(name string) bool {
	for _, fl := range freelancer {
		if name == fl {
			return true
		}
	}
	return false
}

func parseFloat(value string) float64 {
	parseValue := value
	if strings.IndexAny(value, ",") > -1 {
		parseValue = strings.Replace(value, ",", ".", 1)
	}
	float, err := strconv.ParseFloat(parseValue, 64)
	if err != nil {
		fmt.Println(err)
		return 0.0
	}
	return float
}
