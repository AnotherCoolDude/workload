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
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/spf13/viper"

	"github.com/manifoldco/promptui"

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
	freelancer []string
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

		var ttfilePath string
		freelancer = viper.GetStringSlice("freelancer")

		if len(args) == 0 {
			wd, files := checkForPossibleFiles()
			if len(files) == 0 {
				fmt.Println("no files found")
				return
			}
			prompt := promptui.Select{
				Label: "Choose file to be added to workload file",
				Items: files,
			}
			_, result, err := prompt.Run()
			if err != nil {
				fmt.Println(err)
				return
			}
			ttfilePath = wd + "/" + result
		} else if len(args) == 1 {
			_, file := filepath.Split(args[0])
			if file == workloadFileName {
				fmt.Println("cannot take workloadfile as argument")
				return
			}
			suf := filepath.Ext(args[0])
			if suf != ".csv" && suf != ".xlsx" {
				fmt.Println("provided file has wrong suffix")
				return
			}
			ttfilePath = args[0]
		} else {
			fmt.Println("requires one path argument")
			return
		}

		if strings.HasSuffix(ttfilePath, "csv") {
			converted := excel.ConvertCSV(ttfilePath, verbose)
			converted.SaveAs(tempPath)
			ttfilePath = tempPath
		}

		wf := excel.OpenWorkloadFile(workloadFileName)
		read := excel.ReadProadExcel(ttfilePath)
		colmap := read.GetColumns([]int{1, 2, 4, 7, 8, 9})

		currentPeriodColumn := ""
		for _, sheetname := range wf.ModifiableSheetnames() {
			currentPeriodColumn = wf.DeclareNewColumnWithNextPeriod(sheetname)
		}

		for i := 0; i < len(colmap[1]); i++ {
			employeeName := colmap[2][i]
			workhours := parseFloat(colmap[9][i])
			jobnr := colmap[8][i]
			var sheetname string
			if isFreelancer(employeeName) {
				continue
			}

			switch jobnr {
			case jobNrNoWork:
				sheetname = wf.ModifiableSheetnames()[2]
				wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[2], currentPeriodColumn)
			case jobNrOvertime:
				sheetname = wf.ModifiableSheetnames()[7]
				wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[7], currentPeriodColumn)
			case jobNrSick:
				sheetname = wf.ModifiableSheetnames()[5]
				wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[5], currentPeriodColumn)
			case jobNrVacation:
				sheetname = wf.ModifiableSheetnames()[4]
				wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[4], currentPeriodColumn)
			default:
				if caseInsensitiveContains(fmt.Sprintf("%s", colmap[7][i]), "pitch") {
					sheetname = wf.ModifiableSheetnames()[1]
					wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[1], currentPeriodColumn)
				} else if strings.Contains(jobnr, "SEIN") {
					sheetname = wf.ModifiableSheetnames()[3]
					wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[3], currentPeriodColumn)
				} else {
					sheetname = wf.ModifiableSheetnames()[0]
					wf.AddValueToEmployee(employeeName, workhours, wf.ModifiableSheetnames()[0], currentPeriodColumn)
				}
			}
			if verbose {
				fmt.Printf("[%s] %.2f added to %s\n", employeeName, workhours, sheetname)
			}

		}
		wf.Save(workloadFileName)
		// delete temp file
		wd, err := os.Getwd()
		if err != nil {
			fmt.Println(err)
		}
		_, err = os.Stat(wd + "/" + tempPath)
		if os.IsNotExist(err) {
			return
		}
		err = os.Remove(wd + "/" + tempPath)
		if err != nil {
			fmt.Println(err)
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

func checkForPossibleFiles() (workingDir string, files []string) {
	wd, err := os.Getwd()
	if err != nil {
		fmt.Println(err)
	}
	infos, err := ioutil.ReadDir(wd)
	if err != nil {
		fmt.Println(err)
	}
	proadFiles := []string{}
	for _, info := range infos {

		if info.IsDir() {
			continue
		}
		if info.Name() == workloadFileName {
			continue
		}
		ext := filepath.Ext(info.Name())
		if ext != ".csv" && ext != ".xlsx" {
			continue
		}

		proadFiles = append(proadFiles, info.Name())
	}
	return wd, proadFiles

}
