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

// addCmd represents the add command
var addCmd = &cobra.Command{
	Use:   "add",
	Short: "add adds a csv file from proad to the employee workload file",
	Long: `proad exports the workload of each employee as a csv file.
	add sorts the csv file and extracts its content. 
	The content is then added to the emplyee workload file.`,
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Println("add called")
		if len(args) != 1 {
			fmt.Println("requires only one path argument")
			return
		}

		read := excel.Open(args[0])
		colmap := excel.FilterColumns([]int{1, 2, 4, 7, 8, 9}, read)
		// for i, v := range colmap[9] {
		// 	if i%5 == 0 {
		// 		fmt.Printf("%.2f\n", v)
		// 	} else {
		// 		fmt.Printf("%.2f\t", v)
		// 	}
		// }
		for i := 0; i < len(colmap[1]); i++ {
			if caseInsensitiveContains(fmt.Sprintf("%s", colmap[7][i]), "pitch") {
				fmt.Print("pitch")
				continue
			}

			switch colmap[8][i] {
			case jobNrNoWork:
				fmt.Println(jobNrNoWork)
			case jobNrOvertime:
				fmt.Println(jobNrOvertime)
			case jobNrSick:
				fmt.Println(jobNrSick)
			case jobNrVacation:
				fmt.Println(jobNrVacation)
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
