// Copyright © 2019 NAME HERE <EMAIL ADDRESS>
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

package cmd

import (
	"fmt"

	"github.com/AnotherCoolDude/workload/excel"
	"github.com/spf13/cobra"
)

// employeeCmd represents the employee command
var employeeCmd = &cobra.Command{
	Use:   "employee",
	Short: "adds an employee to the workload file",
	Long: `adds an employee to all modifiable sheets of the workload file.
	The new employee will be insertet alphabeticaly and into the chosen department.
	Possible Departments are: 
	
	Geschäftsführung	Beratung	Kreative

	Produktion			Text		Verwaltung

	Auszubildende/Trainee	PR`,
	Run: func(cmd *cobra.Command, args []string) {
		wf := excel.OpenWorkloadFile(workloadFileName)

		if len(args) != 2 {
			fmt.Println("two arguments are required. The name of the new Employee, and the department, e.g. workload add employee Christian Beratung")
		}
		if !validDepartment(args[1]) {
			fmt.Println("the provided department doesnt exist. Type workload add employee --help for more information.")
			return
		}
		wf.AddEmployee(args[0], excel.Department(args[1]))

		wf.Save("newEmployee.xlsx")
	},
}

func init() {
	addCmd.AddCommand(employeeCmd)

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// employeeCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// employeeCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
}

func validDepartment(department string) bool {
	for _, dep := range excel.Departments() {
		if dep == department {
			return true
		}
	}
	return false
}
