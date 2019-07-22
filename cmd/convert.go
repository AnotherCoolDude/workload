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
	"strings"

	"github.com/AnotherCoolDude/workload/excel"
	"github.com/spf13/cobra"
)

// convertCmd represents the convert command
var convertCmd = &cobra.Command{
	Use:   "convert",
	Short: "converts csv to excel",
	Long:  `convert converts a csv file with ; delimeter to an excel file`,
	Run: func(cmd *cobra.Command, args []string) {
		if !validatePathAndFile(args[0]) {
			fmt.Println("wrong input file")
			os.Exit(0)
		}
		f := excel.ConvertCSV(args[0])
		trimmed := strings.TrimSuffix(args[0], "csv")
		f.SaveAs(trimmed + "xlsx")
	},
}

func init() {
	rootCmd.AddCommand(convertCmd)

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// convertCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// convertCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
}

func validatePathAndFile(path string) bool {
	stat, err := os.Stat(path)
	if os.IsNotExist(err) || stat.IsDir() {
		return false
	}
	suffix := strings.HasSuffix(path, "csv")
	if !suffix {
		return false
	}
	return true
}
