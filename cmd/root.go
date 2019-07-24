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

	"github.com/spf13/cobra"

	homedir "github.com/mitchellh/go-homedir"
	"github.com/spf13/viper"
)

var (
	cfgFile          string
	verbose          bool
	workloadFileName string
)

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "workload",
	Short: "workload helps modifying (adding, deleting, ...) the workload file",
	Long: `keeping track of the emplyee workload can be time intense. 
  This programms automates the repetetive task of filling out the excel file,
  that keeps track of how the time of emplyees is used in different kinds of projects.`,
	// Uncomment the following line if your bare application
	// has an action associated with it:
	//	Run: func(cmd *cobra.Command, args []string) { },
}

// Execute adds all child commands to the root command and sets flags appropriately.
// This is called by main.main(). It only needs to happen once to the rootCmd.
func Execute() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

func init() {
	cobra.OnInitialize(InitWorkloadFile)

	// Here you will define your flags and configuration settings.
	// Cobra supports persistent flags, which, if defined here,
	// will be global for your application.

	// rootCmd.PersistentFlags().StringVar(&cfgFile, "config", "", "config file (default is $HOME/.workload.yaml)")

	// Cobra also supports local flags, which will only run
	// when this action is called directly.
	// rootCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
	addCmd.PersistentFlags().BoolVarP(&verbose, "verbose", "v", false, "prints additional information")
}

// InitWorkloadFile checks wether the excel file for employee workload is available
func InitWorkloadFile() {
	wd, err := os.Getwd()
	if err != nil {
		panic(err)
	}
	fmt.Printf("current working directory: %s\n", wd)

	viper.AddConfigPath(wd)
	viper.SetConfigName("workload")
	err = viper.ReadInConfig()
	if err != nil {
		fmt.Println("no config file found. Creating a new one.")
		viper.SetDefault("workloadfilename", "Auslastung.xlsx")
		viper.SetDefault("workloadPath", wd)
		err = viper.SafeWriteConfig()
		if err != nil {
			fmt.Println(err)
		}
	}
	workloadFileName = fmt.Sprintf("%s", viper.Get("workloadfilename"))

	_, err = os.Stat(wd + "/" + workloadFileName)
	if os.IsNotExist(err) {
		fmt.Printf("couldn't find workload file. File must be placed in same ordner as the executable and named as %s\n", workloadFileName)
		fmt.Println(err)
		os.Exit(1)
	}

}

// initConfig reads in config file and ENV variables if set.
func initConfig() {
	if cfgFile != "" {
		// Use config file from the flag.
		viper.SetConfigFile(cfgFile)
	} else {
		// Find home directory.
		home, err := homedir.Dir()
		if err != nil {
			fmt.Println(err)
			os.Exit(1)
		}

		// Search config in home directory with name ".workload" (without extension).
		viper.AddConfigPath(home)
		viper.SetConfigName(".workload")
	}

	viper.AutomaticEnv() // read in environment variables that match

	// If a config file is found, read it in.
	if err := viper.ReadInConfig(); err == nil {
		fmt.Println("Using config file:", viper.ConfigFileUsed())
	}
}
