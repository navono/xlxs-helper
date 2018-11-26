// Copyright © 2018 NAME HERE <EMAIL ADDRESS>
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
	"encoding/json"
	"fmt"
	"io/ioutil"
	"strings"
	"sync"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/spf13/cobra"
)

type searchResult struct {
	SearchName string `json:"searchName"`
	Filename   string `json:"filename"`
	Col        int    `json:"col"`
	Row        int    `json:"row"`
}

// searchCmd represents the search command
var searchCmd = &cobra.Command{
	Use:   "search",
	Short: "查找",
	Long: `用于在文件夹中，通过指定名称，查找匹配的 Sheet 页以及行、列。
	例如，在同层文件夹中的所有 .xlxs 文件中查找 王五：
	xxx search -d "." -n "王五"
	`,
	SilenceUsage: true,
	// Args: func(cmd *cobra.Command, args []string) error {
	// 	fmt.Println(args)
	// 	if len(args) < 1 {
	// 		return errors.New("至少需要一个参数 -n")
	// 	}
	// 	// if myapp.IsValidColor(args[0]) {
	// 	//   return nil
	// 	// }
	// 	return fmt.Errorf("invalid color specified: %s", args[0])
	// },
	Run: search,
}

func init() {
	rootCmd.AddCommand(searchCmd)

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// searchCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	searchCmd.Flags().StringP("name", "n", "", "需要查找的名称")
	searchCmd.Flags().StringP("dir", "d", ".", "xlxs 文件夹")
}

func search(cmd *cobra.Command, args []string) {
	dir, _ := cmd.Flags().GetString("dir")
	name, _ := cmd.Flags().GetString("name")

	files, err := ioutil.ReadDir(dir)
	if err != nil {
		fmt.Println(err)
		return
	}

	wg := &sync.WaitGroup{}

	result := make(chan interface{})

	for _, file := range files {
		wg.Add(1)
		go searchInFile(dir+"/"+file.Name(), name, result, wg)
	}

	// 监听 WaitGroup
	go func() {
		wg.Wait()
		close(result)
	}()

	// 合并查询结果到文件中
	done := make(chan bool, 1)
	go mergeResultToFile(result, done)
	<-done
}

func searchInFile(filename string, name string, resultChan chan<- interface{}, wg *sync.WaitGroup) {
	defer wg.Done()

	groupXlsx, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Printf("读取 %s 错误。%s", filename, err)
		fmt.Println()
		return
	}

	var rows = groupXlsx.GetRows("Sheet1")
	for rowIdx, row := range rows {
		for colIdx, colCell := range row {
			if strings.Compare(name, colCell) != 0 {
				continue
			}

			result := &searchResult{
				SearchName: name,
				Filename:   filename,
				Row:        rowIdx + 1,
				Col:        colIdx + 1,
			}
			data, _ := json.MarshalIndent(result, "", "  ")
			resultChan <- data
		}
	}
}

func mergeResultToFile(result <-chan interface{}, done chan<- bool) {
	var jsonData []byte
	for data := range result {
		bytes, _ := data.([]byte)
		for idx := range bytes {
			jsonData = append(jsonData, bytes[idx])
		}

		fmt.Println(string(jsonData))
	}

	err := ioutil.WriteFile("result.json", jsonData, 0644)
	check(err)
	done <- true
}

func check(e error) {
	if e != nil {
		panic(e)
	}
}
