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
	"fmt"
	"io/ioutil"
	"strconv"

	"xlxs-helper/util"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/spf13/cobra"
)

type household struct {
	name         string
	id           string
	gender       string
	header       bool
	relationship string
}

type family struct {
	members []*household
}

// oldCmd represents the old command
var generateCmd = &cobra.Command{
	Use:   "generate",
	Short: "新建 Sheet 页",
	Long:  `通过指定参数，生成 Sheet 页`,
	Run:   expand,
}

func init() {
	rootCmd.AddCommand(generateCmd)

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// oldCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// oldCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
}

func expand(cmd *cobra.Command, args []string) {
	fmt.Println("old called")

	var fileDir = "./2018??/"

	// ?????
	files, err := ioutil.ReadDir(fileDir)
	if err != nil {
		fmt.Println(err)
		return
	}

	for _, file := range files {
		templateXlsx, err := excelize.OpenFile("./2.xls")
		// templateXlsx, err := excelize.OpenFile("./1.xlsx")
		if err != nil {
			fmt.Println(err)
			return
		}

		groupXlsx, err := excelize.OpenFile(fileDir + file.Name())
		if err != nil {
			fmt.Printf("?? %s ??????%s", file.Name(), err)
			fmt.Println()
			continue
		}

		var rows = groupXlsx.GetRows("Sheet1")
		// ??6??????
		// ? 2 ? ????
		// ? 3 ? ?ID?
		// ? 11 ? ????
		// ? 14 ? ??????

		var familyList []*family
		var singleFamily *family
		bFamily := true

		for rowIdx, rowInfo := range rows {
			if rowIdx <= 4 || rowInfo[2] == "" {
				continue
			}

			var member = new(household)
			member.name = rowInfo[2]
			member.id = rowInfo[3]
			member.gender = rowInfo[11]

			if rowInfo[14] == "?????" {
				singleFamily = new(family)
				bFamily = true
				member.header = true
			} else {
				member.relationship = rowInfo[14]
				member.header = false
				bFamily = false
			}

			if bFamily {
				familyList = append(familyList, singleFamily)
			}
			singleFamily.members = append(singleFamily.members, member)
		}

		for _, f := range familyList {
			var sheetName string
			baseMemberRowIdx := 8
			for _, member := range f.members {
				if member.header {
					// ????sheet??????
					sheetName = member.name
					sheetIndex := templateXlsx.NewSheet(sheetName)
					templateXlsx.CopySheet(1, sheetIndex)

					// ???
					templateXlsx.SetCellValue(sheetName, "C4", member.name)

					// ?????
					templateXlsx.SetCellValue(sheetName, "G4", member.gender)

					// ???ID
					templateXlsx.SetCellValue(sheetName, "M4", member.id)

					cell := templateXlsx.GetCellValue(sheetName, "A3")
					templateXlsx.SetCellValue(sheetName, "A3", cell+"???"+util.GetFilename(file.Name()))
				} else {
					// ??
					templateXlsx.SetCellValue(sheetName, "C"+strconv.Itoa(baseMemberRowIdx), member.name)

					// ??????
					templateXlsx.SetCellValue(sheetName, "D"+strconv.Itoa(baseMemberRowIdx), member.relationship)

					// ID
					templateXlsx.SetCellValue(sheetName, "E"+strconv.Itoa(baseMemberRowIdx), member.id)
					baseMemberRowIdx++
				}
			}
		}

		templateXlsx.SetActiveSheet(2)
		templateXlsx.SaveAs("./???/" + file.Name())
	}
}
