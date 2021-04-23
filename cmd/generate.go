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
	//"io/ioutil"
	//"strconv"

	//"xlxs-helper/util"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/spf13/cobra"
)

type client struct {
	Name string
	Addr string
	Id   string
	// 贷款金额
	LoanAmount string
	// 贷款余额
	LoanBalance string
	// 到期日期
	Expiry string
	// 用途
	LoanPurpose string
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
}

func expand(cmd *cobra.Command, args []string) {
	var inputFile = "D:\\loan\\input.xlsx"
	var outputFile = "D:\\loan\\template.xlsx"

	inputXlsx, err := excelize.OpenFile(inputFile)
	if err != nil {
		fmt.Println(err)
		return
	}

	outputXlsx, err := excelize.OpenFile(outputFile)
	if err != nil {
		fmt.Println(err)
		return
	}

	inputRows := inputXlsx.GetRows("Sheet1")
	for rowIdx, rowInfo := range inputRows {
		// 跳过第一行名称
		if rowIdx == 0 {
			continue
		}
		// 0 名字
		// 1 地址
		// 2 身份证
		// 3 身份证 id
		// 4 贷款金额
		// 5 贷款余额
		// 6 到期日去
		// 7 用途

		clientInfo := new(client)
		clientInfo.Name = rowInfo[0]
		clientInfo.Addr = rowInfo[1]
		clientInfo.Id = rowInfo[3]
		clientInfo.LoanAmount = rowInfo[4]
		clientInfo.LoanBalance = rowInfo[5]
		clientInfo.Expiry = rowInfo[6]
		clientInfo.LoanPurpose = rowInfo[7]

		sheetName := clientInfo.Name
		sheetIndex := outputXlsx.NewSheet(sheetName)
		if err := outputXlsx.CopySheet(1, sheetIndex); err != nil {
			fmt.Println(err)
			return
		}

		outputXlsx.SetCellValue(sheetName, "C3", sheetName)
		outputXlsx.SetCellValue(sheetName, "G3", clientInfo.Addr)
		outputXlsx.SetCellValue(sheetName, "E4", clientInfo.Id)
		outputXlsx.SetCellValue(sheetName, "C5", clientInfo.LoanAmount)
		outputXlsx.SetCellValue(sheetName, "G5", clientInfo.LoanBalance)
		outputXlsx.SetCellValue(sheetName, "C7", clientInfo.Expiry)
		outputXlsx.SetCellValue(sheetName, "E7", clientInfo.LoanPurpose)
	}

	outputXlsx.SaveAs("D:\\loan\\output.xlsx")
	fmt.Println(outputXlsx.SheetCount)

	//groupXlsx, err := excelize.OpenFile(fileDir + file.Name())
	//if err != nil {
	//	fmt.Printf("?? %s ??????%s", file.Name(), err)
	//	fmt.Println()
	//	continue
	//}
	//
	//var rows = groupXlsx.GetRows("Sheet1")
	//// ??6??????
	//// ? 2 ? ????
	//// ? 3 ? ?ID?
	//// ? 11 ? ????
	//// ? 14 ? ??????
	//
	//var familyList []*family
	//var singleFamily *family
	//bFamily := true
	//
	//for rowIdx, rowInfo := range rows {
	//	if rowIdx <= 4 || rowInfo[2] == "" {
	//		continue
	//	}
	//
	//	var member = new(household)
	//	member.name = rowInfo[2]
	//	member.id = rowInfo[3]
	//	member.gender = rowInfo[11]
	//
	//	if rowInfo[14] == "?????" {
	//		singleFamily = new(family)
	//		bFamily = true
	//		member.header = true
	//	} else {
	//		member.relationship = rowInfo[14]
	//		member.header = false
	//		bFamily = false
	//	}
	//
	//	if bFamily {
	//		familyList = append(familyList, singleFamily)
	//	}
	//	singleFamily.members = append(singleFamily.members, member)
	//}
	//
	//for _, f := range familyList {
	//	var sheetName string
	//	baseMemberRowIdx := 8
	//	for _, member := range f.members {
	//		if member.header {
	//			// ????sheet??????
	//			sheetName = member.name
	//			sheetIndex := templateXlsx.NewSheet(sheetName)
	//			templateXlsx.CopySheet(1, sheetIndex)
	//
	//			// ???
	//			templateXlsx.SetCellValue(sheetName, "C4", member.name)
	//
	//			// ?????
	//			templateXlsx.SetCellValue(sheetName, "G4", member.gender)
	//
	//			// ???ID
	//			templateXlsx.SetCellValue(sheetName, "M4", member.id)
	//
	//			cell := templateXlsx.GetCellValue(sheetName, "A3")
	//			templateXlsx.SetCellValue(sheetName, "A3", cell+"???"+util.GetFilename(file.Name()))
	//		} else {
	//			// ??
	//			templateXlsx.SetCellValue(sheetName, "C"+strconv.Itoa(baseMemberRowIdx), member.name)
	//
	//			// ??????
	//			templateXlsx.SetCellValue(sheetName, "D"+strconv.Itoa(baseMemberRowIdx), member.relationship)
	//
	//			// ID
	//			templateXlsx.SetCellValue(sheetName, "E"+strconv.Itoa(baseMemberRowIdx), member.id)
	//			baseMemberRowIdx++
	//		}
	//	}
	//}
	//
	//templateXlsx.SetActiveSheet(2)
	//templateXlsx.SaveAs("./???/" + file.Name())
}
