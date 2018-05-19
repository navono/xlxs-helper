package main

import (
	"fmt"
	"io/ioutil"
	"path"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
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

func main() {

	var fileDir = "./2018麻墩/"

	// 遍历文件夹
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
			fmt.Printf("打开 %s 失败，原因：%s", file.Name(), err)
			fmt.Println()
			continue
		}

		var rows = groupXlsx.GetRows("Sheet1")
		// 从第6行开始读取：
		// 第 2 列 （姓名）
		// 第 3 列 （ID）
		// 第 11 列 （性别）
		// 第 14 列 （家庭关系）

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

			if rowInfo[14] == "本人或户主" {
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
					// 新建一个sheet，以户主为名
					sheetName = member.name
					sheetIndex := templateXlsx.NewSheet(sheetName)
					templateXlsx.CopySheet(1, sheetIndex)

					// 申请人
					templateXlsx.SetCellValue(sheetName, "C4", member.name)

					// 申请人性别
					templateXlsx.SetCellValue(sheetName, "G4", member.gender)

					// 申请人ID
					templateXlsx.SetCellValue(sheetName, "M4", member.id)

					cell := templateXlsx.GetCellValue(sheetName, "A3")
					templateXlsx.SetCellValue(sheetName, "A3", cell+"麻墩村"+getFilename(file.Name()))
				} else {
					// 姓名
					templateXlsx.SetCellValue(sheetName, "C"+strconv.Itoa(baseMemberRowIdx), member.name)

					// 与申请人关系
					templateXlsx.SetCellValue(sheetName, "D"+strconv.Itoa(baseMemberRowIdx), member.relationship)

					// ID
					templateXlsx.SetCellValue(sheetName, "E"+strconv.Itoa(baseMemberRowIdx), member.id)
					baseMemberRowIdx++
				}
			}
		}

		templateXlsx.SetActiveSheet(2)
		templateXlsx.SaveAs("./情况表/" + file.Name())
	}
}

func getFilename(filepath string) string {
	var filenameWithSuffix string
	filenameWithSuffix = path.Base(filepath)
	var fileSuffix string
	fileSuffix = path.Ext(filenameWithSuffix)

	var filenameOnly string
	filenameOnly = strings.TrimSuffix(filenameWithSuffix, fileSuffix)
	return filenameOnly
}
