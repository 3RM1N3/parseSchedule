package main

import (
	"errors"
	"os"
	"path"
	"strings"
)

// ParseFile 解析文件
func ParseFile() error {
	fileName, err := getFile()
	if err != nil {
		return err
	}

	f, err := openFile(fileName)
	if err != nil {
		return err
	}

	sheetList := f.GetSheetList()
	if len(sheetList) == 0 {
		return errors.New("读取工作表名称错误")
	}
	sheet := sheetList[0]

	c3FirstWeek, err := f.GetCellValue(sheet, "C3")
	if err != nil {
		return err
	}

	if !strings.HasPrefix(c3FirstWeek, "第1周") {
		return errors.New("工作表格式不正确")
	}

	return f.parseSheet(sheet)
}

// 搜索运行目录下的xlsx文件，如果文件唯一则返回该文件，否则返回错误
func getFile() (string, error) {
	count := 0
	fileName := ""

	entrys, err := os.ReadDir(".")
	if err != nil {
		return "", err
	}

	for _, entry := range entrys {
		if entry.IsDir() {
			continue
		}

		if path.Ext(entry.Name()) != ".xlsx" {
			continue
		}

		if count > 1 {
			return "", errors.New("发现多个 *.xlsx 文件")
		}

		count++
		fileName = entry.Name()
	}

	if count == 0 {
		return "", errors.New("没有发现 *.xlsx 文件")
	}

	return fileName, nil
}
