package main

import (
	"fmt"
	"os"
	"regexp"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type XLSXFile struct {
	*excelize.File          // 继承excelize.File结构体
	DestFile       *os.File // 目标文件
	ReQuery        string   // 生成目标文件每一行的内容
}

type OneDayClasses struct {
	Date    string   // 日期字符串
	Classes []string // 课程
}

func (o *OneDayClasses) String() string {
	return fmt.Sprintf("\"%s\", \"%s\"", o.Date, strings.Join(o.Classes, "\", \""))
}

// openFile 封装打开文件函数
func openFile(fileName string) (*XLSXFile, error) {
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, err
	}

	destName := strings.TrimSuffix(fileName, ".xlsx") + ".sql"
	destFile, err := os.Create(destName)
	if err != nil {
		return nil, err
	}

	gmg := getGMG(fileName)

	fmt.Fprintf(destFile, "create table "+gmg+` (
    CLASS_DATE TEXT unique not null,
    SUB1       TEXT,
    SUB2       TEXT,
    SUB3       TEXT,
    SUB4       TEXT,
    SUB5       TEXT
);

`)

	xlsx := XLSXFile{
		f,
		destFile,
		"INSERT INTO " + gmg + " (CLASS_DATE, SUB1, SUB2, SUB3, SUB4, SUB5) VALUES (%s);\n",
	}

	return &xlsx, nil
}

// 获取此课表的年级专业班级（未完成）
func getGMG(fileName string) string {
	return "S_18520105"
}

// 解析一个sheet
func (f *XLSXFile) parseSheet(sheetName string) error {
	cols, err := f.GetCols(sheetName)
	if err != nil {
		return err
	}

	cols = cols[2:]            // 去掉前两列（星期、单元）
	for i, col := range cols { // 循环处理每一列
		if len(col) != 40 {
			return fmt.Errorf("第%d列不足40行", i+1)
		}

		col = col[2:38]                     // 取得第3行至第38行的内容（日期+课表）
		weekClasses, err := f.parseCol(col) // 处理一列的内容
		if err != nil {
			return err
		}

		for _, oneDayClasses := range weekClasses {
			_, err = fmt.Fprintf(f.DestFile, f.ReQuery, oneDayClasses.String()) // 向目标文件写入一天课表
			if err != nil {
				return err
			}
		}
	}
	return nil
}

// 解析课表中的一列
func (f *XLSXFile) parseCol(col []string) ([]OneDayClasses, error) {
	monDate, err := f.parseMonDate(col[0]) // 取得周一的日期
	if err != nil {
		return nil, err
	}

	col = col[1:] // 去掉第一行的日期

	var days []OneDayClasses

	for i := 0; i < len(col); i += 5 {
		oneDay := OneDayClasses{
			Date:    monDate.Format("2006-01-02"),
			Classes: (col)[i : i+5],
		}

		days = append(days, oneDay)
		monDate = monDate.AddDate(0, 0, 1)
	}

	return days, nil
}

// 解析每周的第一天日期，返回time.Time和错误；
// 日期文本格式：01.02
func (f *XLSXFile) parseMonDate(colDate string) (time.Time, error) {
	t := time.Now() // 获取当天日期

	re := regexp.MustCompile(`\d\d\.\d\d`)
	timeStr := re.FindString(colDate) // 获取日期字符串

	if timeStr == "" {
		return t, fmt.Errorf("无法解析本周周一日期：%s", colDate)
	}

	return time.Parse("2006.01.02", fmt.Sprintf("%d.%s", t.Year(), timeStr))
}
