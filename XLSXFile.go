package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"regexp"
	"strings"
	"time"
)

type XLSXFile struct {
	*excelize.File
	DestFile *os.File
	ReQuery string
}

type OneDayClasses struct {
	Date string
	Classes []string
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

	xlsx := XLSXFile{
		f,
		destFile,
		"INSERT INTO SCHEDULE (CLASS_DATE, SUB1, SUB2, SUB3, SUB4, SUB5) VALUES (%s);\n",
	}

	return &xlsx, nil
}

// 解析工作表
func (f *XLSXFile) parseSheet(sheet string) error {
	cols, err := f.GetCols(sheet)
	if err != nil {
		return err
	}

	cols = cols[2:]
	for i, col := range cols {
		if len(col) != 40 {
			return fmt.Errorf("第%d列不足40行", i+1)
		}

		col = col[2:38]
		weekClasses, err := f.parseCol(col)
		if err != nil {
			return err
		}

		for _, wc := range weekClasses {
			_, err = fmt.Fprintf(f.DestFile, f.ReQuery, wc.String())
			if err != nil {
				return err
			}
		}
	}
	return nil
}

// 解析课表中的每一列
func (f *XLSXFile) parseCol(col []string) ([]OneDayClasses, error) {
	monDate, err := f.parseMonDate(col[0])
	if err != nil {
		return nil, err
	}

	col = col[1:]

	var days []OneDayClasses

	for i := 0; i < len(col); i += 5 {
		oneDay := OneDayClasses{
			Date: monDate.Format("2006-01-02"),
			Classes: (col)[i:i+5],
		}

		days = append(days, oneDay)
		monDate = monDate.AddDate(0, 0, 1)
	}

	return days, nil
}

// 解析每周的第一天日期，返回time.Time和错误；
// 日期文本格式：01.02
func (f *XLSXFile) parseMonDate(colDate string) (time.Time, error) {
	re := regexp.MustCompile(`\d\d\.\d\d`)
	t := time.Now()
	timeStr := re.FindString(colDate)
	if timeStr == "" {
		return t, fmt.Errorf("无法解析本周周一日期：%s", colDate)
	}

	return time.Parse("2006.01.02", fmt.Sprintf("%d.%s", t.Year(), timeStr))
}
