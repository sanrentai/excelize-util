package excelizeutil

import (
	"strconv"

	"github.com/xuri/excelize/v2"
)

func ExcelDateToDateString(v string) string {
	f, _ := strconv.ParseFloat(v, 64)
	t, _ := excelize.ExcelDateToTime(f, false)
	return t.Format("2006-01-02")
}

func GetDateString(f *excelize.File, sheet, cell string) string {
	v, _ := f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	return ExcelDateToDateString(v)
}
