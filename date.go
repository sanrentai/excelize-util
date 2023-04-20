package excelizeutil

import (
	"strconv"

	"github.com/xuri/excelize/v2"
)

func ExcelDateToDateString(v string) (string, error) {
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return "", err
	}
	t, err := excelize.ExcelDateToTime(f, false)
	if err != nil {
		return "", err
	}
	return t.Format("2006-01-02"), nil
}

func GetDateString(f *excelize.File, sheet, cell string) (string, error) {
	v, err := f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	if err != nil {
		return "", err
	}
	return ExcelDateToDateString(v)
}
