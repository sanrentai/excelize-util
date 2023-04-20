package excelizeutil

import "github.com/xuri/excelize/v2"

func Col(col int) string {
	var i = col + 1
	n, _ := excelize.ColumnNumberToName(i)
	return n
}
