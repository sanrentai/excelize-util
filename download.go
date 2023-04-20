package excelizeutil

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func Putdata(f *excelize.File, s *excelize.Style, n *excelize.Style, sheet string, heads, key []string, data []map[string]interface{}) error {

	style, err := f.NewStyle(s)
	if err != nil {
		return err
	}

	style2, err := f.NewStyle(n)
	if err != nil {
		return err
	}
	// 添加新工作表
	streamWriter, err := f.NewStreamWriter(sheet)
	if err != nil {
		return err
	}

	// var start, end string
	// f.SetActiveSheet(index)
	streamWriter.SetColWidth(2, len(heads), 16)
	var row = 1
	// 头部写入
	var hjarr = make([]int, len(heads))

	realHead := make([]interface{}, len(heads))
	for i, head := range heads {
		if hjs := strings.Split(head, "|"); len(hjs) > 1 {
			realHead[i] = excelize.Cell{
				StyleID: style,
				Value:   hjs[0],
			}
			if hjs[1] == "合计" {
				hjarr[i] = 1
			}
		} else {
			realHead[i] = excelize.Cell{
				StyleID: style,
				Value:   head,
			}
		}
	}
	err = streamWriter.SetPanes(&excelize.Panes{
		Freeze:      true,
		Split:       false,
		XSplit:      0,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "bottomLeft",
		Panes: []excelize.PaneOptions{
			{
				SQRef:      "A2",
				ActiveCell: "A2",
				Pane:       "bottomLeft",
			},
		},
	})
	if err != nil {
		return err
	}
	err = streamWriter.SetRow("A1", realHead)
	if err != nil {
		return err
	}
	// 写入文件头完成
	// 主体写入数据
	for _, record := range data {
		row++
		sli := make([]interface{}, len(key))
		for i, v := range key {
			if hjarr[i] == 1 {
				sli[i] = excelize.Cell{
					StyleID: style2,
					Value:   record[v],
				}
			} else {
				sli[i] = excelize.Cell{
					StyleID: style,
					Value:   record[v],
				}
			}
		}
		colstr, _ := excelize.CoordinatesToCellName(1, row)
		streamWriter.SetRow(colstr, sli)
	}

	hashj := false

	for i := range hjarr {
		if hjarr[i] == 1 {
			hashj = true
			break
		}
	}

	// 写合计公式
	if hashj {
		rowStrEnd := row
		row++
		rowStr := strconv.Itoa(row)
		hjrow := make([]interface{}, len(hjarr))

		for i := range hjarr {
			if hjarr[i] == 1 {
				hjrow[i] = excelize.Cell{
					StyleID: style2,
					Formula: fmt.Sprintf("SUM(%s%d:%s%d)", Col(i), 2, Col(i), rowStrEnd),
				}
			} else {
				hjrow[i] = excelize.Cell{
					StyleID: style,
					Value:   "",
				}
			}

		}
		hjrow[0] = excelize.Cell{
			StyleID: style,
			Value:   "合计",
		}
		streamWriter.SetRow("A"+rowStr, hjrow)
	}

	err = streamWriter.Flush()
	if err != nil {
		return err
	}

	return nil
}
