package xlsx

import (
	"bufio"
	"bytes"
	"fmt"
	stringsUtil "github.com/streamrail/common/util/strings"
	"github.com/tealeg/xlsx"
	"time"
)

type Format int

const (
	FormatIntNumber Format = iota
	FormatPercentNumber
	FormatFloatNumber
	FormatDollarNumber
	FormatString
	FormatDate
)

type Header struct {
	Value  string
	Format Format
	Width  float64
}

func GetExcelData(headers []Header, data [][]interface{}) ([]byte, error) {
	file := xlsx.NewFile()
	sheet, err := file.AddSheet("SR Data Export")
	if err != nil {
		fmt.Printf(err.Error())
	}

	var b bytes.Buffer
	foo := bufio.NewWriter(&b)

	if len(headers) > 0 {
		row := sheet.AddRow()
		for i, h := range headers {
			sheet.SetColWidth(i, i, h.Width)
			cell := row.AddCell()
			style := xlsx.NewStyle()
			style.Fill = *xlsx.NewFill("solid", "00B4CF", "")
			style.Font = *xlsx.NewFont(12, "Lato")
			style.Alignment.Horizontal = "center"
			style.ApplyAlignment = true
			style.Font.Color = "FFFFFF"
			style.Border = *xlsx.NewBorder("thin", "thin", "thin", "thin")

			style.ApplyFill = true
			style.ApplyFont = true
			style.ApplyBorder = true

			cell.SetStyle(style)

			cell.Value = h.Value
		}
	}

	for _, i := range data {
		row := sheet.AddRow()
		for idx, val := range i {
			if len(headers[idx].Value) == 0 {
				continue
			}
			cell := row.AddCell()
			if val, ok := val.(string); ok {
				style := xlsx.NewStyle()
				style.Font = *xlsx.NewFont(12, "Lato")
				style.Alignment.Horizontal = "center"
				style.ApplyAlignment = true
				style.ApplyFont = true
				cell.SetStyle(style)

				if headers[idx].Format == FormatString {
					cell.SetString(val)
				}
				if headers[idx].Format == FormatDollarNumber {
					cell.SetFloatWithFormat(stringsUtil.Float64OrDefault(val, -1), "$#,##0.00")
				}
				if headers[idx].Format == FormatIntNumber {
					cell.SetFloatWithFormat(stringsUtil.Float64OrDefault(val, -1), "#,##0")
				}
				if headers[idx].Format == FormatFloatNumber {
					cell.SetFloatWithFormat(stringsUtil.Float64OrDefault(val, -1), "#,##0.00")
				}
				if headers[idx].Format == FormatPercentNumber {
					cell.SetFloatWithFormat(stringsUtil.Float64OrDefault(val, -1), "0.00%")
				}
				if headers[idx].Format == FormatDate {
					d, err := time.Parse("2006-01-02", val)
					if err != nil {
						return nil, fmt.Errorf("error parsing value as date: %s", val)
					}
					cell.SetDate(d)
				}
			} else {
				return nil, fmt.Errorf("error parsing value as string: %s", val)
			}

		}
	}

	file.Write(foo)
	foo.Flush()

	return b.Bytes(), nil
}
