package xlsx

import (
	"fmt"
	"strings"

	"bufio"
	"bytes"
	"github.com/tealeg/xlsx"
)

func ToCSV(bs []byte, sheetIndex int, delimiter string) ([]byte, error) {
	xlFile, err := xlsx.OpenBinary(bs)

	if err != nil {
		return nil, err
	}
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return nil, fmt.Errorf("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return nil, fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}
	sheet := xlFile.Sheets[sheetIndex]

	var b bytes.Buffer
	writer := bufio.NewWriter(&b)

	for _, row := range sheet.Rows {
		var vals []string
		if row != nil {
			for _, cell := range row.Cells {
				val, _ := cell.String()
				vals = append(vals, fmt.Sprintf("%q", val))
			}
			writer.WriteString(strings.Join(vals, delimiter) + "\n")
		}
	}
	writer.Flush()
	return b.Bytes(), nil
}
