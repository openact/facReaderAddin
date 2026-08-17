package lookup

import (
	"bufio"
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/openact/kit/cache/v3"
)

const (
	StatusOK           = 0
	StatusNotFound     = 1
	StatusDimMismatch  = 2
	StatusInvalid      = 3
	StatusParseError   = 4
	StatusWriteError   = 5
	StatusFileNotFound = 6
)

// ProductFromFilePath returns the filename stem used as default product key.
func ProductFromFilePath(path string) string {
	return strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
}

// EProj_Result resolves one exact-match cell from a cache/v3 table.
// Parameter order aligns with the external function: product, spCode,
// resultType, variable, timePeriod, simID.
// resultType is reserved for interface compatibility and is not used in lookup logic.
func EProj_Result(t *cache.Table, product, spCode, resultType, variable, timePeriod, simID string) (float64, bool) {
	_ = resultType // reserved parameter; intentionally not used.
	v, status := EProj_ResultStatus(t, product, spCode, variable, timePeriod, simID)
	return v, status == StatusOK
}

// ERead_Result resolves one exact-match cell using coordinates in table order.
// The number of coordinates must match the .fac table dimensions exactly:
// all row dimensions first, followed by the column dimension.
func ERead_Result(t *cache.Table, idx ...string) (float64, bool) {
	v, status := ERead_ResultStatus(t, idx...)
	return v, status == StatusOK
}

func ERead_ResultStatus(t *cache.Table, idx ...string) (float64, int) {
	if t == nil || t.NumDims <= 0 || len(idx) != t.NumDims {
		return 0, StatusDimMismatch
	}

	for i := 0; i < t.NumDims; i++ {
		if strings.TrimSpace(idx[i]) == "*" {
			return 0, StatusInvalid
		}
	}

	val := t.RawCell(idx...)
	if val == nil {
		return 0, StatusNotFound
	}
	v, err := strconv.ParseFloat(string(val), 64)
	if err != nil {
		return 0, StatusParseError
	}
	return v, StatusOK
}

func EProj_ResultStatus(t *cache.Table, product, spCode, variable, timePeriod, simID string) (float64, int) {
	row, status := projRowKeys(t, product, spCode, variable, simID)
	if status != StatusOK {
		return 0, status
	}
	col := strings.TrimSpace(timePeriod)
	if col == "" || col == "*" {
		return 0, StatusInvalid
	}
	txt, status := readCellText(t, row, col)
	if status != StatusOK {
		return 0, status
	}
	v, err := strconv.ParseFloat(txt, 64)
	if err != nil {
		return 0, StatusParseError
	}
	return v, StatusOK
}

func WriteEReadTable(t *cache.Table, rows [][]string, cols []string, outputPath string) int {
	if t == nil || t.NumDims <= 0 {
		return StatusInvalid
	}
	if len(cols) == 0 {
		return StatusInvalid
	}
	rowDims := t.NumDims - 1
	for _, row := range rows {
		if len(row) != rowDims {
			return StatusDimMismatch
		}
		if !validParts(row) {
			return StatusInvalid
		}
	}
	if !validParts(cols) {
		return StatusInvalid
	}

	return writeTable(t, rows, cols, outputPath)
}

func WriteEProjTable(t *cache.Table, product, simID string, projKeys [][]string, cols []string, outputPath string) int {
	if t == nil || t.NumDims <= 0 || len(cols) == 0 {
		return StatusInvalid
	}
	if !validParts(cols) {
		return StatusInvalid
	}

	rows := make([][]string, len(projKeys))
	for i, key := range projKeys {
		if len(key) != 2 {
			return StatusDimMismatch
		}
		row, status := projRowKeys(t, product, key[0], key[1], simID)
		if status != StatusOK {
			return status
		}
		rows[i] = row
	}
	return writeTable(t, rows, cols, outputPath)
}

func writeTable(t *cache.Table, rows [][]string, cols []string, outputPath string) int {
	colIdx := make([]int, len(cols))
	for i, col := range cols {
		ci, ok := columnIndex(t, col)
		if !ok {
			colIdx[i] = -1
			continue
		}
		colIdx[i] = ci
	}

	f, err := os.Create(outputPath)
	if err != nil {
		return StatusWriteError
	}
	defer f.Close()

	w := bufio.NewWriterSize(f, 4<<20)
	defer w.Flush()

	for r, row := range rows {
		raw := t.RawRow(row...)
		for c, ci := range colIdx {
			if c > 0 {
				if _, err := w.WriteString("\t"); err != nil {
					return StatusWriteError
				}
			}
			token := "#N/A"
			if raw != nil && ci >= 0 {
				if val := nthColumn(raw, ci); val != nil {
					token = string(val)
					if _, err := strconv.ParseFloat(token, 64); err != nil {
						token = "#VALUE!"
					}
				}
			}
			if _, err := w.WriteString(token); err != nil {
				return StatusWriteError
			}
		}
		if r < len(rows)-1 {
			if _, err := w.WriteString("\n"); err != nil {
				return StatusWriteError
			}
		}
	}
	return StatusOK
}

func projRowKeys(t *cache.Table, product, spCode, variable, simID string) ([]string, int) {
	if t == nil || t.NumDims <= 0 {
		return nil, StatusInvalid
	}
	rowDims := t.NumDims - 1
	if rowDims <= 0 || len(t.DimNames) < rowDims {
		return nil, StatusDimMismatch
	}

	row := make([]string, rowDims)
	for i := 0; i < rowDims; i++ {
		dim := strings.ToUpper(strings.TrimSpace(t.DimNames[i]))
		switch dim {
		case "NAME", "PROD_NAME", "PRODUCT", "PRODUCT_NAME":
			row[i] = strings.TrimSpace(product)
		case "SP_CODE":
			row[i] = strings.TrimSpace(spCode)
		case "VAR_NAME", "VARIABLE", "VARIABLE_NAME":
			row[i] = strings.TrimSpace(variable)
		case "SIM_ID", "SIMULATION":
			row[i] = strings.TrimSpace(simID)
			if row[i] == "" || row[i] == "*" {
				row[i] = "0"
			}
		default:
			return nil, StatusDimMismatch
		}
		if row[i] == "" || row[i] == "*" {
			return nil, StatusInvalid
		}
	}
	return row, StatusOK
}

func readCellText(t *cache.Table, row []string, col string) (string, int) {
	if len(row) != t.NumDims-1 {
		return "", StatusDimMismatch
	}
	raw := t.RawRow(row...)
	if raw == nil {
		return "", StatusNotFound
	}
	ci, ok := columnIndex(t, col)
	if !ok {
		return "", StatusNotFound
	}
	val := nthColumn(raw, ci)
	if val == nil {
		return "", StatusNotFound
	}
	return string(val), StatusOK
}

func columnIndex(t *cache.Table, colName string) (int, bool) {
	return t.ColumnIndex(colName)
}

func validParts(parts []string) bool {
	for _, part := range parts {
		if strings.TrimSpace(part) == "*" {
			return false
		}
	}
	return true
}

func nthColumn(line []byte, n int) []byte {
	if n < 0 {
		return nil
	}
	start, col := 0, 0
	for i := 0; i < len(line); i++ {
		if line[i] == ',' {
			if col == n {
				return bytes.TrimSpace(line[start:i])
			}
			col++
			start = i + 1
		}
	}
	if col == n {
		return bytes.TrimSpace(line[start:])
	}
	return nil
}

func StatusText(status int) string {
	switch status {
	case StatusOK:
		return "OK"
	case StatusNotFound:
		return "not found"
	case StatusDimMismatch:
		return "dimension mismatch"
	case StatusInvalid:
		return "invalid argument"
	case StatusParseError:
		return "parse error"
	case StatusWriteError:
		return "write error"
	default:
		return fmt.Sprintf("status %d", status)
	}
}
