package exceler

import (
	"bytes"
	"errors"
	"fmt"
	"github.com/google/uuid"
	"github.com/xuri/excelize/v2"
	"sync"
	"time"
)

type (
	file struct {
		Name      string
		extension string
		sheets    []sheet
	}

	sheet struct {
		name    string
		headers []header
		rows    []row
	}

	header struct {
		name string
	}

	row struct {
		cells []cell
	}

	cell struct {
		value interface{}
	}
)

// New initializes new File
func New(name string) *file {
	if name == "" {
		name = uuid.NewString()
	}

	return &file{
		Name:      name,
		extension: "xlsx",
	}
}

// NewSheet initializes new sheet
func NewSheet(name string) *sheet {
	return &sheet{name: name}
}

// NewRow initializes new row
func NewRow(cells ...cell) *row {
	row := row{}
	row.cells = make([]cell, len(cells))
	for i, c := range cells {
		row.cells[i] = c
	}

	return &row
}

// NewHeader initializes new header
func NewHeader(name string) header {
	return header{name: name}
}

// NewCell initializes new cell
func NewCell(value interface{}) cell {
	return cell{value: value}
}

// AddSheet adds sheet to the file
func (f *file) AddSheet(sheet sheet) {
	f.sheets = append(f.sheets, sheet)
}

// AddRow adds row to the file
func (s *sheet) AddRow(row row) {
	s.rows = append(s.rows, row)
}

// AddHeader adds header to the file
func (s *sheet) AddHeader(headers ...header) {
	s.headers = append(s.headers, headers...)
}

// AddCell adds cell to the file
func (r *row) AddCell(cell ...cell) {
	r.cells = append(r.cells, cell...)
}

// ToBuffer writes Excel file to the buffer
func (f *file) ToBuffer() (*bytes.Buffer, error) {
	excelFile, err := fill(f)
	if err != nil {
		return nil, err
	}

	buffer, err := excelFile.WriteToBuffer()
	if err != nil {
		return nil, err
	}

	return buffer, nil
}

// Save saves file to the location
func (f *file) Save(location string) error {
	excelFile, err := fill(f)
	if err != nil {
		return err
	}

	if err = excelFile.SaveAs(fmt.Sprintf("%s/%s", location, f.Name)); err != nil {
		return err
	}

	return nil
}

func fill(f *file) (*excelize.File, error) {
	f.Name = fmt.Sprintf("%s.%s", f.Name, f.extension)

	excelFile := excelize.NewFile()
	defer excelFile.Close()

	if len(f.sheets) < 1 {
		return nil, errors.New("there is no any sheet in the file")
	}

	rndSheetName := uuid.NewString()
	excelFile.NewSheet(rndSheetName)
	excelFile.DeleteSheet("Sheet1")

	headerStyle, _ := excelFile.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
			Size: 17,
		},
	})
	bodyStyle, _ := excelFile.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Size: 14,
		},
	})

	var wg sync.WaitGroup

	for i, s := range f.sheets {
		s := s

		if len(s.rows) <= 0 {
			continue
		}

		index := excelFile.NewSheet(s.name)
		if i == 0 {
			excelFile.SetActiveSheet(index)
		}

		wg.Add(1)
		go func() {
			// set headers
			for j, col := range s.headers {
				cell := fmt.Sprintf("%s1", getColumn(j))
				_ = excelFile.SetCellValue(s.name, cell, col.name)
				_ = excelFile.SetCellStyle(s.name, cell, cell, headerStyle)
			}

			// fill the rows
			for j, r := range s.rows {
				for k, c := range r.cells {
					cell := fmt.Sprintf("%s%d", getColumn(k), j+2)
					_ = excelFile.SetCellValue(s.name, cell, getString(c.value))
					_ = excelFile.SetCellStyle(s.name, cell, cell, bodyStyle)
				}
			}

			wg.Done()
		}()
	}

	wg.Wait()

	excelFile.DeleteSheet(rndSheetName)

	return excelFile, nil
}

func getColumn(index int) string {
	// TODO: Fix for index with a greater number than column names

	//const num = 26

	columnNames := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	return columnNames[index]
}

func getString(val interface{}) string {
	if val == nil || val == 0 || val == float64(0) || val == float32(0) || val == int64(0) || val == "" {
		return "-"
	}

	var nilTime *time.Time
	if val == nilTime {
		return "-"
	}

	return fmt.Sprintf("%v", val)
}
