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
	File struct {
		Name      string
		extension string
		sheets    []sheet
	}

	sheet struct {
		name    string
		columns []column
		rows    []row
	}

	column struct {
		name string
	}

	row struct {
		cells []cell
	}

	cell struct {
		value interface{}
		name  string
	}
)

const MaxCellIndexConst = 26
const MaxRowCountConst = 10_000

// New initializes new File
func New(name string) *File {
	if name == "" {
		name = uuid.NewString()
	}

	return &File{
		Name:      name,
		extension: "xlsx",
	}
}

func NewSheet(name string) *sheet {
	return &sheet{name: name}
}

func NewRow(cells ...cell) *row {
	row := row{}
	row.cells = make([]cell, len(cells))
	for i, c := range cells {
		row.cells[i] = c
	}

	return &row
}

func NewColumn(name string) column {
	return column{name: name}
}

func NewCell(value interface{}) cell {
	return cell{value: value}
}

func (f *File) AddSheet(sheet sheet) {
	f.sheets = append(f.sheets, sheet)
}

func (f *File) Sheets() []sheet {
	return f.sheets
}

func (s *sheet) Name() string {
	return s.name
}

func (s *sheet) Rows() []row {
	return s.rows
}

func (s *sheet) Columns() []column {
	return s.columns
}

func (s column) Name() string {
	return s.name
}

func (r row) Cells() []cell {
	return r.cells
}

func (c cell) Value() interface{} {
	return c.value
}

func (c cell) Name() interface{} {
	return c.name
}

func (c *cell) SetName(name string) {
	c.name = name
}

func (s *sheet) AddRow(row row) {
	s.rows = append(s.rows, row)
}

func (s *sheet) AddColumn(columns ...column) {
	s.columns = append(s.columns, columns...)
}

func (r *row) AddCell(cell ...cell) {
	r.cells = append(r.cells, cell...)
}

// ToBuffer writes Excel File to the buffer
func (f *File) ToBuffer() (*bytes.Buffer, error) {
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

// Save saves File to the location
func (f *File) Save(location string) error {
	excelFile, err := fill(f)
	if err != nil {
		return err
	}

	if err = excelFile.SaveAs(fmt.Sprintf("%s/%s", location, f.Name)); err != nil {
		return err
	}

	return nil
}

func fill(f *File) (*excelize.File, error) {
	f.Name = fmt.Sprintf("%s.%s", f.Name, f.extension)

	excelFile := excelize.NewFile()
	defer excelFile.Close()

	if len(f.sheets) < 1 {
		return nil, errors.New("there is no any sheet in the File")
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
			// set columns
			for j, col := range s.columns {
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

/*
ReadFromBuffer reads from buffer and returns *File

Params:

	maxRowCount: the number of row that will read from every sheet, sets as MaxRowCountConst if 0
	maxARowsCellCount: the number of cell that will read from every row, sets as MaxCellIndexConst if 0
*/
func ReadFromBuffer(bf *bytes.Buffer, maxRowCount, maxARowsCellCount int) (*File, error) {
	if maxRowCount <= 0 {
		maxRowCount = MaxRowCountConst
	}
	if maxARowsCellCount <= 0 || maxARowsCellCount > MaxCellIndexConst {
		maxARowsCellCount = MaxCellIndexConst
	}

	f, err := excelize.OpenReader(bf)
	if err != nil {
		return nil, err
	}

	file := New("")

	sheets := f.GetSheetList()

	sheetCh := make(chan *sheet, len(sheets))

	var wg sync.WaitGroup
	wg.Add(len(sheets))

	for _, s := range sheets {
		s := s

		go func(ch chan<- *sheet) {
			sh := NewSheet(s)

			for i := 1; i <= maxRowCount; i++ { // row
				for j := 1; j <= maxARowsCellCount; j++ { // cell
					// set column names
					if i == 1 {
						columnName, err := f.GetCellValue(s, fmt.Sprintf("%s%d", getColumn(j-1), i))
						if err == nil && len([]rune(columnName)) > 0 {
							sh.AddColumn(NewColumn(columnName))
						}
					}

					// set cells
					cellName := fmt.Sprintf("%s%d", getColumn(j-1), i+1)
					val, err := f.GetCellValue(s, cellName)
					if err != nil || len([]rune(val)) <= 0 {
						continue
					}
					cell := NewCell(val)
					cell.SetName(cellName)
					sh.AddRow(*NewRow(
						cell,
					))
				}
			}

			ch <- sh
			wg.Done()
		}(sheetCh)
	}

	go func() {
		wg.Wait()
		close(sheetCh)
	}()

	for s := range sheetCh {
		if len(s.rows) > 0 {
			file.AddSheet(*s)
		}
	}

	return file, nil
}
