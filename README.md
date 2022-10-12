# Exceler

This package is an implementation of github.com/qax-os/excelize

Provides an easy way to create an excel file.

## Usage

### Installation

```bash
go get github.com/ybalcin/exceler
```

### Create Excel

```go
package main

import (
	"bytes"
	"fmt"
	"github.com/ybalcin/exceler"
	"os"
)

func main() {
	// Create excel file
	f := exceler.New("test")

	sheet1 := exceler.NewSheet("sheet1")
	sheet1.AddColumn(
		exceler.NewColumn("header1"),
		exceler.NewColumn("header2"),
	)
	sheet1.AddRow(*exceler.NewRow(
		exceler.NewCell("cell1"),
		exceler.NewCell("cell2"),
	))

	sheet2 := exceler.NewSheet("sheet2")
	sheet2.AddColumn(
		exceler.NewColumn("header3"),
		exceler.NewColumn("header4"),
	)
	sheet2.AddRow(*exceler.NewRow(
		exceler.NewCell("cell3"),
		exceler.NewCell("cell4"),
	))

	f.AddSheet(*sheet1)
	f.AddSheet(*sheet2)

	path, _ := os.Getwd()

	if err := f.Save(path); err != nil {
		fmt.Println(err)
	}

	// Read Excel file from path
	bs, err := os.ReadFile(fmt.Sprintf("%s/%s", path, "test.xlsx"))
	if err != nil {
		panic(err)
	}

	bf := bytes.NewBuffer(bs)
	ff, err := exceler.ReadFromBuffer(bf, 0, 0)
	if err != nil {
		panic(err)
	}

	ss := ff.Sheets()
	for _, v := range ss {
		// get rows of sheet
		rows := v.Rows()
		for _, r := range rows {
			// get cells of row
			cells := r.Cells()
			for _, c := range cells {
				// get cell value
				fmt.Println(c.Value())
			}
		}
	}
}
```
