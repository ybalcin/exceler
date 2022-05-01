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
	"fmt"
	"github.com/ybalcin/exceler"
	"os"
)

func main() {
	f := exceler.New("test")

	sheet1 := exceler.NewSheet("sheet1")
	sheet1.AddHeader(
		exceler.NewHeader("header1"),
		exceler.NewHeader("header2"),
	)
	sheet1.AddRow(*exceler.NewRow(
		exceler.NewCell("cell1"),
		exceler.NewCell("cell2"),
	))

	sheet2 := exceler.NewSheet("sheet2")
	sheet2.AddHeader(
		exceler.NewHeader("header3"),
		exceler.NewHeader("header4"),
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
}
```
