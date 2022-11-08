package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	zero "github.com/wangbei98/excelizero"
	"github.com/wangbei98/excelizero/example"
)

func main() {
	u := example.User{
		Name:   "Jack",
		Age:    11,
		Height: 180,
	}

	f := excelize.NewFile()
	excelizero := zero.NewExcelizero(f)
	err := excelizero.WriteStruct("Sheet1", "A1", u)
	if err != nil {
		fmt.Println(err)
	}
	err = excelizero.SaveAs("example/single/single.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
