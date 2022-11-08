package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	zero "github.com/wangbei98/excelizero"
	"github.com/wangbei98/excelizero/example"
)

func main() {
	var err error
	u := example.User{
		Name:   "Jack",
		Age:    11,
		Height: 180,
	}
	u2 := example.User{
		Name:   "Jack2",
		Age:    222,
		Height: 190,
	}
	u3 := &example.User{
		Name:   "Jack",
		Age:    11,
		Height: 180,
	}
	u4 := &example.User{
		Name:   "Jack2",
		Age:    222,
		Height: 190,
	}
	var l []example.User
	var l2 []*example.User
	l = append(l, u, u2)
	l2 = append(l2, u3, u4)

	f := excelize.NewFile()
	excelizero := zero.NewExcelizero(f)
	// insert slice of structs
	err = excelizero.WriteStructs("Sheet1", "A1", l)
	if err != nil {
		fmt.Println(err)
	}
	err = excelizero.SaveAs("example/multi/multi.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
