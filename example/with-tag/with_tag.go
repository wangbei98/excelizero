package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	zero "github.com/wangbei98/excelizero"
	"github.com/wangbei98/excelizero/example"
)

func main() {
	var err error
	u := example.UserWithTag{
		Name:   "Jack",
		Age:    11,
		Height: 180,
	}
	u2 := example.UserWithTag{
		Name:   "Jack2",
		Age:    222,
		Height: 190,
	}
	var l []example.UserWithTag
	l = append(l, u, u2)

	f := excelize.NewFile()
	excelizero := zero.NewExcelizero(f)
	// insert slice of structs
	err = excelizero.WriteStructWithTag("Sheet1", 1, l)
	if err != nil {
		fmt.Println(err)
	}
	err = excelizero.SaveAs("example/with-tag/multi_with_tag.xlsx")
	if err != nil {
		fmt.Println(err)
	}

	// ################################

	uu := example.UserWithTag2{
		Name:   "Jack",
		Age:    11,
		Height: 180,
	}
	uu2 := example.UserWithTag2{
		Name:   "Jack2",
		Age:    222,
		Height: 190,
	}
	var ll []example.UserWithTag2
	ll = append(ll, uu, uu2)

	ff := excelize.NewFile()
	excelizero2 := zero.NewExcelizero(ff)
	// insert slice of structs
	err = excelizero2.WriteStructWithTag("Sheet1", 1, ll)
	if err != nil {
		fmt.Println(err)
	}
	err = excelizero2.SaveAs("example/with-tag/multi_with_tag2.xlsx")
	if err != nil {
		fmt.Println(err)
	}

}
