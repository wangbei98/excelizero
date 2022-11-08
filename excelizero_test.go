package excelizero

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/stretchr/testify/assert"
	"testing"
)

func TestWriteHeader(t *testing.T) {
	f := excelize.NewFile()
	excelizero := NewExcelizero(f)
	err := excelizero.WriteHeader("Sheet1", []string{"Name", "Age", "Height"})

	assert.Equal(t, nil, err)
}

func TestWriteStruct(t *testing.T) {
	type User struct {
		Name   string
		Age    int
		Height int
	}
	u := User{
		Name:   "Jack",
		Age:    11,
		Height: 180,
	}

	f := excelize.NewFile()
	excelizero := NewExcelizero(f)
	err1 := excelizero.WriteStruct("Sheet1", "A1", u)
	err2 := excelizero.WriteStruct("Sheet1", "A1", &u)
	assert.Equal(t, nil, err1)
	assert.Equal(t, nil, err2)
}

func TestWriteStructs(t *testing.T) {
	type User struct {
		Name   string
		Age    int
		Height int
	}
	u := User{
		Name:   "Jack",
		Age:    11,
		Height: 180,
	}
	u2 := User{
		Name:   "Jack2",
		Age:    222,
		Height: 190,
	}
	u3 := &User{
		Name:   "Jack",
		Age:    11,
		Height: 180,
	}
	u4 := &User{
		Name:   "Jack2",
		Age:    222,
		Height: 190,
	}
	var l []User
	var l2 []*User
	l = append(l, u, u2)
	l2 = append(l2, u3, u4)
	f := excelize.NewFile()
	excelizero := NewExcelizero(f)
	err1 := excelizero.WriteStructs("Sheet1", "A1", l)
	err2 := excelizero.WriteStructs("Sheet1", "A1", &l)
	err3 := excelizero.WriteStructs("Sheet1", "A1", l2)
	err4 := excelizero.WriteStructs("Sheet1", "A1", &l2)
	assert.Equal(t, nil, err1)
	assert.Equal(t, nil, err2)
	assert.Equal(t, nil, err3)
	assert.Equal(t, nil, err4)
}

func TestWriteStructWithTag(t *testing.T) {
	type User struct {
		Name   string `xlsx:"1"`
		Age    int    `xlsx:"2"`
		Height int    `xlsx:"3"`
	}
	u := User{
		Name:   "Jack",
		Age:    11,
		Height: 180,
	}
	u2 := User{
		Name:   "Jack2",
		Age:    222,
		Height: 190,
	}
	u3 := &User{
		Name:   "Jack",
		Age:    11,
		Height: 180,
	}
	u4 := &User{
		Name:   "Jack2",
		Age:    222,
		Height: 190,
	}
	var l []User
	var l2 []*User
	l = append(l, u, u2)
	l2 = append(l2, u3, u4)
	f := excelize.NewFile()
	excelizero := NewExcelizero(f)
	err1 := excelizero.WriteStructWithTag("Sheet1", 1, l)
	err2 := excelizero.WriteStructWithTag("Sheet1", 1, &l)
	err3 := excelizero.WriteStructWithTag("Sheet1", 1, l2)
	err4 := excelizero.WriteStructWithTag("Sheet1", 1, &l2)
	assert.Equal(t, nil, err1)
	assert.Equal(t, nil, err2)
	assert.Equal(t, nil, err3)
	assert.Equal(t, nil, err4)
}
