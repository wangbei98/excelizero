package excelizero

import (
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize"
	"reflect"
)

type Excelizero struct {
	f *excelize.File
}

func NewExcelizero(f *excelize.File) *Excelizero {
	return &Excelizero{
		f: f,
	}
}

func (x *Excelizero) WriteHeader(sheetName string, hs []string) error {
	// notice: GetSheetIndex returns different value
	// in different version of excelize
	// current version: excelize@v1.4.1
	if idx := x.f.GetSheetIndex(sheetName); idx == 0 {
		return errors.New("sheet not exist")
	}
	data := make([]interface{}, len(hs))
	for i, s := range hs {
		data[i] = s
	}
	x.f.SetSheetRow(sheetName, "A1", &data)
	return nil
}

// WriteStruct Writes a struct to row
// f. a pointer to excelizero object
// sheetName. sheetName to write to, if sheet not exist, return error
// axis. where to write from
// obj. struct need to be written to excel
func (x *Excelizero) WriteStruct(sheetName string, axis string, obj interface{}) error {
	// notice: GetSheetIndex returns different value
	// in different version of excelize
	// current version: excelize@v1.4.1
	if idx := x.f.GetSheetIndex(sheetName); idx == 0 {
		return errors.New("sheet not exist")
	}

	t := reflect.TypeOf(obj)
	v := reflect.ValueOf(obj)
	if v.Kind() == reflect.Ptr {
		t = t.Elem()
		v = v.Elem()
	}

	if t.Kind() != reflect.Struct {
		return errors.New("invalid input type, need struct or ptr to struct")
	}

	n := v.NumField()
	data := make([]interface{}, n)
	for i := 0; i < n; i++ {
		f := v.Field(i).Interface()
		data[i] = f
	}
	x.f.SetSheetRow(sheetName, axis, &data)
	return nil
}

// WriteStructs Writes structs to rows
// f. a pointer to excelizero object
// sheetName. sheetName to write to, if sheet not exist, return error
// axis. where to write from
// sl. slice of struct that need to be written to excel
func (x *Excelizero) WriteStructs(sheetName string, axis string, sl interface{}) error {
	// notice: GetSheetIndex returns different value
	// in different version of excelize
	// current version: excelize@v1.4.1
	var err error
	if idx := x.f.GetSheetIndex(sheetName); idx == 0 {
		return errors.New("sheet not exist")
	}

	t := reflect.TypeOf(sl)
	v := reflect.ValueOf(sl)
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
		v = v.Elem()
	}
	if t.Kind() != reflect.Slice {
		return errors.New("sl not Slice")
	}
	var next = axis
	for i := 0; i < v.Len(); i++ {
		if i > 0 {
			next, err = GetNextAxis(next)
			if err != nil {
				return nil
			}
		}
		err = x.WriteStruct(sheetName, next, v.Index(i).Interface())
		if err != nil {
			return err
		}
	}
	return nil
}

// WriteStructWithTag: write structs to excel rows
// only insert field that has tag: xlsx
// sheetName. sheetName to write to, if sheet not exist, return error
// fromRow. no of row where to start to insert
// sl. slice of struct that need to be written to excel
func (x *Excelizero) WriteStructWithTag(sheetName string, fromRow int, sl interface{}) error {
	// notice: GetSheetIndex returns different value
	// in different version of excelize
	// current version: excelize@v1.4.1
	if idx := x.f.GetSheetIndex(sheetName); idx == 0 {
		return errors.New("sheet not exist")
	}

	t := reflect.TypeOf(sl)
	v := reflect.ValueOf(sl)
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
		v = v.Elem()
	}
	if t.Kind() != reflect.Slice {
		return errors.New("sl not Slice")
	}
	if v.Len() < 1 {
		return nil
	}

	fieldTagMap := make(map[string]string)
	tt := v.Index(0).Type()
	if tt.Kind() == reflect.Ptr {
		tt = tt.Elem()
	}
	for i := 0; i < tt.NumField(); i++ {
		tag := tt.Field(i).Tag.Get("xlsx")
		if tag == "" || tag == "-" {
			continue
		}
		col, err := GetColumn(tag)
		if err != nil {
			continue
		}
		if _, ok := fieldTagMap[tag]; ok {
			return errors.New("duplicate tag")
		} else {
			fieldTagMap[tag] = col
		}
	}
	for i := 0; i < v.Len(); i++ {
		row := GetNextRow(fromRow, i)
		vv := v.Index(i)
		if vv.Kind() == reflect.Ptr {
			vv = vv.Elem()
		}
		for j := 0; j < tt.NumField(); j++ {
			fv := vv.Field(j).Interface()
			tag := tt.Field(j).Tag.Get("xlsx")
			if _, ok := fieldTagMap[tag]; !ok {
				return errors.New("tag not exists")
			}
			col := fieldTagMap[tag]
			axis := col + row
			x.f.SetCellValue(sheetName, axis, fv)
		}
	}
	return nil
}

func (x *Excelizero) SaveAs(name string) error {
	err := x.f.SaveAs(name)
	if err != nil {
		return err
	}
	return nil
}
