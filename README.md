# Excelizero

Excelizero is simple encapsulation of [Excelize](https://github.com/qax-os/excelize).
It is used for insert go objects into excel files.

## Example

### 1. Write Single Object

```go
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
err := excelizero.WriteStruct("Sheet1", "A1", u)
err = excelizero.SaveAs("test.xlsx")
if err != nil {
    fmt.Println(err)
}
```

### 2.  Write Multiple Object

```go
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
// insert slice of structs
err1 := excelizero.WriteStructs("Sheet1", "A1", l)
// insert slice of pointers to struct
err2 := excelizero.WriteStructs("Sheet1", "A1", &l)
// insert pointer to slice of structs
err3 := excelizero.WriteStructs("Sheet1", "A1", l2)
// insert pointer to slice of pointers to structs
err4 := excelizero.WriteStructs("Sheet1", "A1", &l2)
err = excelizero.SaveAs("test.xlsx")
if err != nil {
    fmt.Println(err)
}
```

### 3. Write Multiple Object With Tag
```go
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
// insert slice of structs
err1 := excelizero.WriteStructs("Sheet1", "A1", l)
// insert slice of pointers to struct
err2 := excelizero.WriteStructs("Sheet1", "A1", &l)
// insert pointer to slice of structs
err3 := excelizero.WriteStructs("Sheet1", "A1", l2)
// insert pointer to slice of pointers to structs
err4 := excelizero.WriteStructs("Sheet1", "A1", &l2)
err = excelizero.SaveAs("test.xlsx")
if err != nil {
fmt.Println(err)
}
```