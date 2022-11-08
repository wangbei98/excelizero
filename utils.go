package excelizero

import (
	"errors"
	strconv2 "github.com/savsgio/gotils/strconv"
	"regexp"
	"strconv"
)

// A2 -> A3
func GetNextAxis(cur string) (string, error) {
	valid := regexp.MustCompile("[0-9]+$")
	r := valid.FindAllString(cur, -1)
	if len(r) != 1 {
		return "", errors.New("Get Row Name Fail")
	}
	c := r[0]
	column, err := strconv.Atoi(c)
	if err != nil {
		return "", err
	}
	column += 1
	row0 := cur[:len(cur)-len(c)]
	row := strconv2.S2B(row0)
	next0 := strconv.AppendInt(row, int64(column), 10)
	next := strconv2.B2S(next0)
	return next, nil
}

func GetNextRow(from, cur int) string {
	next := from + cur
	return strconv.Itoa(next)
}

func GetColumn(col string) (string, error) {
	nCol, err := strconv.Atoi(col)
	if err != nil {
		return "", err
	}
	var c []byte
	const A = 'A' - 1
	for nCol > 0 {
		cc := nCol % 26
		c = append([]byte{byte(A + cc)}, c...)
		nCol = nCol / 26
	}
	return strconv2.B2S(c), nil
}
