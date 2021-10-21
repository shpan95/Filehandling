package main

import (
	"flag"
	"fmt"
	"os"
	"strings"

	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
)

func main() {
	fp := flag.String("filepath", "input_file.xlsx", "location of the input file (Required)")
	flag.Parse()

	var sh1 = "Sheet1"
	var sh2 = "Sheet2"

	if *fp == "" {
		flag.PrintDefaults()
		os.Exit(1)
	}

	m := getHashMap(*fp, sh2)
	createOutputFile(*fp, sh1, m)

}

func getHashMap(fp string, sheet string) map[string]string {
	f, err := excelize.OpenFile(fp)
	if err != nil {
		fmt.Println(err)
		return nil
	}

	m := make(map[string]string)

	// Get all the rows in the Sheet1 as an array
	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println(err)
		return nil
	}

	for _, row := range rows {
		//assuming the row will only have 2 columns in a predefined manner
		var key = row[0]
		var value = row[1]
		m[key] = value
	}

	//remove empty strings from the map, if present
	for k := range m {
		if k == "" {
			delete(m, k)
		}

	}

	//fmt.Println(m)
	return m
}

func createOutputFile(fp string, sheet string, m map[string]string) {

	f, err := excelize.OpenFile(fp)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println(err)
		return
	}

	//set start index
	l := len(rows[0]) - 1
	ax, _ := f.SearchSheet(sheet, rows[0][l])
	initAxis := strings.Split(ax[0], "")
	var startIndex = ascii2Char(char2Ascii(initAxis[0]) + 2)

	var rowNo = 0
	for _, row := range rows {
		links := make(map[string]string)
		rowNo++
		var currColIndex = startIndex
		rowElem := strings.FieldsFunc(row[2], func(r rune) bool { return strings.ContainsRune(" -,:.!§$%&/=?``", r) })

		//get all keyword-link in this row
		for _, elem := range rowElem {
			if val, ok := m[elem]; ok {
				links[elem] = val

			}
		}

		for k, v := range links {
			var nexIndx = getNextColIndex(currColIndex)
			var index = fmt.Sprintf("%s%d", nexIndx, rowNo)

			f.SetCellValue(sheet, index, k)
			style, _ := f.NewStyle(`{"font":{"color":"blue","underline":"single"}}`)
			f.SetCellStyle(sheet, index, index, style)

			f.SetCellHyperLink(sheet, index, v, "External")

			currColIndex = nexIndx
		}

	}

	// Save spreadsheet by the given path.
	if err := f.SaveAs("Output.xlsx"); err != nil {
		fmt.Println(err)
	}

}

func getNextColIndex(colNo string) string {
	return ascii2Char(char2Ascii(colNo) + 1)

}

func ascii2Char(ascii int) string {
	ch := rune(ascii)
	return string(ch)
}

func char2Ascii(str string) int {
	r := []rune(str)
	return int(r[0])
}
