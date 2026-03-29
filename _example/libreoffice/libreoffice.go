//go:build windows
// +build windows

/*
	Demonstrates basic LibreOffce (OpenOffice) automation with OLE using GO-OLE.
	Usage: 	cd [...]\go-ole\example\libreoffice
			go run libreoffice.go
	References:
			http://www.openoffice.org/api/basic/man/tutorial/tutorial.pdf
			http://api.libreoffice.org/examples/examples.html#OLE_examples
			https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide

	Tested environment:
			go 1.6.2 (windows/amd64)
			LibreOffice 5.1.0.3 (32 bit)
			Windows 10 (64 bit)

	The MIT License (MIT)
	Copyright (c) 2016 Sebastian Schleemilch <https://github.com/itschleemilch>.

	Permission is hereby granted, free of charge, to any person obtaining a copy of
	this software and associated documentation files (the "Software"), to deal in
	the Software without restriction, including without limitation the rights to use,
	copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
	and to permit persons to whom the Software is furnished to do so, subject to the
	following conditions:

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
	INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
	PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
	LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
	TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
	OR OTHER DEALINGS IN THE SOFTWARE.
*/

package main

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"log"
)

func checkError(err error, msg string) {
	if err != nil {
		log.Fatal(msg)
	}
}

// LOGetCell returns an handle to a cell within a worksheet
// LibreOffice Basic: GetCell = oSheet.getCellByPosition (nColumn , nRow)
func LOGetCell(worksheet *ole.IDispatch, nColumn int, nRow int) (cell *ole.IDispatch) {
	return worksheet.MustCallMethod("getCellByPosition", nColumn, nRow).ToIDispatch()
}

// LOGetCellRangeByName returns a named range (e.g. "A1:B4")
func LOGetCellRangeByName(worksheet *ole.IDispatch, rangeName string) (cells *ole.IDispatch) {
	return worksheet.MustCallMethod("getCellRangeByName", rangeName).ToIDispatch()
}

// LOGetCellString returns the displayed value
func LOGetCellString(cell *ole.IDispatch) (value string) {
	return cell.MustGetProperty("string").ToString()
}

// LOGetCellValue returns the cell's internal value (not formatted, dummy code, FIXME)
func LOGetCellValue(cell *ole.IDispatch) (value string) {
	val := cell.MustGetProperty("value")
	fmt.Printf("Cell: %+v\n", val)
	return val.ToString()
}

// LOGetCellError returns the error value of a cell (dummy code, FIXME)
func LOGetCellError(cell *ole.IDispatch) (result *ole.VARIANT) {
	return cell.MustGetProperty("error")
}

// LOSetCellString sets the text value of a cell
func LOSetCellString(cell *ole.IDispatch, text string) {
	cell.MustPutProperty("string", text)
}

// LOSetCellValue sets the numeric value of a cell
func LOSetCellValue(cell *ole.IDispatch, value float64) {
	cell.MustPutProperty("value", value)
}

// LOSetCellFormula sets the formula (in englisch language)
func LOSetCellFormula(cell *ole.IDispatch, formula string) {
	cell.MustPutProperty("formula", formula)
}

// LOSetCellFormulaLocal sets the formula in the user's language (e.g. German =SUMME instead of =SUM)
func LOSetCellFormulaLocal(cell *ole.IDispatch, formula string) {
	cell.MustPutProperty("FormulaLocal", formula)
}

// LONewSpreadsheet creates a new spreadsheet in a new window and returns a document handle.
func LONewSpreadsheet(desktop *ole.IDispatch) (document *ole.IDispatch) {
	var args = []string{}
	document = desktop.MustCallMethod(
		"loadComponentFromURL", "private:factory/scalc", // alternative: private:factory/swriter
		"_blank", 0, args).ToIDispatch()
	return
}

// LOOpenFile opens a file (text, spreadsheet, ...) in a new window and returns a document
// handle. Example: /home/testuser/spreadsheet.ods
func LOOpenFile(desktop *ole.IDispatch, fullpath string) (document *ole.IDispatch) {
	var args = []string{}
	document = desktop.MustCallMethod(
		"loadComponentFromURL", "file://"+fullpath,
		"_blank", 0, args).ToIDispatch()
	return
}

// LOSaveFile saves the current document.
// Only works if a file already exists,
// see https://wiki.openoffice.org/wiki/Saving_a_document
func LOSaveFile(document *ole.IDispatch) {
	// use storeAsURL if neccessary with third URL parameter
	document.MustCallMethod("store")
}

// LOGetWorksheet returns a worksheet (index starts at 0)
func LOGetWorksheet(document *ole.IDispatch, index int) (worksheet *ole.IDispatch) {
	sheets := document.MustGetProperty("Sheets").ToIDispatch()
	worksheet = sheets.MustCallMethod("getByIndex", index).ToIDispatch()
	return
}

// This example creates a new spreadsheet, reads and modifies cell values and style.
func main() {
	ole.Initialize(ole.Multithreaded)
	defer ole.Uninitialize()
	clsid, err := ole.LookupClassId("com.sun.star.ServiceManager")
	checkError(err, "Couldn't create a OLE connection to LibreOffice")
	unknown, errCreate := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	checkError(errCreate, "Couldn't create a OLE connection to LibreOffice")
	ServiceManager, errSM := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	checkError(errSM, "Couldn't start a LibreOffice instance")
	desktop := ServiceManager.MustCallMethod(
		"createInstance", "com.sun.star.frame.Desktop").ToIDispatch()

	document := LONewSpreadsheet(desktop)
	sheet0 := LOGetWorksheet(document, 0)

	cell1_1 := LOGetCell(sheet0, 1, 1) // cell B2
	cell1_2 := LOGetCell(sheet0, 1, 2) // cell B3
	cell1_3 := LOGetCell(sheet0, 1, 3) // cell B4
	cell1_4 := LOGetCell(sheet0, 1, 4) // cell B5
	LOSetCellString(cell1_1, "Hello World")
	LOSetCellValue(cell1_2, 33.45)
	LOSetCellFormula(cell1_3, "=B3+5")
	b4Value := LOGetCellString(cell1_3)
	LOSetCellString(cell1_4, b4Value)
	// set background color yellow:
	cell1_1.MustPutProperty("cellbackcolor", 0xFFFF00)

	fmt.Printf("Press [ENTER] to exit")
	fmt.Scanf("%s")
	ServiceManager.Release()
}
