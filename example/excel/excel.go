// +build windows

package main

import (
	"fmt"
	"log"
	"os"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func writeExample(workbooks *ole.IDispatch) {

	workbook := oleutil.MustCallMethod(workbooks, "Add", nil).ToIDispatch()
	worksheet := oleutil.MustGetProperty(workbook, "Worksheets", 1).ToIDispatch()
	cell := oleutil.MustGetProperty(worksheet, "Cells", 1, 1).ToIDispatch()
	oleutil.PutProperty(cell, "Value", 12345)

	time.Sleep(2 * time.Second)

	// let excel could close without asking
	// oleutil.PutProperty(workbook, "Saved", true)
	//oleutil.CallMethod(workbook, "Close", false)
}
func readExample(fileName string, workbooks *ole.IDispatch) {
	workbook, err := oleutil.CallMethod(workbooks, "Open", fileName)
	if err != nil {
		log.Fatalln(err)
	}
	worksheet := oleutil.MustGetProperty(workbook.ToIDispatch(), "Worksheets", 1).ToIDispatch()
	for row := 1; row <= 2; row++ {
		for col := 1; col <= 5; col++ {
			cell := oleutil.MustGetProperty(worksheet, "Cells", row, col).ToIDispatch()
			val, err := oleutil.GetProperty(cell, "Value")
			if err != nil {
				break
			}
			fmt.Printf("(%d,%d)=%+v toString=%s\n", col, row, val.Value(), val.ToString())
		}
	}
}

func main() {
	ole.CoInitialize(0)
	unknown, _ := oleutil.CreateObject("Excel.Application")
	excel, _ := unknown.QueryInterface(ole.IID_IDispatch)
	oleutil.PutProperty(excel, "Visible", true)
	workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
	cwd, _ := os.Getwd()
	writeExample(workbooks)
	readExample(cwd+"\\excel97-2003.xls", workbooks)
	// oleutil.CallMethod(excel, "Quit")
	excel.Release()

	ole.CoUninitialize()
}
