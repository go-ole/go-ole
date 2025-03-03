//go:build windows

package main

import (
	"time"

	"github.com/go-ole/go-ole"
)

func main() {
	ole.Initialize()
	defer ole.Uninitialize()
	excelCLSID, _ := ole.LookupClassId("Excel.Application")
	excel, _ := ole.GetActiveObject[ole.IDispatch](excelCLSID)
	ole.PutProperty(excel, "Visible", true)
	workbooks := ole.MustGetProperty(excel, "Workbooks").ToIDispatch()
	workbook := ole.MustCallMethod(workbooks, "Add", nil).ToIDispatch()
	worksheet := ole.MustGetProperty(workbook, "Worksheets", 1).ToIDispatch()
	cell := ole.MustGetProperty(worksheet, "Cells", 1, 1).ToIDispatch()
	ole.PutProperty(cell, "Value", 12345)

	time.Sleep(2000000000)

	ole.PutProperty(workbook, "Saved", true)
	ole.CallMethod(workbook, "Close", false)
	ole.CallMethod(excel, "Quit")
	excel.Release()
}
