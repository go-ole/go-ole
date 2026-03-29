//go:build windows

package main

import (
	"time"

	"github.com/go-ole/go-ole"
)

func main() {
	ole.Initialize(ole.Multithreaded)
	defer ole.Uninitialize()
	excelCLSID, _ := ole.LookupClassId("Excel.Application")
	excel, _ := ole.GetActiveObject[ole.IDispatch](excelCLSID)
	defer excel.Release()
	excel.PutProperty("Visible", true)
	workbooks := excel.MustGetProperty("Workbooks").ToIDispatch()
	workbook := workbooks.MustCallMethod("Add").ToIDispatch()
	worksheet := workbook.MustGetProperty("Worksheets", 1).ToIDispatch()
	cell := worksheet.MustGetProperty("Cells", 1, 1).ToIDispatch()
	cell.PutProperty("Value", 12345)

	time.Sleep(2000000000)

	workbook.PutProperty("Saved", true)
	workbook.CallMethod("Close", false)
	excel.CallMethod("Quit")
}
