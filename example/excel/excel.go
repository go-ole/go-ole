package main

import "ole"
import "ole/oleutil"
import "syscall"

func main() {
	ole.CoInitialize(0)
	excel, _ := oleutil.CreateDispatch("Excel.Application")
	oleutil.PutProperty(excel, "Visible", true)
	result, _ := oleutil.GetProperty(excel, "Workbooks")
	workbooks := result.ToIDispatch()
	result, _ = oleutil.CallMethod(workbooks, "Add", nil)
	workbook := result.ToIDispatch()
	result, _ = oleutil.GetProperty(workbook, "Worksheets", 1)
	worksheet := result.ToIDispatch()
	result, _ = oleutil.GetProperty(worksheet, "Cells", 1, 1)
	cell := result.ToIDispatch()
	oleutil.PutProperty(cell, "Value", 12345)

	syscall.Sleep(2000000000)

	oleutil.PutProperty(workbook, "Saved", true)
	oleutil.CallMethod(excel, "Quit")
	excel.Release()
}
