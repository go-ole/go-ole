//go:build windows

package main

import (
	"time"

	"github.com/go-ole/go-ole"
)

func main() {
	ole.InitializeMultithreaded()
	defer ole.Uninitialize()
	ole.RegisterVariantConverters()

	excel, err := ole.GetActiveObjectFromString[ole.IDispatch]("Excel.Application", ole.IID_IDispatch)
	if err != nil {
		panic("unable to load Excel")
	}
	defer excel.Release()

	excel.PutProperty("Visible", ole.BoolToVariant(true))
	workbooks := ole.VariantToComObject[ole.IDispatch](excel.MustGetProperty("Workbooks"))
	workbook := ole.VariantToComObject[ole.IDispatch](workbooks.CallMethod("Add"))
	worksheet := ole.VariantToComObject[ole.IDispatch](workbook.MustGetProperty("Worksheets", ole.Int64ToVariant(1)))
	cell := ole.VariantToComObject[ole.IDispatch](worksheet.MustGetProperty("Cells", ole.Int64ToVariant(1), ole.Int64ToVariant(1)))
	cell.PutProperty("Value", ole.Int64ToVariant(12345))

	time.Sleep(2000000000)

	workbook.PutProperty("Saved", ole.BoolToVariant(true))
	workbook.CallMethod("Close", ole.BoolToVariant(false))
	excel.CallMethod("Quit")
}
