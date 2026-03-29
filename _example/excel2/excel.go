//go:build windows
// +build windows

package main

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"log"
	"os"
)

func writeExample(excel, workbooks *ole.IDispatch, filepath string) {
	// ref: https://msdn.microsoft.com/zh-tw/library/office/ff198017.aspx
	// http://stackoverflow.com/questions/12159513/what-is-the-correct-xlfileformat-enumeration-for-excel-97-2003
	const xlExcel8 = 56
	workbook := ole.VariantToComObject[ole.IDispatch](workbooks.MustCallMethod("Add"))
	defer workbook.Release()
	worksheet := ole.VariantToComObject[ole.IDispatch](workbook.MustGetProperty("Worksheets", ole.Int32ToVariant(1)))
	defer worksheet.Release()
	cell := ole.VariantToComObject[ole.IDispatch](worksheet.MustGetProperty("Cells", ole.Int32ToVariant(1), ole.Int32ToVariant(1)))
	cell.PutProperty("Value", ole.Int32ToVariant(12345))
	cell.Release()
	activeWorkBook := ole.VariantToComObject[ole.IDispatch](excel.MustGetProperty("ActiveWorkBook"))
	defer activeWorkBook.Release()

	os.Remove(filepath)
	// ref: https://msdn.microsoft.com/zh-tw/library/microsoft.office.tools.excel.workbook.saveas.aspx
	activeWorkBook.MustCallMethod("SaveAs", ole.StringToBStrVariant(filepath), ole.Int32ToVariant(xlExcel8))
}

func readExample(fileName string, excel, workbooks *ole.IDispatch) {
	workbook, err := workbooks.CallMethod("Open", ole.StringToBStrVariant(fileName))

	if err != nil {
		log.Fatalln(err)
	}
	workbookDispatch := ole.VariantToComObject[ole.IDispatch](workbook)
	defer workbookDispatch.Release()

	sheets := ole.VariantToComObject[ole.IDispatch](excel.MustGetProperty("Sheets"))
	sheetCount := int(ole.UnwrapVariant[int32](sheets.MustGetProperty("Count")))
	fmt.Println("sheet count=", sheetCount)
	sheets.Release()

	worksheet := ole.VariantToComObject[ole.IDispatch](workbookDispatch.MustGetProperty("Worksheets", ole.Int32ToVariant(1)))
	defer worksheet.Release()
	for row := 1; row <= 2; row++ {
		for col := 1; col <= 5; col++ {
			cell := ole.VariantToComObject[ole.IDispatch](worksheet.MustGetProperty("Cells", ole.Int32ToVariant(row), ole.Int32ToVariant(col)))
			val, err := cell.GetProperty("Value")
			if err != nil {
				cell.Release()
				break
			}
			fmt.Printf("(%d,%d)=%+v\n", col, row, val)
			cell.Release()
		}
	}
}

func showMethodsAndProperties(i *ole.IDispatch) {
	if !i.HasTypeInfo() {
		return
	}

	typeInfo := i.GetTypeInfo()
	if typeInfo == nil {
		return
	}

	fmt.Println("typeInfo=", typeInfo)
}

func main() {
	log.SetFlags(log.Flags() | log.Lshortfile)
	ole.InitializeMultithreaded()
	defer ole.Uninitialize()
	ole.RegisterVariantConverters()

	clsid, err := ole.ClassIdFromString("Excel.Application")
	if err != nil {
		panic("Unable to get class ID for 'Excel.Application'")
	}

	excel, _ := ole.GetActiveObject[ole.IDispatch](clsid, ole.IID_IDispatch)
	defer excel.Release()

	excel.PutProperty("Visible", ole.BoolToVariant(true))

	workbooks := ole.VariantToComObject[ole.IDispatch](excel.MustGetProperty("Workbooks"))
	defer workbooks.Release()

	cwd, _ := os.Getwd()
	writeExample(excel, workbooks, cwd+"\\write.xls")
	readExample(cwd+"\\excel97-2003.xls", excel, workbooks)
	showMethodsAndProperties(workbooks)
}
