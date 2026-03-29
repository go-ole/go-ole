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
	workbook := workbooks.MustCallMethod("Add", nil).ToIDispatch()
	defer workbook.Release()
	worksheet := workbook.MustGetProperty("Worksheets", 1).ToIDispatch()
	defer worksheet.Release()
	cell := worksheet.MustGetProperty("Cells", 1, 1).ToIDispatch()
	cell.PutProperty("Value", 12345)
	cell.Release()
	activeWorkBook := excel.MustGetProperty("ActiveWorkBook").ToIDispatch()
	defer activeWorkBook.Release()

	os.Remove(filepath)
	// ref: https://msdn.microsoft.com/zh-tw/library/microsoft.office.tools.excel.workbook.saveas.aspx
	activeWorkBook.MustCallMethod("SaveAs", filepath, xlExcel8, nil, nil).ToIDispatch()

	//time.Sleep(2 * time.Second)

	// let excel could close without asking
	// workbook.PutProperty("Saved", true)
	// workbook.CallMethod("Close", false)
}

func readExample(fileName string, excel, workbooks *ole.IDispatch) {
	workbook, err := workbooks.CallMethod("Open", fileName)

	if err != nil {
		log.Fatalln(err)
	}
	defer workbook.ToIDispatch().Release()

	sheets := excel.MustGetProperty("Sheets").ToIDispatch()
	sheetCount := (int)(sheets.MustGetProperty("Count").Val)
	fmt.Println("sheet count=", sheetCount)
	sheets.Release()

	worksheet := workbook.ToIDispatch().MustGetProperty("Worksheets", 1).ToIDispatch()
	defer worksheet.Release()
	for row := 1; row <= 2; row++ {
		for col := 1; col <= 5; col++ {
			cell := worksheet.MustGetProperty("Cells", row, col).ToIDispatch()
			val, err := cell.GetProperty("Value")
			if err != nil {
				break
			}
			fmt.Printf("(%d,%d)=%+v toString=%s\n", col, row, val.Value(), val.ToString())
			cell.Release()
		}
	}
}

func showMethodsAndProperties(i *ole.IDispatch) {
	n, err := i.GetTypeInfoCount()
	if err != nil {
		log.Fatalln(err)
	}
	tinfo, err := i.GetTypeInfo()
	if err != nil {
		log.Fatalln(err)
	}

	fmt.Println("n=", n, "tinfo=", tinfo)
}

func main() {
	log.SetFlags(log.Flags() | log.Lshortfile)
	ole.Initialize(ole.Multithreaded)
	defer ole.Uninitialize()
	clsid, _ := ole.LookupClassId("Excel.Application")
	unknown, _ := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	excel, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	excel.PutProperty("Visible", true)

	workbooks := excel.MustGetProperty("Workbooks").ToIDispatch()
	cwd, _ := os.Getwd()
	writeExample(excel, workbooks, cwd+"\\write.xls")
	readExample(cwd+"\\excel97-2003.xls", excel, workbooks)
	showMethodsAndProperties(workbooks)
	workbooks.Release()
	// excel.CallMethod("Quit")
	excel.Release()
}
