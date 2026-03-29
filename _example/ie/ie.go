//go:build windows
// +build windows

package main

import (
	"time"

	"github.com/go-ole/go-ole"
)

func main() {
	ole.InitializeMultithreaded()
	defer ole.Uninitialize()
	ole.RegisterVariantConverters()

	clsid, _ := ole.ClassIdFromString("InternetExplorer.Application")
	unknown, _ := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	defer unknown.Release()

	ie, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	defer ie.Release()

	ie.PutProperty("Visible", ole.BoolToVariant(true))
	ie.CallMethod("Navigate", ole.StringToBStrVariant("http://www.google.com"))
	for {
		busy := ole.UnwrapVariant[int](ie.MustGetProperty("Busy"))
		if busy == 0 {
			break
		}
	}

	time.Sleep(1e9)

	document := ole.VariantToComObject[ole.IDispatch](ie.MustGetProperty("document"))

	// set 'golang' to text box.
	elems := ole.VariantToComObject[ole.IDispatch](document.MustCallMethod("getElementsByName", ole.StringToBStrVariant("q")))
	q := ole.VariantToComObject[ole.IDispatch](elems.MustCallMethod("item", ole.Int32ToVariant(0)))
	q.MustPutProperty("value", ole.StringToBStrVariant("golang"))

	// click btnK.
	elems = ole.VariantToComObject[ole.IDispatch](document.MustCallMethod("getElementsByName", ole.StringToBStrVariant("btnK")))
	btnG := ole.VariantToComObject[ole.IDispatch](elems.MustCallMethod("item", ole.Int32ToVariant(0)))
	btnG.MustCallMethod("click")
}
