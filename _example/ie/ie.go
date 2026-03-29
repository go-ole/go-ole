//go:build windows
// +build windows

package main

import (
	"time"

	"github.com/go-ole/go-ole"
)

func main() {
	ole.Initialize(ole.Multithreaded)
	defer ole.Uninitialize()
	clsid, _ := ole.LookupClassId("InternetExplorer.Application")
	unknown, _ := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	ie, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	ie.PutProperty("Visible", true)
	ie.CallMethod("Navigate", "http://www.google.com")
	for {
		if ie.MustGetProperty("Busy").Val == 0 {
			break
		}
	}

	time.Sleep(1e9)

	document := ie.MustGetProperty("document").ToIDispatch()

	// set 'golang' to text box.
	elems := document.MustCallMethod("getElementsByName", "q").ToIDispatch()
	q := elems.MustCallMethod("item", 0).ToIDispatch()
	q.MustPutProperty("value", "golang")

	// click btnK.
	elems = document.MustCallMethod("getElementsByName", "btnK").ToIDispatch()
	btnG := elems.MustCallMethod("item", 0).ToIDispatch()
	btnG.MustCallMethod("click")
}
