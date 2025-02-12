//go:build windows
// +build windows

package main

import (
	"github.com/go-ole/go-ole/legacy"
	"time"

	"github.com/go-ole/go-ole/oleutil"
)

func main() {
	legacy.CoInitialize(0)
	unknown, _ := oleutil.CreateObject("InternetExplorer.Application")
	ie, _ := unknown.QueryInterface(legacy.IID_IDispatch)
	oleutil.PutProperty(ie, "Visible", true)
	oleutil.CallMethod(ie, "Navigate", "http://www.google.com")
	for {
		if oleutil.MustGetProperty(ie, "Busy").Val == 0 {
			break
		}
	}

	time.Sleep(1e9)

	document := oleutil.MustGetProperty(ie, "document").ToIDispatch()

	// set 'golang' to text box.
	elems := oleutil.MustCallMethod(document, "getElementsByName", "q").ToIDispatch()
	q := oleutil.MustCallMethod(elems, "item", 0).ToIDispatch()
	oleutil.MustPutProperty(q, "value", "golang")

	// click btnK.
	elems = oleutil.MustCallMethod(document, "getElementsByName", "btnK").ToIDispatch()
	btnG := oleutil.MustCallMethod(elems, "item", 0).ToIDispatch()
	oleutil.MustCallMethod(btnG, "click")
}
