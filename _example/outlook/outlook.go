//go:build windows
// +build windows

package main

import (
	"fmt"
	"github.com/go-ole/go-ole"
)

func main() {
	ole.Initialize(ole.Multithreaded)
	defer ole.Uninitialize()
	clsid, _ := ole.LookupClassId("Outlook.Application")
	unknown, _ := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	outlook, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	ns := outlook.MustCallMethod("GetNamespace", "MAPI").ToIDispatch()
	folder := ns.MustCallMethod("GetDefaultFolder", 10).ToIDispatch()
	contacts := folder.MustCallMethod("Items").ToIDispatch()
	count := contacts.MustGetProperty("Count").Value().(int32)
	for i := 1; i <= int(count); i++ {
		item, err := contacts.GetProperty("Item", i)
		if err == nil && item.VT == ole.VT_DISPATCH {
			if value, err := item.ToIDispatch().GetProperty("FullName"); err == nil {
				fmt.Println(value.Value())
			}
		}
	}
	outlook.MustCallMethod("Quit")
}
