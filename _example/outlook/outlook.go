//go:build windows
// +build windows

package main

import (
	"fmt"
	"github.com/go-ole/go-ole"
)

func main() {
	ole.InitializeMultithreaded()
	defer ole.Uninitialize()

	clsid, _ := ole.ClassIdFromString("Outlook.Application")

	unknown, _ := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	defer unknown.Release()

	outlook, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	defer unknown.Release()

	ns := ole.VariantToComObject[ole.IDispatch](outlook.MustCallMethod("GetNamespace", ole.StringToBStrVariant("MAPI")))
	defer ns.Release()

	folder := ole.VariantToComObject[ole.IDispatch](ns.MustCallMethod("GetDefaultFolder", ole.Int32ToVariant(10)))
	defer folder.Release()

	contacts := ole.VariantToComObject[ole.IDispatch](folder.MustCallMethod("Items"))
	defer contacts.Release()

	count := ole.UnwrapVariant[int32](contacts.MustGetProperty("Count"))
	for i := 1; i <= int(count); i++ {
		item, err := contacts.GetProperty("Item", ole.Int32ToVariant(i))
		if err == nil && item.VT == ole.VT_DISPATCH {
			value := ole.VariantToComObject[ole.IDispatch](item)
			if value == nil {
				continue
			}
			
			fullName := ole.UnwrapVariant[string](value.GetProperty("FullName"))
			if fullName != nil {
				fmt.Println(fullName)
			}
		}
	}

	outlook.MustCallMethod("Quit")
}
