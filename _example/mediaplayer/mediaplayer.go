//go:build windows
// +build windows

package main

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"log"
)

func main() {
	ole.InitializeMultithreaded()
	defer ole.Uninitialize()
	ole.RegisterVariantConverters()

	clsid, err := ole.ClassIdFromString("WMPlayer.OCX")
	if err != nil {
		log.Fatal(err)
	}
	unknown, err := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	if err != nil {
		log.Fatal(err)
	}
	defer unknown.Release()

	wmp, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	defer wmp.Release()

	collection := ole.VariantToComObject[ole.IDispatch](wmp.MustGetProperty("MediaCollection"))
	defer collection.Release()

	list := ole.VariantToComObject[ole.IDispatch](collection.MustCallMethod("getAll"))
	defer list.Release()

	count := int(ole.UnwrapVariant[int32](list.MustGetProperty("count")))
	for i := 0; i < count; i++ {
		item := ole.VariantToComObject[ole.IDispatch](list.MustGetProperty("item", ole.Int32ToVariant(i)))
		name := ole.UnwrapVariant[string](item.MustGetProperty("name"))
		sourceURL := ole.UnwrapVariant[string](item.MustGetProperty("sourceURL"))
		fmt.Println(name, sourceURL)
		item.Release()
	}
}
