//go:build windows
// +build windows

package main

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"log"
)

func main() {
	ole.Initialize(ole.Multithreaded)
	defer ole.Uninitialize()
	clsid, err := ole.LookupClassId("WMPlayer.OCX")
	if err != nil {
		log.Fatal(err)
	}
	unknown, err := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	if err != nil {
		log.Fatal(err)
	}
	wmp, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	collection := wmp.MustGetProperty("MediaCollection").ToIDispatch()
	list := collection.MustCallMethod("getAll").ToIDispatch()
	count := int(list.MustGetProperty("count").Val)
	for i := 0; i < count; i++ {
		item := list.MustGetProperty("item", i).ToIDispatch()
		name := item.MustGetProperty("name").ToString()
		sourceURL := item.MustGetProperty("sourceURL").ToString()
		fmt.Println(name, sourceURL)
	}
}
