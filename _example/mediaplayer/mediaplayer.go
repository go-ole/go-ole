//go:build windows
// +build windows

package main

import (
	"fmt"
	"github.com/go-ole/go-ole/legacy"
	"log"

	"github.com/go-ole/go-ole/oleutil"
)

func main() {
	legacy.CoInitialize(0)
	unknown, err := oleutil.CreateObject("WMPlayer.OCX")
	if err != nil {
		log.Fatal(err)
	}
	wmp := unknown.MustQueryInterface(legacy.IID_IDispatch)
	collection := oleutil.MustGetProperty(wmp, "MediaCollection").ToIDispatch()
	list := oleutil.MustCallMethod(collection, "getAll").ToIDispatch()
	count := int(oleutil.MustGetProperty(list, "count").Val)
	for i := 0; i < count; i++ {
		item := oleutil.MustGetProperty(list, "item", i).ToIDispatch()
		name := oleutil.MustGetProperty(item, "name").ToString()
		sourceURL := oleutil.MustGetProperty(item, "sourceURL").ToString()
		fmt.Println(name, sourceURL)
	}
}
