//go:build windows
// +build windows

package main

import (
	"fmt"
	"time"

	"github.com/go-ole/go-ole"
)

func main() {
	ole.InitializeMultithreaded()
	defer ole.Uninitialize()
	ole.RegisterVariantConverters()

	clsid, _ := ole.ClassIdFromString("Microsoft.XMLHTTP")
	unknown, _ := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	defer unknown.Release()

	xmlhttp, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	defer xmlhttp.Release()

	_, err := xmlhttp.CallMethod("open", ole.StringToBStrVariant("GET"), ole.StringToBStrVariant("http://rss.slashdot.org/Slashdot/slashdot"), ole.BoolToVariant(false))
	if err != nil {
		panic(err.Error())
	}
	_, err = xmlhttp.CallMethod("send")
	if err != nil {
		panic(err.Error())
	}
	state := -1
	for state != 4 {
		state = int(ole.UnwrapVariant[int32](xmlhttp.MustGetProperty("readyState")))
		time.Sleep(10000000)
	}
	responseXml := ole.VariantToComObject[ole.IDispatch](xmlhttp.MustGetProperty("responseXml"))
	items := ole.VariantToComObject[ole.IDispatch](responseXml.MustCallMethod("selectNodes", ole.StringToBStrVariant("/rdf:RDF/item")))
	length := int(ole.UnwrapVariant[int32](items.MustGetProperty("length")))

	println(length)
	for n := 0; n < length; n++ {
		item := ole.VariantToComObject[ole.IDispatch](items.MustGetProperty("item", ole.Int32ToVariant(n)))

		title := ole.VariantToComObject[ole.IDispatch](item.MustCallMethod("selectSingleNode", ole.StringToBStrVariant("title")))
		fmt.Println(ole.UnwrapVariant[string](title.MustGetProperty("text")))

		link := ole.VariantToComObject[ole.IDispatch](item.MustCallMethod("selectSingleNode", ole.StringToBStrVariant("link")))
		fmt.Println("  " + ole.UnwrapVariant[string](link.MustGetProperty("text")))

		title.Release()
		link.Release()
		item.Release()
	}
	items.Release()
	responseXml.Release()
}
