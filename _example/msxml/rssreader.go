//go:build windows
// +build windows

package main

import (
	"fmt"
	"time"

	"github.com/go-ole/go-ole"
)

func main() {
	ole.Initialize(ole.Multithreaded)
	defer ole.Uninitialize()
	clsid, _ := ole.LookupClassId("Microsoft.XMLHTTP")
	unknown, _ := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	xmlhttp, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	_, err := xmlhttp.CallMethod("open", "GET", "http://rss.slashdot.org/Slashdot/slashdot", false)
	if err != nil {
		panic(err.Error())
	}
	_, err = xmlhttp.CallMethod("send", nil)
	if err != nil {
		panic(err.Error())
	}
	state := -1
	for state != 4 {
		state = int(xmlhttp.MustGetProperty("readyState").Val)
		time.Sleep(10000000)
	}
	responseXml := xmlhttp.MustGetProperty("responseXml").ToIDispatch()
	items := responseXml.MustCallMethod("selectNodes", "/rdf:RDF/item").ToIDispatch()
	length := int(items.MustGetProperty("length").Val)

	println(length)
	for n := 0; n < length; n++ {
		item := items.MustGetProperty("item", n).ToIDispatch()

		title := item.MustCallMethod("selectSingleNode", "title").ToIDispatch()
		fmt.Println(title.MustGetProperty("text").ToString())

		link := item.MustCallMethod("selectSingleNode", "link").ToIDispatch()
		fmt.Println("  " + link.MustGetProperty("text").ToString())

		title.Release()
		link.Release()
		item.Release()
	}
	items.Release()
	xmlhttp.Release()
}
