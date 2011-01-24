package main

import "ole"
import "ole/oleutil"
import "syscall"

func main() {
	ole.CoInitialize(0)
	xmlhttp, _ := oleutil.CreateDispatch("Microsoft.XMLHTTP")
	oleutil.CallMethod(xmlhttp, "open", "GET", "http://rss.slashdot.org/Slashdot/slashdot", false)
	oleutil.CallMethod(xmlhttp, "send", nil)
	state := -1
	for state != 4 {
		result, _ := oleutil.GetProperty(xmlhttp, "readyState")
		state = int(result.Val)
		syscall.Sleep(10000000)
	}
	result, _ := oleutil.GetProperty(xmlhttp, "responseXml")
	responseXml := result.ToIDispatch()
	result, _ = oleutil.CallMethod(responseXml, "selectNodes", "rdf:RDF/item")
	items := result.ToIDispatch()
	result, _ = oleutil.GetProperty(items, "length")

	for n := 0; n < int(result.Val); n++ {
		result, _ := oleutil.GetProperty(items, "item", n)
		item := result.ToIDispatch()
		
		result, _ = oleutil.CallMethod(item, "selectSingleNode", "title")
		title := result.ToIDispatch()
		result, _ = oleutil.GetProperty(title, "text")
		println(result.ToString())

		result, _ = oleutil.CallMethod(item, "selectSingleNode", "link")
		link := result.ToIDispatch()
		result, _ = oleutil.GetProperty(link, "text")
		println("  " + result.ToString())

		title.Release()
		link.Release()
		item.Release()
	}
	items.Release()
	xmlhttp.Release()
}
