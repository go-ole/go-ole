//go:build windows
// +build windows

package main

import (
	"github.com/go-ole/go-ole"
	"time"
)

func main() {
	ole.InitializeMultithreaded()
	defer ole.Uninitialize()
	ole.RegisterVariantConverters()

	clsid, _ := ole.ClassIdFromString("Agent.Control.1")
	unknown, _ := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	defer unknown.Release()

	agent, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	defer agent.Release()

	agent.PutProperty("Connected", ole.BoolToVariant(true))
	characters := ole.VariantToComObject[ole.IDispatch](agent.MustGetProperty("Characters"))
	defer characters.Release()

	characters.CallMethod("Load", ole.StringToBStrVariant("Merlin"), ole.StringToBStrVariant("c:\\windows\\msagent\\chars\\Merlin.acs"))
	character := ole.VariantToComObject[ole.IDispatch](characters.MustCallMethod("Character", ole.StringToBStrVariant("Merlin")))
	defer character.Release()

	character.CallMethod("Show")
	character.CallMethod("Speak", ole.StringToBStrVariant("こんにちわ世界"))

	time.Sleep(4000000000)
}
