//go:build windows
// +build windows

package main

import (
	"github.com/go-ole/go-ole"
	"time"
)

func main() {
	ole.Initialize(ole.Multithreaded)
	defer ole.Uninitialize()
	clsid, _ := ole.LookupClassId("Agent.Control.1")
	unknown, _ := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	agent, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	agent.PutProperty("Connected", true)
	characters := agent.MustGetProperty("Characters").ToIDispatch()
	characters.CallMethod("Load", "Merlin", "c:\\windows\\msagent\\chars\\Merlin.acs")
	character := characters.MustCallMethod("Character", "Merlin").ToIDispatch()
	character.CallMethod("Show")
	character.CallMethod("Speak", "こんにちわ世界")

	time.Sleep(4000000000)
}
