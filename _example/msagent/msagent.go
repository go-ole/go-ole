//go:build windows
// +build windows

package main

import (
	"github.com/go-ole/go-ole/legacy"
	"time"

	"github.com/go-ole/go-ole/oleutil"
)

func main() {
	legacy.CoInitialize(0)
	unknown, _ := oleutil.CreateObject("Agent.Control.1")
	agent, _ := unknown.QueryInterface(legacy.IID_IDispatch)
	oleutil.PutProperty(agent, "Connected", true)
	characters := oleutil.MustGetProperty(agent, "Characters").ToIDispatch()
	oleutil.CallMethod(characters, "Load", "Merlin", "c:\\windows\\msagent\\chars\\Merlin.acs")
	character := oleutil.MustCallMethod(characters, "Character", "Merlin").ToIDispatch()
	oleutil.CallMethod(character, "Show")
	oleutil.CallMethod(character, "Speak", "こんにちわ世界")

	time.Sleep(4000000000)
}
