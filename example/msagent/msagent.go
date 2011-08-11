package main

import "github.com/mattn/go-ole"
import "github.com/mattn/go-ole/oleutil"
import "syscall"

func main() {
	ole.CoInitialize(0)
	unknown, _ := oleutil.CreateObject("Agent.Control.1")
	agent, _ := unknown.QueryInterface(ole.IID_IDispatch)
	oleutil.PutProperty(agent, "Connected", true)
	result, _ := oleutil.GetProperty(agent, "Characters")
	characters := result.ToIDispatch()
	oleutil.CallMethod(characters, "Load", "Merlin", "c:\\windows\\msagent\\chars\\Merlin.acs")
	result, _ = oleutil.CallMethod(characters, "Character", "Merlin")
	character := result.ToIDispatch()
	oleutil.CallMethod(character, "Show")
	oleutil.CallMethod(character, "Speak", "こんにちわ世界")

	syscall.Sleep(4000000000)
}
