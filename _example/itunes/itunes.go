//go:build windows
// +build windows

package main

import (
	"github.com/go-ole/go-ole"
	"log"
	"os"
	"strings"

	"github.com/gonuts/commander"
)

func iTunes() *ole.IDispatch {
	ole.Initialize(ole.Multithreaded)
	clsid, err := ole.LookupClassId("iTunes.Application")
	if err != nil {
		log.Fatal(err)
	}
	unknown, err := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	if err != nil {
		log.Fatal(err)
	}
	itunes, err := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	if err != nil {
		log.Fatal(err)
	}
	return itunes
}

func main() {
	defer ole.Uninitialize()
	command := &commander.Command{
		UsageLine: os.Args[0],
		Short:     "itunes cmd",
	}
	command.Subcommands = []*commander.Command{}
	for _, name := range []string{"Play", "Stop", "Pause", "Quit"} {
		command.Subcommands = append(command.Subcommands, &commander.Command{
			Run: func(cmd *commander.Command, args []string) error {
				itunes := iTunes()
				defer itunes.Release()
				_, err := itunes.CallMethod(name)
				return err
			},
			UsageLine: strings.ToLower(name),
		})
	}
	err := command.Dispatch(os.Args[1:])
	if err != nil {
		log.Fatal(err)
	}
}
