//go:build windows

package ole

import "unsafe"

var (
	procGetMessageW      = moduser32.NewProc("GetMessageW")
	procDispatchMessageW = moduser32.NewProc("DispatchMessageW")
)

// Msg is message between processes.
type Msg struct {
	Hwnd    uint32
	Message uint32
	Wparam  int32
	Lparam  int32
	Time    uint32
	Pt      Point
}

// GetMessage in message queue from runtime.
//
// This function appears to block. PeekMessage does not block.
func GetMessage(msg *Msg, hwnd uint32, MsgFilterMin uint32, MsgFilterMax uint32) (ret int32, err error) {
	r0, _, err := procGetMessageW.Call(uintptr(unsafe.Pointer(msg)), uintptr(hwnd), uintptr(MsgFilterMin), uintptr(MsgFilterMax))
	ret = int32(r0)
	return
}

// DispatchMessage to window procedure.
func DispatchMessage(msg *Msg) (ret int32) {
	r0, _, _ := procDispatchMessageW.Call(uintptr(unsafe.Pointer(msg)))
	ret = int32(r0)
	return
}
