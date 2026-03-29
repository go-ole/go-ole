//go:build windows

package ole

import "unsafe"

var (
	procGetMessageW        = moduser32.NewProc("GetMessageW")
	procDispatchMessageW   = moduser32.NewProc("DispatchMessageW")
	procPostThreadMessageW = moduser32.NewProc("PostThreadMessageW")
	procPostMessageW       = moduser32.NewProc("PostMessageW")
)

// Msg is message between processes.
type Msg struct {
	Hwnd    uintptr
	Message uint32
	Wparam  uintptr
	Lparam  uintptr
	Time    uint32
	Pt      Point
}

// GetMessage in message queue from runtime.
//
// This function appears to block. PeekMessage does not block.
func GetMessage(msg *Msg, hwnd uintptr, MsgFilterMin uint32, MsgFilterMax uint32) (ret int32, err error) {
	r0, _, err := procGetMessageW.Call(uintptr(unsafe.Pointer(msg)), hwnd, uintptr(MsgFilterMin), uintptr(MsgFilterMax))
	ret = int32(r0)
	return
}

// DispatchMessage to window procedure.
func DispatchMessage(msg *Msg) (ret int32) {
	r0, _, _ := procDispatchMessageW.Call(uintptr(unsafe.Pointer(msg)))
	ret = int32(r0)
	return
}

// PostThreadMessage posts a message to the message queue of the specified thread.
func PostThreadMessage(threadID uint32, msg uint32, wParam uintptr, lParam uintptr) (ret int32, err error) {
	r0, _, err := procPostThreadMessageW.Call(uintptr(threadID), uintptr(msg), wParam, lParam)
	ret = int32(r0)
	return
}

// PostMessage posts a message to the message queue of the specified window.
func PostMessage(hwnd uintptr, msg uint32, wParam uintptr, lParam uintptr) (ret int32, err error) {
	r0, _, err := procPostMessageW.Call(hwnd, uintptr(msg), wParam, lParam)
	ret = int32(r0)
	return
}
