//go:build windows

package ole

import (
	"runtime"
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestMessageLoop(t *testing.T) {
	// We need to lock the OS thread because PostThreadMessage targets a specific thread
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	threadID := windows.GetCurrentThreadId()

	const WM_USER = 0x0400
	const MY_MSG = WM_USER + 1

	// Post a message to ourselves
	_, err := PostThreadMessage(threadID, MY_MSG, 123, 456)
	if err != nil {
		t.Fatalf("PostThreadMessage failed: %v", err)
	}

	var msg Msg
	// GetMessage should return non-zero if successful and not WM_QUIT
	ret, err := GetMessage(&msg, 0, 0, 0)
	if ret == 0 {
		t.Fatalf("GetMessage returned 0 (WM_QUIT), but we expected our message")
	}
	if ret == -1 {
		t.Fatalf("GetMessage failed: %v", err)
	}

	if msg.Message != MY_MSG {
		t.Errorf("Expected message %v, got %v", MY_MSG, msg.Message)
	}
	if msg.Wparam != 123 {
		t.Errorf("Expected Wparam 123, got %v", msg.Wparam)
	}
	if msg.Lparam != 456 {
		t.Errorf("Expected Lparam 456, got %v", msg.Lparam)
	}

	// Test DispatchMessage (for thread messages it doesn't do much, but we can verify it doesn't crash)
	DispatchMessage(&msg)
}

func TestPostQuitMessage(t *testing.T) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	threadID := windows.GetCurrentThreadId()

	const WM_QUIT = 0x0012

	_, err := PostThreadMessage(threadID, WM_QUIT, 0, 0)
	if err != nil {
		t.Fatalf("PostThreadMessage failed: %v", err)
	}

	var msg Msg
	ret, _ := GetMessage(&msg, 0, 0, 0)
	if ret != 0 {
		t.Errorf("Expected GetMessage to return 0 for WM_QUIT, got %v", ret)
	}
}

func TestDispatchMessageActual(t *testing.T) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	moduser32 := windows.NewLazySystemDLL("user32.dll")
	procDefWindowProcW = moduser32.NewProc("DefWindowProcW")
	procRegisterClassExW = moduser32.NewProc("RegisterClassExW")
	procCreateWindowExW = moduser32.NewProc("CreateWindowExW")
	procDestroyWindow = moduser32.NewProc("DestroyWindow")

	const (
		WM_USER = 0x0400
		MY_MSG  = WM_USER + 1
	)

	var receivedMessage uint32
	var receivedWparam uintptr
	var receivedLparam uintptr

	wndProc := func(hwnd uintptr, msg uint32, wparam uintptr, lparam uintptr) uintptr {
		if msg == MY_MSG {
			receivedMessage = msg
			receivedWparam = wparam
			receivedLparam = lparam
			return 0
		}
		r0, _, _ := procDefWindowProcW.Call(hwnd, uintptr(msg), wparam, lparam)
		return r0
	}

	className, _ := windows.UTF16PtrFromString("TestWindowClass")

	type WNDCLASSEXW struct {
		Size       uint32
		Style      uint32
		WndProc    uintptr
		ClsExtra   int32
		WndExtra   int32
		Instance   uintptr
		Icon       uintptr
		Cursor     uintptr
		Background uintptr
		MenuName   *uint16
		ClassName  *uint16
		IconSm     uintptr
	}

	wc := WNDCLASSEXW{
		WndProc:   syscall.NewCallback(wndProc),
		ClassName: className,
	}
	wc.Size = uint32(unsafe.Sizeof(wc))

	r0, _, err := procRegisterClassExW.Call(uintptr(unsafe.Pointer(&wc)))
	if r0 == 0 {
		t.Fatalf("RegisterClassExW failed: %v", err)
	}

	hwnd, _, err := procCreateWindowExW.Call(
		0,
		uintptr(unsafe.Pointer(className)),
		0,
		0, 0, 0, 0, 0,
		0, 0, 0, 0)
	if hwnd == 0 {
		t.Fatalf("CreateWindowExW failed: %v", err)
	}
	defer procDestroyWindow.Call(hwnd)

	// Post message to the window
	PostMessage(hwnd, MY_MSG, 777, 888)

	var msg Msg
	for {
		ret, _ := GetMessage(&msg, 0, 0, 0)
		if ret <= 0 {
			break
		}
		DispatchMessage(&msg)
		if receivedMessage == MY_MSG {
			break
		}
	}

	if receivedMessage != MY_MSG {
		t.Errorf("WndProc did not receive MY_MSG")
	}
	if receivedWparam != 777 {
		t.Errorf("Expected Wparam 777, got %v", receivedWparam)
	}
	if receivedLparam != 888 {
		t.Errorf("Expected Lparam 888, got %v", receivedLparam)
	}
}

var (
	procDefWindowProcW   *windows.LazyProc
	procRegisterClassExW *windows.LazyProc
	procCreateWindowExW  *windows.LazyProc
	procDestroyWindow    *windows.LazyProc
)
