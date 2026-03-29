//go:build windows

package ole

import (
	"errors"
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestIConnectionPointGetConnectionInterface(t *testing.T) {
	want := IID_IDispatch
	virtualTable := &IConnectionPointVirtualTable{
		getConnectionInterface: syscall.NewCallback(func(this uintptr, iid uintptr) uintptr {
			*(*windows.GUID)(unsafe.Pointer(iid)) = want
			return uintptr(windows.S_OK)
		}),
	}
	connectionPoint := &IConnectionPoint{VirtualTable: virtualTable}

	got, err := connectionPoint.GetConnectionInterface()
	if err != nil {
		t.Fatalf("GetConnectionInterface failed: %v", err)
	}
	if got != want {
		t.Fatalf("GetConnectionInterface = %v, want %v", got, want)
	}
}

func TestIConnectionPointGetConnectionPointContainer(t *testing.T) {
	want := &IConnectionPointContainer{}
	virtualTable := &IConnectionPointVirtualTable{
		GetConnectionPointContainer: syscall.NewCallback(func(this uintptr, container uintptr) uintptr {
			*(**IConnectionPointContainer)(unsafe.Pointer(container)) = want
			return uintptr(windows.S_OK)
		}),
	}
	connectionPoint := &IConnectionPoint{VirtualTable: virtualTable}

	got, err := connectionPoint.GetConnectionPointContainer()
	if err != nil {
		t.Fatalf("GetConnectionPointContainer failed: %v", err)
	}
	if got != want {
		t.Fatalf("GetConnectionPointContainer = %p, want %p", got, want)
	}
}

func TestIConnectionPointAdviseAndUnadvise(t *testing.T) {
	const wantCookie = uint32(77)
	var gotAdviseUnknown uintptr
	var gotUnadviseCookie uint32

	sink := &IUnknown{}
	sinkInterface := IsIUnknown(sink)
	virtualTable := &IConnectionPointVirtualTable{
		advise: syscall.NewCallback(func(this uintptr, unknown uintptr, cookie uintptr) uintptr {
			gotAdviseUnknown = unknown
			*(*uint32)(unsafe.Pointer(cookie)) = wantCookie
			return uintptr(windows.S_OK)
		}),
		unadvise: syscall.NewCallback(func(this uintptr, cookie uintptr) uintptr {
			gotUnadviseCookie = uint32(cookie)
			return uintptr(windows.S_OK)
		}),
	}
	connectionPoint := &IConnectionPoint{VirtualTable: virtualTable}

	cookie, err := connectionPoint.Advise(&sinkInterface)
	if err != nil {
		t.Fatalf("Advise failed: %v", err)
	}
	if cookie != wantCookie {
		t.Fatalf("Advise cookie = %d, want %d", cookie, wantCookie)
	}
	if gotAdviseUnknown != uintptr(unsafe.Pointer(&sinkInterface)) {
		t.Fatalf("Advise unknown pointer = %#x, want %#x", gotAdviseUnknown, uintptr(unsafe.Pointer(&sinkInterface)))
	}

	if err := connectionPoint.Unadvise(cookie); err != nil {
		t.Fatalf("Unadvise failed: %v", err)
	}
	if gotUnadviseCookie != wantCookie {
		t.Fatalf("Unadvise cookie = %d, want %d", gotUnadviseCookie, wantCookie)
	}
}

func TestIConnectionPointEnumConnections(t *testing.T) {
	want := &IEnumConnections{}
	virtualTable := &IConnectionPointVirtualTable{
		enumConnections: syscall.NewCallback(func(this uintptr, connections uintptr) uintptr {
			*(**IEnumConnections)(unsafe.Pointer(connections)) = want
			return uintptr(windows.S_OK)
		}),
	}
	connectionPoint := &IConnectionPoint{VirtualTable: virtualTable}

	got, err := connectionPoint.EnumConnections()
	if err != nil {
		t.Fatalf("EnumConnections failed: %v", err)
	}
	if got != want {
		t.Fatalf("EnumConnections = %p, want %p", got, want)
	}
}

func TestIConnectionPointReportsHRESULTErrors(t *testing.T) {
	hr := uintptr(windows.E_POINTER)
	virtualTable := &IConnectionPointVirtualTable{
		getConnectionInterface: syscall.NewCallback(func(this uintptr, iid uintptr) uintptr {
			return hr
		}),
		GetConnectionPointContainer: syscall.NewCallback(func(this uintptr, container uintptr) uintptr {
			return hr
		}),
		advise: syscall.NewCallback(func(this uintptr, unknown uintptr, cookie uintptr) uintptr {
			return hr
		}),
		unadvise: syscall.NewCallback(func(this uintptr, cookie uintptr) uintptr {
			return hr
		}),
		enumConnections: syscall.NewCallback(func(this uintptr, connections uintptr) uintptr {
			return hr
		}),
	}
	connectionPoint := &IConnectionPoint{VirtualTable: virtualTable}
	sink := &IUnknown{}
	sinkInterface := IsIUnknown(sink)

	if _, err := connectionPoint.GetConnectionInterface(); !errors.Is(err, windows.Errno(hr)) {
		t.Fatalf("GetConnectionInterface error = %v, want %v", err, windows.Errno(hr))
	}
	if _, err := connectionPoint.GetConnectionPointContainer(); !errors.Is(err, windows.Errno(hr)) {
		t.Fatalf("GetConnectionPointContainer error = %v, want %v", err, windows.Errno(hr))
	}
	if _, err := connectionPoint.Advise(&sinkInterface); !errors.Is(err, windows.Errno(hr)) {
		t.Fatalf("Advise error = %v, want %v", err, windows.Errno(hr))
	}
	if err := connectionPoint.Unadvise(1); !errors.Is(err, windows.Errno(hr)) {
		t.Fatalf("Unadvise error = %v, want %v", err, windows.Errno(hr))
	}
	if _, err := connectionPoint.EnumConnections(); !errors.Is(err, windows.Errno(hr)) {
		t.Fatalf("EnumConnections error = %v, want %v", err, windows.Errno(hr))
	}
}

func TestQueryIConnectionPointFromIUnknownNil(t *testing.T) {
	got, err := QueryIConnectionPointFromIUnknown(nil)
	if got != nil {
		t.Fatalf("QueryIConnectionPointFromIUnknown(nil) = %p, want nil", got)
	}
	if !errors.Is(err, ComInterfaceIsNilPointer) {
		t.Fatalf("QueryIConnectionPointFromIUnknown(nil) error = %v, want %v", err, ComInterfaceIsNilPointer)
	}
}
