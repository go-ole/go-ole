//go:build windows

package ole

import (
	"errors"
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestIConnectionPointContainerEnumConnectionPoints(t *testing.T) {
	want := &IEnumConnections{}
	virtualTable := &IConnectionPointContainerVirtualTable{
		enumConnectionPoints: syscall.NewCallback(func(this uintptr, points uintptr) uintptr {
			*(**IEnumConnections)(unsafe.Pointer(points)) = want
			return uintptr(windows.S_OK)
		}),
	}
	container := &IConnectionPointContainer{VirtualTable: virtualTable}

	got, err := container.EnumConnectionPoints()
	if err != nil {
		t.Fatalf("EnumConnectionPoints failed: %v", err)
	}
	if got != want {
		t.Fatalf("EnumConnectionPoints = %p, want %p", got, want)
	}
}

func TestIConnectionPointContainerFindConnectionPoint(t *testing.T) {
	wantIID := IID_IConnectionPoint
	want := &IConnectionPoint{}
	var gotIID windows.GUID

	virtualTable := &IConnectionPointContainerVirtualTable{
		findConnectionPoint: syscall.NewCallback(func(this uintptr, iid uintptr, point uintptr) uintptr {
			gotIID = *(*windows.GUID)(unsafe.Pointer(iid))
			*(**IConnectionPoint)(unsafe.Pointer(point)) = want
			return uintptr(windows.S_OK)
		}),
	}
	container := &IConnectionPointContainer{VirtualTable: virtualTable}

	got, err := container.FindConnectionPoint(wantIID)
	if err != nil {
		t.Fatalf("FindConnectionPoint failed: %v", err)
	}
	if got != want {
		t.Fatalf("FindConnectionPoint = %p, want %p", got, want)
	}
	if gotIID != wantIID {
		t.Fatalf("FindConnectionPoint iid = %v, want %v", gotIID, wantIID)
	}
}

func TestIConnectionPointContainerReportsHRESULTErrors(t *testing.T) {
	hr := uintptr(windows.E_POINTER)
	virtualTable := &IConnectionPointContainerVirtualTable{
		enumConnectionPoints: syscall.NewCallback(func(this uintptr, points uintptr) uintptr {
			return hr
		}),
		findConnectionPoint: syscall.NewCallback(func(this uintptr, iid uintptr, point uintptr) uintptr {
			return hr
		}),
	}
	container := &IConnectionPointContainer{VirtualTable: virtualTable}

	if _, err := container.EnumConnectionPoints(); !errors.Is(err, windows.Errno(hr)) {
		t.Fatalf("EnumConnectionPoints error = %v, want %v", err, windows.Errno(hr))
	}
	if _, err := container.FindConnectionPoint(IID_IConnectionPoint); !errors.Is(err, windows.Errno(hr)) {
		t.Fatalf("FindConnectionPoint error = %v, want %v", err, windows.Errno(hr))
	}
}

func TestQueryIConnectionPointContainerFromIUnknownNil(t *testing.T) {
	got, err := QueryIConnectionPointContainerFromIUnknown(nil)
	if got != nil {
		t.Fatalf("QueryIConnectionPointContainerFromIUnknown(nil) = %p, want nil", got)
	}
	if !errors.Is(err, ComInterfaceIsNilPointer) {
		t.Fatalf("QueryIConnectionPointContainerFromIUnknown(nil) error = %v, want %v", err, ComInterfaceIsNilPointer)
	}
}
