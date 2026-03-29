//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

// IsIUnknown describes COM interface wrappers that expose the three IUnknown entry points.
//
// Example:
//
//	var unknown ole.IsIUnknown = dispatch
//	_ = unknown.QueryInterfaceAddress()
type IsIUnknown interface {
	QueryInterfaceAddress() uintptr
	AddRefAddress() uintptr
	ReleaseAddress() uintptr
}

// IUnknown is the base COM interface shared by every COM object.
type IUnknown struct {
	VirtualTable *IUnknownVirtualTable
}

// IUnknownVirtualTable contains the native function pointers for IUnknown.
type IUnknownVirtualTable struct {
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
}

// QueryInterfaceAddress returns the QueryInterface entry point for obj.
func (obj *IUnknown) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

// AddRefAddress returns the AddRef entry point for obj.
func (obj *IUnknown) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

// ReleaseAddress returns the Release entry point for obj.
func (obj *IUnknown) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

// AddRef increments the COM reference count for obj.
func (obj *IUnknown) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

// Release decrements the COM reference count for obj.
func (obj *IUnknown) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

// QueryInterfaceOnIUnknown converts IUnknown to another COM interface.
//
// T must be a COM interface virtual table, this is an unsafe action.
func QueryInterfaceOnIUnknown[T any](unknown IsIUnknown, interfaceID windows.GUID) (*T, error) {
	if unknown == nil {
		return nil, ComInterfaceIsNilPointer
	}

	var ret *T
	hr, _, _ := syscall.Syscall(
		unknown.QueryInterfaceAddress(),
		3,
		uintptr(unsafe.Pointer(&unknown)),
		uintptr(unsafe.Pointer(&interfaceID)),
		uintptr(unsafe.Pointer(&ret)))

	switch windows.Handle(hr) {
	case windows.S_OK:
		return ret, nil
	case windows.E_NOINTERFACE:
		return nil, ComInterfaceNotImplementedError
	case windows.E_POINTER:
		return nil, ComInterfaceIsNullPointer
	default:
		return ret, windows.Errno(hr)
	}
}

// MustQueryInterfaceOnIUnknown converts IUnknown to another COM interface or panics.
//
// T must be a COM interface virtual table, this is an unsafe action.
func MustQueryInterfaceOnIUnknown[T any](unknown IsIUnknown, interfaceID windows.GUID) *T {
	if unknown == nil {
		panic(ComInterfaceIsNilPointer)
	}
	ret, err := QueryInterfaceOnIUnknown[T](unknown, interfaceID)
	if err != nil {
		panic(err)
	}
	return ret
}

// AddRefOnIUnknown increments the COM reference count for any IsIUnknown value.
func AddRefOnIUnknown(unknown IsIUnknown) uint32 {
	if unknown == nil {
		return 0
	}
	ret, _, _ := syscall.Syscall(unknown.AddRefAddress(), 1, uintptr(unsafe.Pointer(&unknown)), 0, 0)
	return uint32(ret)
}

// ReleaseOnIUnknown decrements the COM reference count for any IsIUnknown value.
func ReleaseOnIUnknown(unknown IsIUnknown) uint32 {
	if unknown == nil {
		return 0
	}
	ret, _, _ := syscall.Syscall(unknown.ReleaseAddress(), 1, uintptr(unsafe.Pointer(&unknown)), 0, 0)
	return uint32(ret)
}
