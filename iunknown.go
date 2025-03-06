//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

type IsIUnknown interface {
	QueryInterfaceAddress() uintptr
	AddRefAddress() uintptr
	ReleaseAddress() uintptr
}

type IUnknown struct {
	QueryInterface uintptr
	addRef         uintptr
	release        uintptr
}

func (obj *IUnknown) QueryInterfaceAddress() uintptr {
	return obj.QueryInterface
}

func (obj *IUnknown) AddRefAddress() uintptr {
	return obj.addRef
}

func (obj *IUnknown) ReleaseAddress() uintptr {
	return obj.release
}

func (obj *IUnknown) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

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

	switch hr {
	case windows.S_OK:
		return out, nil
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

func AddRefOnIUnknown(unknown IsIUnknown) uint32 {
	if unknown == nil {
		return 0
	}
	ret, _, _ := syscall.Syscall(unknown.AddRefAddress(), 1, uintptr(unsafe.Pointer(&unknown)), 0, 0)
	return uint32(ret)
}

func ReleaseOnIUnknown(unknown IsIUnknown) uint32 {
	if unknown == nil {
		return 0
	}
	ret, _, _ := syscall.Syscall(unknown.ReleaseAddress(), 1, uintptr(unsafe.Pointer(&unknown)), 0, 0)
	return uint32(ret)
}
