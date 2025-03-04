//go:build windows

package ole

import (
	"errors"
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	ComInterfaceNotImplementedError = errors.New("IUnknown: COM Interface is not castable to attempted COM interface")
	ComInterfaceIsNullPointer       = errors.New("IUnknown: COM Interface is null pointer (this is a bug, please report this error)")
	ComInterfaceIsNilPointer        = errors.New("InvalidArgument: COM Interface is nil")
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

func (v *IUnknown) QueryInterfaceAddress() uintptr {
	return v.QueryInterface
}

func (v *IUnknown) AddRefAddress() uintptr {
	return v.addRef
}

func (v *IUnknown) ReleaseAddress() uintptr {
	return v.release
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
		uintptr(unsafe.Pointer(unknown)),
		uintptr(unsafe.Pointer(interfaceID)),
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
	ret, _, _ := syscall.Syscall(unknown.AddRefAddress(), 1, uintptr(unsafe.Pointer(unknown)), 0, 0)
	return uint32(ret)
}

func ReleaseOnIUnknown(unknown IsIUnknown) uint32 {
	if unknown == nil {
		return 0
	}
	ret, _, _ := syscall.Syscall(unknown.ReleaseAddress(), 1, uintptr(unsafe.Pointer(unknown)), 0, 0)
	return uint32(ret)
}
