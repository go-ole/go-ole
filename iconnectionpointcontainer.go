//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

type IConnectionPointContainer struct {
	// IUnknown
	QueryInterface uintptr
	addRef         uintptr
	release        uintptr
	// IConnectionPointContainer
	enumConnectionPoints uintptr
	findConnectionPoint  uintptr
}

func (obj *IConnectionPointContainer) EnumConnectionPoints(points interface{}) error {
	return NewError(E_NOTIMPL)
}

func (obj *IConnectionPointContainer) FindConnectionPoint(iid windows.GUID) (point *IConnectionPoint, err error) {
	hr, _, _ := syscall.Syscall(
		obj.findConnectionPoint,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&point)))
	if hr != windows.S_OK {
		err = hr
	}
	return
}

func QueryIConnectionPointContainerFromIUnknown(unknown *IsIUnknown) (obj *IConnectionPointContainer, err error) {
	if unknown == nil {
		return nil, ComInterfaceIsNilPointer
	}

	enum, err = QueryInterfaceOnIUnknown[IConnectionPointContainer](unknown, IID_IConnectionPointContainer)
	if err != nil {
		return nil, err
	}
	return
}
