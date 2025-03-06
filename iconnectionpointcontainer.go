//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

type IConnectionPointContainer struct {
	VirtualTable *IConnectionPointContainerVirtualTable
}

type IConnectionPointContainerVirtualTable struct {
	// IUnknown
	QueryInterface uintptr
	addRef         uintptr
	release        uintptr
	// IConnectionPointContainer
	enumConnectionPoints uintptr
	findConnectionPoint  uintptr
}

func (obj *IConnectionPointContainer) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *IConnectionPointContainer) AddRefAddress() uintptr {
	return obj.VirtualTable.addRef
}

func (obj *IConnectionPointContainer) ReleaseAddress() uintptr {
	return obj.VirtualTable.release
}

func (obj *IConnectionPointContainer) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IConnectionPointContainer) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IConnectionPointContainer) EnumConnectionPoints(points interface{}) error {
	return MethodNotImplementedError
}

func (obj *IConnectionPointContainer) FindConnectionPoint(iid windows.GUID) (point *IConnectionPoint, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.findConnectionPoint,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&iid)),
		uintptr(unsafe.Pointer(&point)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

func QueryIConnectionPointContainerFromIUnknown(unknown IsIUnknown) (obj *IConnectionPointContainer, err error) {
	if unknown == nil {
		return nil, ComInterfaceIsNilPointer
	}

	obj, err = QueryInterfaceOnIUnknown[IConnectionPointContainer](unknown, IID_IConnectionPointContainer)
	if err != nil {
		return nil, err
	}
	return
}
