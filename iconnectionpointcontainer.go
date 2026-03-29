//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

// IConnectionPointContainerAddresses describes the IConnectionPointContainer vtable entries.
//
// Example:
//
//	container, err := ole.QueryIConnectionPointContainerFromIUnknown(dispatch)
//	if err != nil {
//		return err
//	}
//	defer container.Release()
type IConnectionPointContainerAddresses interface {
	IsIUnknown
	EnumConnectionPointsAddress() uintptr
	FindConnectionPointAddress() uintptr
}

// IConnectionPointContainer represents a COM object that exposes connection points.
type IConnectionPointContainer struct {
	VirtualTable *IConnectionPointContainerVirtualTable
}

// IConnectionPointContainerVirtualTable contains the native function pointers for IConnectionPointContainer.
type IConnectionPointContainerVirtualTable struct {
	// IUnknown
	QueryInterface uintptr
	addRef         uintptr
	release        uintptr
	// IConnectionPointContainer
	enumConnectionPoints uintptr
	findConnectionPoint  uintptr
}

// QueryInterfaceAddress returns the QueryInterface entry point for obj.
func (obj *IConnectionPointContainer) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

// AddRefAddress returns the AddRef entry point for obj.
func (obj *IConnectionPointContainer) AddRefAddress() uintptr {
	return obj.VirtualTable.addRef
}

// ReleaseAddress returns the Release entry point for obj.
func (obj *IConnectionPointContainer) ReleaseAddress() uintptr {
	return obj.VirtualTable.release
}

// EnumConnectionPointsAddress returns the EnumConnectionPoints entry point for obj.
func (obj *IConnectionPointContainer) EnumConnectionPointsAddress() uintptr {
	return obj.VirtualTable.enumConnectionPoints
}

// FindConnectionPointAddress returns the FindConnectionPoint entry point for obj.
func (obj *IConnectionPointContainer) FindConnectionPointAddress() uintptr {
	return obj.VirtualTable.findConnectionPoint
}

// AddRef increments the COM reference count for obj.
func (obj *IConnectionPointContainer) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

// Release decrements the COM reference count for obj.
func (obj *IConnectionPointContainer) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

// EnumConnectionPoints returns an enumerator for the connection points exposed by obj.
func (obj *IConnectionPointContainer) EnumConnectionPoints() (points *IEnumConnections, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.enumConnectionPoints,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&points)),
		0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// FindConnectionPoint looks up the connection point for the requested interface IID.
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

// QueryIConnectionPointContainerFromIUnknown casts unknown to IConnectionPointContainer.
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
