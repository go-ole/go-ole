//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

// IConnectionPointAddresses describes the IConnectionPoint vtable entries.
//
// Example:
//
//	point, err := container.FindConnectionPoint(ole.IID_IDispatch)
//	if err != nil {
//		return err
//	}
//	defer point.Release()
type IConnectionPointAddresses interface {
	IsIUnknown
	GetConnectionInterfaceAddress() uintptr
	GetConnectionPointContainerAddress() uintptr
	AdviseAddress() uintptr
	UnadviseAddress() uintptr
	EnumConnectionsAddress() uintptr
}

// IConnectionPoint represents a COM connection point used for event sinks.
type IConnectionPoint struct {
	VirtualTable *IConnectionPointVirtualTable
}

// IConnectionPointVirtualTable contains the native function pointers for IConnectionPoint.
type IConnectionPointVirtualTable struct {
	// IUnknown
	QueryInterface uintptr
	addRef         uintptr
	release        uintptr
	// IConnectionPoint
	getConnectionInterface      uintptr
	GetConnectionPointContainer uintptr
	advise                      uintptr
	unadvise                    uintptr
	enumConnections             uintptr
}

// QueryInterfaceAddress returns the QueryInterface entry point for obj.
func (obj *IConnectionPoint) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

// AddRefAddress returns the AddRef entry point for obj.
func (obj *IConnectionPoint) AddRefAddress() uintptr {
	return obj.VirtualTable.addRef
}

// ReleaseAddress returns the Release entry point for obj.
func (obj *IConnectionPoint) ReleaseAddress() uintptr {
	return obj.VirtualTable.release
}

// GetConnectionInterfaceAddress returns the GetConnectionInterface entry point for obj.
func (obj *IConnectionPoint) GetConnectionInterfaceAddress() uintptr {
	return obj.VirtualTable.getConnectionInterface
}

// GetConnectionPointContainerAddress returns the container lookup entry point for obj.
func (obj *IConnectionPoint) GetConnectionPointContainerAddress() uintptr {
	return obj.VirtualTable.GetConnectionPointContainer
}

// AdviseAddress returns the Advise entry point for obj.
func (obj *IConnectionPoint) AdviseAddress() uintptr {
	return obj.VirtualTable.advise
}

// UnadviseAddress returns the Unadvise entry point for obj.
func (obj *IConnectionPoint) UnadviseAddress() uintptr {
	return obj.VirtualTable.unadvise
}

// EnumConnectionsAddress returns the EnumConnections entry point for obj.
func (obj *IConnectionPoint) EnumConnectionsAddress() uintptr {
	return obj.VirtualTable.enumConnections
}

// AddRef increments the COM reference count for obj.
func (obj *IConnectionPoint) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

// Release decrements the COM reference count for obj.
func (obj *IConnectionPoint) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

// GetConnectionInterface returns the event interface IID supported by obj.
func (obj *IConnectionPoint) GetConnectionInterface() (interfaceID windows.GUID, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.getConnectionInterface,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&interfaceID)),
		0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// GetConnectionPointContainer returns the container that owns obj.
func (obj *IConnectionPoint) GetConnectionPointContainer() (container *IConnectionPointContainer, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetConnectionPointContainer,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&container)),
		0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// Advise connects an event sink to obj and returns the subscription cookie.
func (obj *IConnectionPoint) Advise(unknown *IsIUnknown) (cookie uint32, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.advise,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(unknown)),
		uintptr(unsafe.Pointer(&cookie)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// Unadvise disconnects a sink previously registered with Advise.
func (obj *IConnectionPoint) Unadvise(cookie uint32) (err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.unadvise,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(cookie),
		0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// EnumConnections returns an enumerator over active event sink registrations.
func (obj *IConnectionPoint) EnumConnections() (connections *IEnumConnections, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.enumConnections,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&connections)),
		0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// QueryIConnectionPointFromIUnknown casts unknown to IConnectionPoint.
func QueryIConnectionPointFromIUnknown(unknown IsIUnknown) (obj *IConnectionPoint, err error) {
	if unknown == nil {
		return nil, ComInterfaceIsNilPointer
	}

	obj, err = QueryInterfaceOnIUnknown[IConnectionPoint](unknown, IID_IConnectionPoint)
	if err != nil {
		return nil, err
	}
	return
}
