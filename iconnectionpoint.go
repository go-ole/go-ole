//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

type IConnectionPointAddresses interface {
	IsIUnknown
	GetConnectionInterfaceAddress() uintptr
	GetConnectionPointContainerAddress() uintptr
	AdviseAddress() uintptr
	UnadviseAddress() uintptr
	EnumConnectionsAddress() uintptr
}

type IConnectionPoint struct {
	VirtualTable *IConnectionPointVirtualTable
}

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

func (obj *IConnectionPoint) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *IConnectionPoint) AddRefAddress() uintptr {
	return obj.VirtualTable.addRef
}

func (obj *IConnectionPoint) ReleaseAddress() uintptr {
	return obj.VirtualTable.release
}

func (obj *IConnectionPoint) GetConnectionInterfaceAddress() uintptr {
	return obj.VirtualTable.getConnectionInterface
}

func (obj *IConnectionPoint) GetConnectionPointContainerAddress() uintptr {
	return obj.VirtualTable.GetConnectionPointContainer
}

func (obj *IConnectionPoint) AdviseAddress() uintptr {
	return obj.VirtualTable.advise
}

func (obj *IConnectionPoint) UnadviseAddress() uintptr {
	return obj.VirtualTable.unadvise
}

func (obj *IConnectionPoint) EnumConnectionsAddress() uintptr {
	return obj.VirtualTable.enumConnections
}

func (obj *IConnectionPoint) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IConnectionPoint) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

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
