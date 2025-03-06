//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

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

func (obj *IConnectionPoint) EnumConnections(p *unsafe.Pointer) error {
	return MethodNotImplementedError
}
