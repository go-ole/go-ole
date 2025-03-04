//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

type IConnectionPoint struct {
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
	return obj.QueryInterface
}

func (obj *IConnectionPoint) AddRefAddress() uintptr {
	return obj.addRef
}

func (obj *IConnectionPoint) ReleaseAddress() uintptr {
	return obj.release
}

func (obj *IConnectionPoint) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IConnectionPoint) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IConnectionPoint) GetConnectionInterface() (interfaceID windows.GUID, err error) {
	hr, _, _ := syscall.Syscall(
		obj.getConnectionInterface,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&interfaceID)),
		0)
	if hr != windows.S_OK {
		err = hr
	}
	return
}

func (obj *IConnectionPoint) Advise(unknown *IsIUnknown) (cookie uint32, err error) {
	hr, _, _ := syscall.Syscall(
		obj.advise,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(unknown)),
		uintptr(unsafe.Pointer(&cookie)))
	if hr != windows.S_OK {
		err = hr
	}
	return
}

func (obj *IConnectionPoint) Unadvise(cookie uint32) (err error) {
	hr, _, _ := syscall.Syscall(
		obj.unadvise,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(cookie),
		0)
	if hr != windows.S_OK {
		err = hr
	}
	return
}

func (obj *IConnectionPoint) EnumConnections(p *unsafe.Pointer) error {
	return windows.E_NOTIMPL
}
