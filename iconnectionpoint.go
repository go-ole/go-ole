//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
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
	EnumConnections             uintptr
}

func (v *IConnectionPoint) QueryInterfaceAddress() uintptr {
	return v.QueryInterface
}

func (v *IConnectionPoint) AddRefAddress() uintptr {
	return v.addRef
}

func (v *IConnectionPoint) ReleaseAddress() uintptr {
	return v.release
}

func (obj *IConnectionPoint) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IConnectionPoint) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IConnectionPoint) GetConnectionInterface() (interfaceID *windows.GUID, err error) {
	hr, _, _ := windows.Syscall(
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
	hr, _, _ := windows.Syscall(
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
	hr, _, _ := windows.Syscall(
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

func (v *IConnectionPoint) EnumConnections(p *unsafe.Pointer) error {
	return NewError(E_NOTIMPL)
}
