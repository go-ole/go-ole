//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

type ConnectData struct {
	unknown uintptr
	Cookie  uint32
}

type IEnumConnections struct {
	// IUnknown
	QueryInterface uintptr
	addRef         uintptr
	release        uintptr
	// IEnumVARIANT
	next  uintptr
	skip  uintptr
	reset uintptr
	clone uintptr
}

func (obj *IEnumConnections) QueryInterfaceAddress() uintptr {
	return obj.QueryInterface
}

func (obj *IEnumConnections) AddRefAddress() uintptr {
	return obj.addRef
}

func (obj *IEnumConnections) ReleaseAddress() uintptr {
	return obj.release
}

func (obj *IEnumConnections) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IEnumConnections) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IEnumConnections) Clone() (cloned *IEnumConnections, err error) {
	hr, _, _ := syscall.Syscall(
		obj.clone,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&cloned)),
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.E_OUTOFMEMORY:
		return nil, EnumOutOfMemoryError
	default:
		return cloned, windows.Errno(hr)
	}
}

func (obj *IEnumConnections) Reset() bool {
	hr, _, _ := syscall.Syscall(
		obj.reset,
		1,
		uintptr(unsafe.Pointer(obj)),
		0,
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return true
	case windows.S_FALSE:
		return false
	default:
		return false
	}
}

func (obj *IEnumConnections) Skip(numSkip uint) bool {
	hr, _, _ := syscall.Syscall(
		obj.skip,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(numSkip),
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return true
	case windows.S_FALSE:
		return false
	default:
		return false
	}
}

func (obj *IEnumConnections) Next(numRetrieve uint) (connectData []ConnectData) {
	var length int
	var array []ConnectData
	syscall.Syscall6(
		obj.next,
		4,
		uintptr(unsafe.Pointer(obj)),
		uintptr(numRetrieve),
		uintptr(unsafe.Pointer(&array)),
		uintptr(unsafe.Pointer(&length)),
		0,
		0)

	// New unsafe array conversion since Go 1.17.
	connectData = (*[length]ConnectData)(unsafe.Pointer(array))[:]

	return
}

func (v *IEnumConnections) ForEach(callback func(v *VARIANT) error) (err error) {
	v.Reset()
	items := v.Next(100)
	for len(items) > 0 {
		for _, item := range items {
			err = callback(&item)
			if err != nil {
				return err
			}
		}
		items = v.Next(100)
	}
	return nil
}
