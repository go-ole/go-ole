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
	VirtualTable *IEnumConnectionsVirtualTable
}

type IEnumConnectionsVirtualTable struct {
	// IUnknown
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
	// IEnumVARIANT
	Next  uintptr
	Skip  uintptr
	Reset uintptr
	Clone uintptr
}

func (obj *IEnumConnections) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *IEnumConnections) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

func (obj *IEnumConnections) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

func (obj *IEnumConnections) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IEnumConnections) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IEnumConnections) Clone() (cloned *IEnumConnections, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.Clone,
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
		obj.VirtualTable.Reset,
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
		obj.VirtualTable.Skip,
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

func (obj *IEnumConnections) Next(numRetrieve uint32) (ret []ConnectData) {
	var length uint32
	var array []ConnectData
	syscall.Syscall6(
		obj.VirtualTable.Next,
		4,
		uintptr(unsafe.Pointer(obj)),
		uintptr(numRetrieve),
		uintptr(unsafe.Pointer(&array[0])),
		uintptr(unsafe.Pointer(&length)),
		0,
		0)

	ret = unsafe.Slice(&array[0], length)

	return
}

func (v *IEnumConnections) ForEach(callback func(v *ConnectData) error) (err error) {
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
