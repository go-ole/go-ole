//go:build windows

package ole

import (
	"errors"
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

type IEnumVariant struct {
	VirtualTable *IEnumVariantVirtualTable
}

type IEnumVariantVirtualTable struct {
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

var (
	EnumOutOfMemoryError = errors.New("IEnumVariant: OutOfMemoryError")
)

func (v *IEnumVariant) QueryInterfaceAddress() uintptr {
	return v.VirtualTable.QueryInterface
}

func (v *IEnumVariant) AddRefAddress() uintptr {
	return v.VirtualTable.AddRef
}

func (v *IEnumVariant) ReleaseAddress() uintptr {
	return v.VirtualTable.Release
}

func (obj *IEnumVariant) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IEnumVariant) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IEnumVariant) Clone() (cloned *IEnumVariant, err error) {
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

func (obj *IEnumVariant) Reset() bool {
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

func (obj *IEnumVariant) Skip(numSkip uint) bool {
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

func (obj *IEnumVariant) Next(numRetrieve uint32) (ret []*VARIANT) {
	var length uint32
	var array []*VARIANT
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

func (obj *IEnumVariant) ForEach(callback func(v *VARIANT) error) (err error) {
	obj.Reset()
	items := obj.Next(100)
	for len(items) > 0 {
		for _, item := range items {
			err = callback(item)
			if err != nil {
				return err
			}
		}
		items = obj.Next(100)
	}
	return nil
}

func QueryIEnumVariantFromIUnknown(unknown IsIUnknown) (enum *IEnumVariant, err error) {
	if unknown == nil {
		return nil, ComInterfaceIsNilPointer
	}

	enum, err = QueryInterfaceOnIUnknown[IEnumVariant](unknown, IID_IEnumVariant)
	if err != nil {
		return nil, err
	}
	return
}
