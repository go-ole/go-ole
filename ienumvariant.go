//go:build windows

package ole

import (
	"errors"
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

type IEnumVariant struct {
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

var (
	EnumOutOfMemoryError = errors.New("IEnumVariant: OutOfMemoryError")
)

func (v *IEnumVariant) QueryInterfaceAddress() uintptr {
	return v.QueryInterface
}

func (v *IEnumVariant) AddRefAddress() uintptr {
	return v.addRef
}

func (v *IEnumVariant) ReleaseAddress() uintptr {
	return v.release
}

func (obj *IEnumVariant) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IEnumVariant) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IEnumVariant) Clone() (cloned *IEnumVariant, err error) {
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

func (obj *IEnumVariant) Reset() bool {
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

func (obj *IEnumVariant) Skip(numSkip uint) bool {
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

func (obj *IEnumVariant) Next(numRetrieve uint) (ret []*VARIANT) {
	var length int
	var array []*VARIANT
	syscall.Syscall6(
		obj.next,
		4,
		uintptr(unsafe.Pointer(obj)),
		uintptr(numRetrieve),
		uintptr(unsafe.Pointer(&array[0])),
		uintptr(unsafe.Pointer(length)),
		0,
		0)

	// New unsafe array conversion since Go 1.17.
	ret = (*[length]*VARIANT)(unsafe.Pointer(&array[0]))[:]

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
