//go:build windows

package ole

import (
	"errors"
	"golang.org/x/sys/windows"
	"syscall"
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
		v.clone,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&cloned)),
		0)

	switch hr {
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

	switch hr {
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
		enum.skip,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(numSkip),
		0)

	switch hr {
	case windows.S_OK:
		return true
	case windows.S_FALSE:
		return false
	default:
		return false
	}
}

func (obj *IEnumVariant) Next(numRetrieve uint) (array VARIANT, length uint, hasLess bool) {
	hr, _, _ := syscall.Syscall6(
		v.next,
		4,
		uintptr(unsafe.Pointer(obj)),
		uintptr(numRetrieve),
		uintptr(unsafe.Pointer(&array)),
		uintptr(unsafe.Pointer(&length)),
		0,
		0)

	switch hr {
	case windows.S_OK:
		hasLess = false
	case windows.S_FALSE:
		hasLess = true
	default:
		hasLess = false
	}
	return
}

func (v *IEnumVariant) ForEach(callback func(v *VARIANT) error) error {
	v.Reset()
	for item, length, hasLess := v.Next(1); length > 0; item, length, err = v.Next(1) {
		if hasLess {
			return nil
		}
		err = callback(&item)
		if err != nil {
			return err
		}
	}
	return nil
}

func QueryIEnumVariantFromIUnknown(unknown *IsIUnknown) (enum *IEnumVariant, err error) {
	if unknown == nil {
		return nil, ComInterfaceIsNilPointer
	}

	enum, err = QueryInterfaceOnIUnknown[IEnumVariant](unknown, IID_IEnumVariant)
	if err != nil {
		return nil, err
	}
	return
}
