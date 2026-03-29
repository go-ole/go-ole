//go:build windows

package ole

import (
	"errors"
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

// IEnumVariantAddresses describes the IEnumVARIANT vtable entries.
//
// Example:
//
//	enum, err := ole.QueryIEnumVariantFromIUnknown(collection)
//	if err != nil {
//		return err
//	}
//	defer enum.Release()
type IEnumVariantAddresses interface {
	IsIUnknown
	NextAddress() uintptr
	SkipAddress() uintptr
	ResetAddress() uintptr
	CloneAddress() uintptr
}

// IEnumVariant represents the COM IEnumVARIANT enumerator interface.
type IEnumVariant struct {
	VirtualTable *IEnumVariantVirtualTable
}

// IEnumVariantVirtualTable contains the native function pointers for IEnumVARIANT.
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

// QueryInterfaceAddress returns the QueryInterface entry point for v.
func (v *IEnumVariant) QueryInterfaceAddress() uintptr {
	return v.VirtualTable.QueryInterface
}

// AddRefAddress returns the AddRef entry point for v.
func (v *IEnumVariant) AddRefAddress() uintptr {
	return v.VirtualTable.AddRef
}

// ReleaseAddress returns the Release entry point for v.
func (v *IEnumVariant) ReleaseAddress() uintptr {
	return v.VirtualTable.Release
}

// NextAddress returns the Next entry point for obj.
func (obj *IEnumVariant) NextAddress() uintptr {
	return obj.VirtualTable.Next
}

// SkipAddress returns the Skip entry point for obj.
func (obj *IEnumVariant) SkipAddress() uintptr {
	return obj.VirtualTable.Skip
}

// ResetAddress returns the Reset entry point for obj.
func (obj *IEnumVariant) ResetAddress() uintptr {
	return obj.VirtualTable.Reset
}

// CloneAddress returns the Clone entry point for obj.
func (obj *IEnumVariant) CloneAddress() uintptr {
	return obj.VirtualTable.Clone
}

// AddRef increments the COM reference count for obj.
func (obj *IEnumVariant) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

// Release decrements the COM reference count for obj.
func (obj *IEnumVariant) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

// Clone duplicates the enumeration state of obj.
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

// Reset moves obj back to the start of the enumeration.
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

// Skip advances obj by numSkip elements.
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

// Next retrieves up to numRetrieve elements from obj.
func (obj *IEnumVariant) Next(numRetrieve uint32) (ret []*VARIANT) {
	if numRetrieve == 0 {
		return nil
	}

	var length uint32
	array := make([]*VARIANT, numRetrieve)
	syscall.Syscall6(
		obj.VirtualTable.Next,
		4,
		uintptr(unsafe.Pointer(obj)),
		uintptr(numRetrieve),
		uintptr(unsafe.Pointer(&array[0])),
		uintptr(unsafe.Pointer(&length)),
		0,
		0)

	ret = array[:length]

	return
}

// ForEach calls yield for each item until the enumeration ends or yield returns false.
func (obj *IEnumVariant) ForEach(yield func(v *VARIANT) bool) {
	obj.Reset()
	items := obj.Next(100)
	for len(items) > 0 {
		for _, item := range items {
			if !yield(item) {
				return
			}
		}
		items = obj.Next(100)
	}
}

// QueryIEnumVariantFromIUnknown casts unknown to IEnumVARIANT.
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
