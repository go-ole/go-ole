//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

// ConnectData mirrors the COM CONNECTDATA structure returned by IEnumConnections.
type ConnectData struct {
	unknown uintptr
	Cookie  uint32
}

// IEnumConnectionsAddresses describes the IEnumConnections vtable entries.
//
// Example:
//
//	enum, err := point.EnumConnections()
//	if err != nil {
//		return err
//	}
//	defer enum.Release()
type IEnumConnectionsAddresses interface {
	IsIUnknown
	NextAddress() uintptr
	SkipAddress() uintptr
	ResetAddress() uintptr
	CloneAddress() uintptr
}

// IEnumConnections represents the COM IEnumConnections enumerator interface.
type IEnumConnections struct {
	VirtualTable *IEnumConnectionsVirtualTable
}

// IEnumConnectionsVirtualTable contains the native function pointers for IEnumConnections.
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

// QueryInterfaceAddress returns the QueryInterface entry point for obj.
func (obj *IEnumConnections) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

// AddRefAddress returns the AddRef entry point for obj.
func (obj *IEnumConnections) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

// ReleaseAddress returns the Release entry point for obj.
func (obj *IEnumConnections) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

// NextAddress returns the Next entry point for obj.
func (obj *IEnumConnections) NextAddress() uintptr {
	return obj.VirtualTable.Next
}

// SkipAddress returns the Skip entry point for obj.
func (obj *IEnumConnections) SkipAddress() uintptr {
	return obj.VirtualTable.Skip
}

// ResetAddress returns the Reset entry point for obj.
func (obj *IEnumConnections) ResetAddress() uintptr {
	return obj.VirtualTable.Reset
}

// CloneAddress returns the Clone entry point for obj.
func (obj *IEnumConnections) CloneAddress() uintptr {
	return obj.VirtualTable.Clone
}

// AddRef increments the COM reference count for obj.
func (obj *IEnumConnections) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

// Release decrements the COM reference count for obj.
func (obj *IEnumConnections) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

// Clone duplicates the enumeration state of obj.
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

// Reset moves obj back to the start of the enumeration.
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

// Skip advances obj by numSkip elements.
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

// Next retrieves up to numRetrieve connection records from obj.
func (obj *IEnumConnections) Next(numRetrieve uint32) (ret []ConnectData) {
	if numRetrieve == 0 {
		return nil
	}

	var length uint32
	array := make([]ConnectData, numRetrieve)
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

// ForEach calls yield for each connection until the enumeration ends or yield returns false.
func (v *IEnumConnections) ForEach(yield func(v *ConnectData) bool) {
	v.Reset()
	items := v.Next(100)
	for len(items) > 0 {
		for index := range items {
			if !yield(&items[index]) {
				return
			}
		}
		items = v.Next(100)
	}
}
