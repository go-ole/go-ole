//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

// IProvideClassInfoAddresses describes the IProvideClassInfo vtable entries.
//
// Example:
//
//	info, err := ole.QueryIProvideClassInfoFromIUnknown(dispatch)
//	if err != nil {
//		return err
//	}
//	defer info.Release()
type IProvideClassInfoAddresses interface {
	IsIUnknown
	GetClassInfoAddress() uintptr
}

// IProvideClassInfo represents the COM IProvideClassInfo interface.
type IProvideClassInfo struct {
	VirtualTable *IProvideClassInfoVirtualTable
}

// IProvideClassInfoVirtualTable contains the native function pointers for IProvideClassInfo.
type IProvideClassInfoVirtualTable struct {
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
	GetClassInfo   uintptr
}

// QueryInterfaceAddress returns the QueryInterface entry point for obj.
func (obj *IProvideClassInfo) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

// AddRefAddress returns the AddRef entry point for obj.
func (obj *IProvideClassInfo) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

// ReleaseAddress returns the Release entry point for obj.
func (obj *IProvideClassInfo) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

// GetClassInfoAddress returns the GetClassInfo entry point for obj.
func (obj *IProvideClassInfo) GetClassInfoAddress() uintptr {
	return obj.VirtualTable.GetClassInfo
}

// AddRef increments the COM reference count for obj.
func (obj *IProvideClassInfo) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

// Release decrements the COM reference count for obj.
func (obj *IProvideClassInfo) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

// GetClassInfo returns the type information for the COM class that owns obj.
func (obj *IProvideClassInfo) GetClassInfo() (info *ITypeInfo, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetClassInfo,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&info)),
		0)

	if windows.Handle(hr) == windows.S_OK {
		return
	}

	err = windows.Errno(hr)

	return
}
