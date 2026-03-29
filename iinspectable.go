//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

// TrustLevel reports the WinRT trust level returned by IInspectable.
type TrustLevel uint32

const (
	BaseTrust TrustLevel = iota
	PartialTrust
	FullTrust
)

// IsIInspectable describes values that expose the IInspectable vtable entries.
//
// Example:
//
//	name, err := ole.GetRuntimeClassNameOnIInspectable(factory)
//	if err != nil {
//		return err
//	}
//	_ = name
type IsIInspectable interface {
	GetInterfaceIdsAddress() uintptr
	GetRuntimeClassNameAddress() uintptr
	GetTrustLevelAddress() uintptr
}

// IInspectable is the base WinRT interface implemented by runtime classes.
type IInspectable struct {
	VirtualTable *IInspectableVirtualTable
}

// IInspectableVirtualTable contains the native function pointers for IInspectable.
type IInspectableVirtualTable struct {
	QueryInterface      uintptr
	AddRef              uintptr
	Release             uintptr
	GetIIds             uintptr
	GetRuntimeClassName uintptr
	GetTrustLevel       uintptr
}

// QueryInterfaceAddress returns the QueryInterface entry point for obj.
func (obj *IInspectable) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

// AddRefAddress returns the AddRef entry point for obj.
func (obj *IInspectable) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

// ReleaseAddress returns the Release entry point for obj.
func (obj *IInspectable) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

// GetInterfaceIdsAddress returns the GetIids entry point for obj.
func (obj *IInspectable) GetInterfaceIdsAddress() uintptr {
	return obj.VirtualTable.GetIIds
}

// GetRuntimeClassNameAddress returns the GetRuntimeClassName entry point for obj.
func (obj *IInspectable) GetRuntimeClassNameAddress() uintptr {
	return obj.VirtualTable.GetRuntimeClassName
}

// GetTrustLevelAddress returns the GetTrustLevel entry point for obj.
func (obj *IInspectable) GetTrustLevelAddress() uintptr {
	return obj.VirtualTable.GetTrustLevel
}

// AddRef increments the COM reference count for obj.
func (obj *IInspectable) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

// Release decrements the COM reference count for obj.
func (obj *IInspectable) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

// GetInterfaceIds returns the interface IDs implemented by obj.
func (obj *IInspectable) GetInterfaceIds() ([]windows.GUID, error) {
	return GetInterfaceIdsOnIInspectable(obj)
}

// GetRuntimeClassName returns the WinRT runtime class name for obj.
func (obj *IInspectable) GetRuntimeClassName() (string, error) {
	return GetRuntimeClassNameOnIInspectable(obj)
}

// GetTrustLevel returns the WinRT trust level reported by obj.
func (obj *IInspectable) GetTrustLevel() TrustLevel {
	return GetTrustLevelOnIInspectable(obj)
}

// GetInterfaceIdsOnIInspectable returns the interface IDs implemented by obj.
func GetInterfaceIdsOnIInspectable(obj IsIInspectable) (interfaceIds []windows.GUID, err error) {
	var count uint32
	var array *windows.GUID
	hr, _, _ := syscall.Syscall(
		obj.GetInterfaceIdsAddress(),
		3,
		comPointer(obj),
		uintptr(unsafe.Pointer(&count)),
		uintptr(unsafe.Pointer(&array)),
	)

	if windows.Handle(hr) != windows.S_OK {
		err = windows.Errno(hr)
		return
	}
	if array == nil || count == 0 {
		return nil, nil
	}
	defer TaskMemoryFreePointer(unsafe.Pointer(array))

	interfaceIds = unsafe.Slice(array, count)

	return
}

// GetRuntimeClassNameOnIInspectable returns the WinRT runtime class name for obj.
func GetRuntimeClassNameOnIInspectable(obj IsIInspectable) (s string, err error) {
	var hString HString
	hr, _, _ := syscall.Syscall(
		obj.GetRuntimeClassNameAddress(),
		2,
		comPointer(obj),
		uintptr(unsafe.Pointer(&hString)),
		0)

	if windows.Handle(hr) != windows.S_OK {
		err = windows.Errno(hr)
		return
	}
	defer DeleteHString(hString)

	s = hString.String()
	return
}

// GetTrustLevelOnIInspectable returns the WinRT trust level for obj.
func GetTrustLevelOnIInspectable(obj IsIInspectable) TrustLevel {
	var level uint32
	syscall.Syscall(
		obj.GetTrustLevelAddress(),
		2,
		comPointer(obj),
		uintptr(unsafe.Pointer(&level)),
		0)

	return TrustLevel(level)
}
