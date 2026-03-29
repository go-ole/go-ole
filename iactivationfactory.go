//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

// IActivationFactoryAddresses describes the WinRT activation factory vtable entries.
//
// Example:
//
//	factory, err := ole.RoGetActivationFactory("Windows.Foundation.Uri", ole.IID_IActivationFactory)
//	if err != nil {
//		return err
//	}
//	defer factory.Release()
type IActivationFactoryAddresses interface {
	IsIInspectable
	ActivateInstanceAddress() uintptr
}

// IActivationFactory represents the WinRT IActivationFactory interface.
type IActivationFactory struct {
	VirtualTable *IActivationFactoryVirtualTable
}

// IActivationFactoryVirtualTable contains the native function pointers for IActivationFactory.
type IActivationFactoryVirtualTable struct {
	QueryInterface      uintptr
	addRef              uintptr
	release             uintptr
	getIIds             uintptr
	getRuntimeClassName uintptr
	getTrustLevel       uintptr
	activateInstance    uintptr
}

// QueryInterfaceAddress returns the QueryInterface entry point for obj.
func (obj *IActivationFactory) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

// AddRefAddress returns the AddRef entry point for obj.
func (obj *IActivationFactory) AddRefAddress() uintptr {
	return obj.VirtualTable.addRef
}

// ReleaseAddress returns the Release entry point for obj.
func (obj *IActivationFactory) ReleaseAddress() uintptr {
	return obj.VirtualTable.release
}

// GetInterfaceIdsAddress returns the GetIids entry point for obj.
func (obj *IActivationFactory) GetInterfaceIdsAddress() uintptr {
	return obj.VirtualTable.getIIds
}

// GetRuntimeClassNameAddress returns the GetRuntimeClassName entry point for obj.
func (obj *IActivationFactory) GetRuntimeClassNameAddress() uintptr {
	return obj.VirtualTable.getRuntimeClassName
}

// GetTrustLevelAddress returns the GetTrustLevel entry point for obj.
func (obj *IActivationFactory) GetTrustLevelAddress() uintptr {
	return obj.VirtualTable.getTrustLevel
}

// ActivateInstanceAddress returns the ActivateInstance entry point for obj.
func (obj *IActivationFactory) ActivateInstanceAddress() uintptr {
	return obj.VirtualTable.activateInstance
}

// AddRef increments the COM reference count for obj.
func (obj *IActivationFactory) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

// Release decrements the COM reference count for obj.
func (obj *IActivationFactory) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

// GetInterfaceIds returns the interface IDs implemented by obj.
func (obj *IActivationFactory) GetInterfaceIds() ([]windows.GUID, error) {
	return GetInterfaceIdsOnIInspectable(obj)
}

// GetRuntimeClassName returns the runtime class name reported by obj.
func (obj *IActivationFactory) GetRuntimeClassName() (string, error) {
	return GetRuntimeClassNameOnIInspectable(obj)
}

// GetTrustLevel returns the trust level reported by obj.
func (obj *IActivationFactory) GetTrustLevel() TrustLevel {
	return GetTrustLevelOnIInspectable(obj)
}

// ActivateInstance constructs a new WinRT instance through the factory.
func (obj *IActivationFactory) ActivateInstance() (ret *IInspectable, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.activateInstance,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&ret)),
		0,
	)

	if hr != 0 {
		err = windows.Errno(hr)
	}

	return
}
