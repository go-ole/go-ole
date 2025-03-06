//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

type IActivationFactory struct {
	VirtualTable *IActivationFactoryVirtualTable
}

type IActivationFactoryVirtualTable struct {
	QueryInterface      uintptr
	addRef              uintptr
	release             uintptr
	getIIds             uintptr
	getRuntimeClassName uintptr
	getTrustLevel       uintptr
	activateInstance    uintptr
}

func (obj *IActivationFactory) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *IActivationFactory) AddRefAddress() uintptr {
	return obj.VirtualTable.addRef
}

func (obj *IActivationFactory) ReleaseAddress() uintptr {
	return obj.VirtualTable.release
}

func (obj *IActivationFactory) GetInterfaceIdsAddress() uintptr {
	return obj.VirtualTable.getIIds
}

func (obj *IActivationFactory) GetRuntimeClassNameAddress() uintptr {
	return obj.VirtualTable.getRuntimeClassName
}

func (obj *IActivationFactory) GetTrustLevelAddress() uintptr {
	return obj.VirtualTable.getTrustLevel
}

func (obj *IActivationFactory) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IActivationFactory) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IActivationFactory) GetInterfaceIds() ([]windows.GUID, error) {
	return GetInterfaceIdsOnIInspectable(obj)
}

func (obj *IActivationFactory) GetRuntimeClassName() (string, error) {
	return GetRuntimeClassNameOnIInspectable(obj)
}

func (obj *IActivationFactory) GetTrustLevel() TrustLevel {
	return GetTrustLevelOnIInspectable(obj)
}

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
