//go:build windows

package ole

import (
	"unsafe"

	"golang.org/x/sys/windows"
)

type IActivationFactory struct {
	QueryInterface      uintptr
	addRef              uintptr
	release             uintptr
	getIIds             uintptr
	getRuntimeClassName uintptr
	getTrustLevel       uintptr
	activateInstance    uintptr
}

func (obj *IActivationFactory) QueryInterfaceAddress() uintptr {
	return obj.QueryInterface
}

func (obj *IActivationFactory) AddRefAddress() uintptr {
	return obj.addRef
}

func (obj *IActivationFactory) ReleaseAddress() uintptr {
	return obj.release
}

func (obj *IActivationFactory) GetInterfaceIdsAddress() uintptr {
	return obj.getIIds
}

func (obj *IActivationFactory) GetRuntimeClassNameAddress() uintptr {
	return obj.getRuntimeClassName
}

func (obj *IActivationFactory) GetTrustLevelAddress() uintptr {
	return obj.getTrustLevel
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
	hr, _, _ := windows.Syscall(
		obj.activateInstance,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&ret)),
		0,
	)

	if hr != windows.S_OK {
		err = hr
	}

	return
}
