//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

type TrustLevel uint32

const (
	BaseTrust TrustLevel = iota
	PartialTrust
	FullTrust
)

type IsIInspectable interface {
	GetInterfaceIdsAddress() uintptr
	GetRuntimeClassNameAddress() uintptr
	GetTrustLevelAddress() uintptr
}

type IInspectable struct {
	VirtualTable *IInspectableVirtualTable
}

type IInspectableVirtualTable struct {
	QueryInterface      uintptr
	AddRef              uintptr
	Release             uintptr
	GetIIds             uintptr
	GetRuntimeClassName uintptr
	GetTrustLevel       uintptr
}

func (obj *IInspectable) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *IInspectable) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

func (obj *IInspectable) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

func (obj *IInspectable) GetInterfaceIdsAddress() uintptr {
	return obj.VirtualTable.GetIIds
}

func (obj *IInspectable) GetRuntimeClassNameAddress() uintptr {
	return obj.VirtualTable.GetRuntimeClassName
}

func (obj *IInspectable) GetTrustLevelAddress() uintptr {
	return obj.VirtualTable.GetTrustLevel
}

func (obj *IInspectable) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IInspectable) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IInspectable) GetInterfaceIds() ([]windows.GUID, error) {
	return GetInterfaceIdsOnIInspectable(obj)
}

func (obj *IInspectable) GetRuntimeClassName() (string, error) {
	return GetRuntimeClassNameOnIInspectable(obj)
}

func (obj *IInspectable) GetTrustLevel() TrustLevel {
	return GetTrustLevelOnIInspectable(obj)
}

func GetInterfaceIdsOnIInspectable(obj IsIInspectable) (interfaceIds []windows.GUID, err error) {
	var count uint32
	var array []windows.GUID
	hr, _, _ := syscall.Syscall(
		obj.GetInterfaceIdsAddress(),
		3,
		uintptr(unsafe.Pointer(&obj)),
		uintptr(unsafe.Pointer(&count)),
		uintptr(unsafe.Pointer(&array[0])),
	)

	if windows.Handle(hr) != windows.S_OK {
		err = windows.Errno(hr)
		return
	}
	defer TaskMemoryFreePointer(unsafe.Pointer(&array[0]))

	interfaceIds = unsafe.Slice(&array[0], count)

	return
}

func GetRuntimeClassNameOnIInspectable(obj IsIInspectable) (s string, err error) {
	var hString HString
	hr, _, _ := syscall.Syscall(
		obj.GetRuntimeClassNameAddress(),
		2,
		uintptr(unsafe.Pointer(&obj)),
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

func GetTrustLevelOnIInspectable(obj IsIInspectable) TrustLevel {
	var level uint32
	syscall.Syscall(
		obj.GetTrustLevelAddress(),
		2,
		uintptr(unsafe.Pointer(&obj)),
		uintptr(unsafe.Pointer(&level)),
		0)

	return TrustLevel(level)
}
