//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

type IProvideClassInfo struct {
	VirtualTable *IProvideClassInfoVirtualTable
}

type IProvideClassInfoVirtualTable struct {
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
	GetClassInfo   uintptr
}

func (obj *IProvideClassInfo) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *IProvideClassInfo) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

func (obj *IProvideClassInfo) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

func (obj *IProvideClassInfo) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IProvideClassInfo) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

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
