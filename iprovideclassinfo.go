//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

type IProvideClassInfo struct {
	QueryInterface uintptr
	addRef         uintptr
	release        uintptr
	getClassInfo   uintptr
}

func (v *IProvideClassInfo) QueryInterfaceAddress() uintptr {
	return v.QueryInterface
}

func (v *IProvideClassInfo) AddRefAddress() uintptr {
	return v.addRef
}

func (v *IProvideClassInfo) ReleaseAddress() uintptr {
	return v.release
}

func (obj *IProvideClassInfo) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IProvideClassInfo) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (v *IProvideClassInfo) GetClassInfo() (info *ITypeInfo, err error) {
	hr, _, _ := syscall.Syscall(
		v.getClassInfo,
		2,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(&info)),
		0)

	if hr == windows.S_OK {
		return
	}

	err = hr

	return
}
