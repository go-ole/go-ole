//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

type ITypeInfo struct {
	// IUnknown
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
	// ITypeInfo
	getTypeAttr          uintptr
	GetTypeComp          uintptr
	GetFuncDesc          uintptr
	GetVarDesc           uintptr
	GetNames             uintptr
	GetRefTypeOfImplType uintptr
	GetImplTypeFlags     uintptr
	GetIDsOfNames        uintptr
	Invoke               uintptr
	GetDocumentation     uintptr
	GetDllEntry          uintptr
	GetRefTypeInfo       uintptr
	AddressOfMember      uintptr
	CreateInstance       uintptr
	GetMops              uintptr
	GetContainingTypeLib uintptr
	ReleaseTypeAttr      uintptr
	ReleaseFuncDesc      uintptr
	ReleaseVarDesc       uintptr
}

func (obj *ITypeInfo) QueryInterfaceAddress() uintptr {
	return obj.QueryInterface
}

func (obj *ITypeInfo) AddRefAddress() uintptr {
	return obj.AddRef
}

func (obj *ITypeInfo) ReleaseAddress() uintptr {
	return obj.Release
}

// TODO: refactor to not be a function pointer.
func (obj *ITypeInfo) GetTypeAttr() (tattr *TYPEATTR, err error) {
	hr, _, _ := syscall.Syscall(
		uintptr(obj.getTypeAttr),
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&tattr)),
		0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}
