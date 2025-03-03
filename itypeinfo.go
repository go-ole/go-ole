//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"unsafe"
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

func (v *ITypeInfo) QueryInterfaceAddress() uintptr {
	return v.QueryInterface
}

func (v *ITypeInfo) AddRefAddress() uintptr {
	return v.AddRef
}

func (v *ITypeInfo) ReleaseAddress() uintptr {
	return v.Release
}

// TODO: refactor to not be a function pointer.
func (v *ITypeInfo) GetTypeAttr() (tattr *TYPEATTR, err error) {
	hr, _, _ := windows.Syscall(
		uintptr(v.getTypeAttr),
		2,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(&tattr)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}
