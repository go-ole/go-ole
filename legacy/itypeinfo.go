package legacy

import (
	"github.com/go-ole/go-ole"
	"unsafe"
)

type ITypeInfo struct {
	ole.IUnknown
}

type ITypeInfoVtbl struct {
	ole.IUnknownVtbl
	GetTypeAttr          uintptr
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

func (v *ITypeInfo) VTable() *ITypeInfoVtbl {
	return (*ITypeInfoVtbl)(unsafe.Pointer(v.RawVTable))
}
