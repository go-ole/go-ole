package legacy

import (
	"github.com/go-ole/go-ole"
	"unsafe"
)

type IProvideClassInfo struct {
	ole.IUnknown
}

type IProvideClassInfoVtbl struct {
	ole.IUnknownVtbl
	GetClassInfo uintptr
}

func (v *IProvideClassInfo) VTable() *IProvideClassInfoVtbl {
	return (*IProvideClassInfoVtbl)(unsafe.Pointer(v.RawVTable))
}

func (v *IProvideClassInfo) GetClassInfo() (cinfo *ITypeInfo, err error) {
	cinfo, err = getClassInfo(v)
	return
}
