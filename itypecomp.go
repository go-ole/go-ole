package ole

import "unsafe"

type ITypeComp struct {
	IUnknown
}

type ITypeCompVtbl struct {
	IUnknownVtbl
	Bind     uintptr
	BindType uintptr
}

func (v *ITypeComp) VTable() *ITypeCompVtbl {
	return (*ITypeCompVtbl)(unsafe.Pointer(v.RawVTable))
}
