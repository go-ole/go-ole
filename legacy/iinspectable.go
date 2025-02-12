package legacy

import (
	"github.com/go-ole/go-ole"
	"unsafe"
)

type IInspectable struct {
	ole.IUnknown
}

type IInspectableVtbl struct {
	ole.IUnknownVtbl
	GetIIds             uintptr
	GetRuntimeClassName uintptr
	GetTrustLevel       uintptr
}

func (v *IInspectable) VTable() *IInspectableVtbl {
	return (*IInspectableVtbl)(unsafe.Pointer(v.RawVTable))
}
