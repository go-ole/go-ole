package legacy

import (
	"github.com/go-ole/go-ole"
	"unsafe"
)

type IConnectionPointContainer struct {
	ole.IUnknown
}

type IConnectionPointContainerVtbl struct {
	ole.IUnknownVtbl
	EnumConnectionPoints uintptr
	FindConnectionPoint  uintptr
}

func (v *IConnectionPointContainer) VTable() *IConnectionPointContainerVtbl {
	return (*IConnectionPointContainerVtbl)(unsafe.Pointer(v.RawVTable))
}
