package legacy

import (
	"github.com/go-ole/go-ole"
	"unsafe"
)

type IConnectionPoint struct {
	ole.IUnknown
}

type IConnectionPointVtbl struct {
	ole.IUnknownVtbl
	GetConnectionInterface      uintptr
	GetConnectionPointContainer uintptr
	Advise                      uintptr
	Unadvise                    uintptr
	EnumConnections             uintptr
}

func (v *IConnectionPoint) VTable() *IConnectionPointVtbl {
	return (*IConnectionPointVtbl)(unsafe.Pointer(v.RawVTable))
}
