package ole

import "unsafe"

type IUnknown struct {
	RawVTable *interface{}
}

type IUnknownVtbl struct {
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
}

type UnknownLike interface {
	QueryInterface(iid *GUID) (disp *IDispatch, err error)
	AddRef() int32
	Release() int32
}

func (v *IUnknown) VTable() *IUnknownVtbl {
	return (*IUnknownVtbl)(unsafe.Pointer(v.RawVTable))
}

func (v *IUnknown) QueryInterface(iid *GUID) (*IDispatch, error) {
	return queryInterface(v, iid)
}

func (v *IUnknown) MustQueryInterface(iid *GUID) *IDispatch {
	disp, _ = queryInterface(v, iid)
	return disp
}

func (v *IUnknown) AddRef() int32 {
	return addRef(v)
}

func (v *IUnknown) Release() int32 {
	return release(v)
}
