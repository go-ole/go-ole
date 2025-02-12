//go:build windows
// +build windows

package legacy

import (
	"github.com/go-ole/go-ole"
	"syscall"
	"unsafe"
)

func (v *IConnectionPoint) GetConnectionInterface(piid **GUID) int32 {
	// XXX: This doesn't look like it does what it's supposed to
	return release((*ole.IUnknown)(unsafe.Pointer(v)))
}

func (v *IConnectionPoint) Advise(unknown *ole.IUnknown) (cookie uint32, err error) {
	hr, _, _ := syscall.Syscall(
		v.VTable().Advise,
		3,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(unknown)),
		uintptr(unsafe.Pointer(&cookie)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func (v *IConnectionPoint) Unadvise(cookie uint32) (err error) {
	hr, _, _ := syscall.Syscall(
		v.VTable().Unadvise,
		2,
		uintptr(unsafe.Pointer(v)),
		uintptr(cookie),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func (v *IConnectionPoint) EnumConnections(p *unsafe.Pointer) error {
	return NewError(E_NOTIMPL)
}
