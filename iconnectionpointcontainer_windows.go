// +build windows

package ole

import (
	"unsafe"

	syscall "golang.org/x/sys/windows"
)

func (v *IConnectionPointContainer) EnumConnectionPoints(points interface{}) error {
	return NewError(E_NOTIMPL)
}

func (v *IConnectionPointContainer) FindConnectionPoint(iid *GUID, point **IConnectionPoint) (err error) {
	hr, _, _ := syscall.Syscall(
		v.VTable().FindConnectionPoint,
		3,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(point)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}
