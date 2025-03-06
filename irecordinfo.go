//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

type IRecordInfo struct {
	QueryInterface uintptr
	addRef         uintptr
	release        uintptr
	// IRecordInfo
	recordInit       uintptr
	recordClear      uintptr
	recordCopy       uintptr
	getGuid          uintptr
	GetName          uintptr
	getSize          uintptr
	GetTypeInfo      uintptr
	GetField         uintptr
	GetFieldNoCopy   uintptr
	PutField         uintptr
	PutFieldNoCopy   uintptr
	GetFieldNames    uintptr
	isMatchingType   uintptr
	recordCreate     uintptr
	recordCreateCopy uintptr
	recordDestroy    uintptr
}

func (v *IRecordInfo) QueryInterfaceAddress() uintptr {
	return v.QueryInterface
}

func (v *IRecordInfo) AddRefAddress() uintptr {
	return v.addRef
}

func (v *IRecordInfo) ReleaseAddress() uintptr {
	return v.release
}

func (obj *IRecordInfo) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IRecordInfo) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IRecordInfo) GetGuid() (ret windows.GUID, err error) {
	hr, _, _ := syscall.Syscall(
		v.getGuid,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&ret)),
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.TYPE_E_INVALIDSTATE:
		return InvalidArgumentComError
	case windows.E_INVALIDARG:
		return InvalidArgumentComError
	default:
		return windows.Errno(hr)
	}
}

func (obj *IRecordInfo) GetSize() (ret uint32, err error) {
	hr, _, _ := syscall.Syscall(
		v.getSize,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&ret)),
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.E_INVALIDARG:
		return InvalidArgumentComError
	default:
		return windows.Errno(hr)
	}
}

func (obj *IRecordInfo) RecordInit() (ret uintptr, err error) {
	hr, _, _ := syscall.Syscall(
		v.recordInit,
		2,
		uintptr(unsafe.Pointer(&obj)),
		uintptr(unsafe.Pointer(&ret)),
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.E_INVALIDARG:
		return InvalidArgumentComError
	default:
		return windows.Errno(hr)
	}
}

func (obj *IRecordInfo) RecordClear(existing uintptr) (err error) {
	hr, _, _ := syscall.Syscall(
		v.recordClear,
		2,
		uintptr(unsafe.Pointer(obj)),
		existing,
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.E_INVALIDARG:
		return InvalidArgumentComError
	default:
		return windows.Errno(hr)
	}
}

func (obj *IRecordInfo) RecordCopy(existing uintptr) (copy uintptr, err error) {
	hr, _, _ := syscall.Syscall(
		v.recordCopy,
		3,
		uintptr(unsafe.Pointer(obj)),
		existing,
		uintptr(unsafe.Pointer(&copy)))

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.E_INVALIDARG:
		return InvalidArgumentComError
	default:
		return windows.Errno(hr)
	}
}

func (obj *IRecordInfo) RecordCreate() (ret uintptr, err error) {
	ret, _, err := syscall.Syscall(
		v.recordCreate,
		1,
		uintptr(unsafe.Pointer(obj)),
		0,
		0)
	return
}

func (obj *IRecordInfo) RecordCreateCopy(existing uintptr) (ret uintptr, err error) {
	hr, _, _ := syscall.Syscall(
		v.recordCreateCopy,
		3,
		uintptr(unsafe.Pointer(obj)),
		existing,
		uintptr(unsafe.Pointer(&ret)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_OUTOFMEMORY:
		err = OutOfMemoryComError
	case windows.E_INVALIDARG:
		err = InvalidArgumentComError
	default:
		err = windows.Errno(hr)
	}

	return
}

func (obj *IRecordInfo) RecordDestroy(existing uintptr) (err error) {
	hr, _, _ := syscall.Syscall(
		v.recordDestroy,
		2,
		uintptr(unsafe.Pointer(obj)),
		existing,
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.E_INVALIDARG:
		return InvalidArgumentComError
	default:
		return windows.Errno(hr)
	}
}

func (obj *IRecordInfo) Equals(recordInfo *IRecordInfo) (ret bool, err error) {
	hr, _, err := syscall.Syscall(
		v.isMatchingType,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&recordInfo)),
		0)

	ret = hr != 0

	return
}
