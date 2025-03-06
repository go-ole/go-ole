//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

type IRecordInfo struct {
	VirtualTable *IRecordInfoVirtualTable
}

type IRecordInfoVirtualTable struct {
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
	// IRecordInfo
	RecordInit       uintptr
	RecordClear      uintptr
	RecordCopy       uintptr
	GetGuid          uintptr
	GetName          uintptr
	GetSize          uintptr
	GetTypeInfo      uintptr
	GetField         uintptr
	GetFieldNoCopy   uintptr
	PutField         uintptr
	PutFieldNoCopy   uintptr
	GetFieldNames    uintptr
	IsMatchingType   uintptr
	RecordCreate     uintptr
	RecordCreateCopy uintptr
	RecordDestroy    uintptr
}

func (obj *IRecordInfo) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *IRecordInfo) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

func (obj *IRecordInfo) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

func (obj *IRecordInfo) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IRecordInfo) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IRecordInfo) GetGuid() (ret windows.GUID, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetGuid,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&ret)),
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.TYPE_E_INVALIDSTATE:
		err = InvalidArgumentComError
	case windows.E_INVALIDARG:
		err = InvalidArgumentComError
	default:
		err = windows.Errno(hr)
	}
	return
}

func (obj *IRecordInfo) GetSize() (ret uint32, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetSize,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&ret)),
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.E_INVALIDARG:
		err = InvalidArgumentComError
	default:
		err = windows.Errno(hr)
	}
	return
}

func (obj *IRecordInfo) RecordInit() (ret uintptr, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.RecordInit,
		2,
		uintptr(unsafe.Pointer(&obj)),
		uintptr(unsafe.Pointer(&ret)),
		0)

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.E_INVALIDARG:
		err = InvalidArgumentComError
	default:
		err = windows.Errno(hr)
	}
	return
}

func (obj *IRecordInfo) RecordClear(existing uintptr) (err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.RecordClear,
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
		obj.VirtualTable.RecordCopy,
		3,
		uintptr(unsafe.Pointer(obj)),
		existing,
		uintptr(unsafe.Pointer(&copy)))

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.E_INVALIDARG:
		err = InvalidArgumentComError
	default:
		err = windows.Errno(hr)
	}
	return
}

func (obj *IRecordInfo) RecordCreate() (ret uintptr, err error) {
	ret, _, err = syscall.Syscall(
		obj.VirtualTable.RecordCreate,
		1,
		uintptr(unsafe.Pointer(obj)),
		0,
		0)
	return
}

func (obj *IRecordInfo) RecordCreateCopy(existing uintptr) (ret uintptr, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.RecordCreateCopy,
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
		obj.VirtualTable.RecordDestroy,
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
		obj.VirtualTable.IsMatchingType,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&recordInfo)),
		0)

	ret = hr != 0

	return
}
