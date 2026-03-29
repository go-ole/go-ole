//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

// IRecordInfoAddresses describes the IRecordInfo vtable entries.
//
// Example:
//
//	size, err := recordInfo.GetSize()
//	if err != nil {
//		return err
//	}
//	_ = size
type IRecordInfoAddresses interface {
	IsIUnknown
	RecordInitAddress() uintptr
	RecordClearAddress() uintptr
	RecordCopyAddress() uintptr
	GetGuidAddress() uintptr
	GetNameAddress() uintptr
	GetSizeAddress() uintptr
	GetTypeInfoAddress() uintptr
	GetFieldAddress() uintptr
	GetFieldNoCopyAddress() uintptr
	PutFieldAddress() uintptr
	PutFieldNoCopyAddress() uintptr
	GetFieldNamesAddress() uintptr
	IsMatchingTypeAddress() uintptr
	RecordCreateAddress() uintptr
	RecordCreateCopyAddress() uintptr
	RecordDestroyAddress() uintptr
}

// IRecordInfo represents the COM IRecordInfo interface for user-defined record types.
type IRecordInfo struct {
	VirtualTable *IRecordInfoVirtualTable
}

// IRecordInfoVirtualTable contains the native function pointers for IRecordInfo.
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

// QueryInterfaceAddress returns the QueryInterface entry point for obj.
func (obj *IRecordInfo) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

// AddRefAddress returns the AddRef entry point for obj.
func (obj *IRecordInfo) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

// ReleaseAddress returns the Release entry point for obj.
func (obj *IRecordInfo) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

// RecordInitAddress returns the RecordInit entry point for obj.
func (obj *IRecordInfo) RecordInitAddress() uintptr {
	return obj.VirtualTable.RecordInit
}

// RecordClearAddress returns the RecordClear entry point for obj.
func (obj *IRecordInfo) RecordClearAddress() uintptr {
	return obj.VirtualTable.RecordClear
}

// RecordCopyAddress returns the RecordCopy entry point for obj.
func (obj *IRecordInfo) RecordCopyAddress() uintptr {
	return obj.VirtualTable.RecordCopy
}

// GetGuidAddress returns the GetGuid entry point for obj.
func (obj *IRecordInfo) GetGuidAddress() uintptr {
	return obj.VirtualTable.GetGuid
}

// GetNameAddress returns the GetName entry point for obj.
func (obj *IRecordInfo) GetNameAddress() uintptr {
	return obj.VirtualTable.GetName
}

// GetSizeAddress returns the GetSize entry point for obj.
func (obj *IRecordInfo) GetSizeAddress() uintptr {
	return obj.VirtualTable.GetSize
}

// GetTypeInfoAddress returns the GetTypeInfo entry point for obj.
func (obj *IRecordInfo) GetTypeInfoAddress() uintptr {
	return obj.VirtualTable.GetTypeInfo
}

// GetFieldAddress returns the GetField entry point for obj.
func (obj *IRecordInfo) GetFieldAddress() uintptr {
	return obj.VirtualTable.GetField
}

// GetFieldNoCopyAddress returns the GetFieldNoCopy entry point for obj.
func (obj *IRecordInfo) GetFieldNoCopyAddress() uintptr {
	return obj.VirtualTable.GetFieldNoCopy
}

// PutFieldAddress returns the PutField entry point for obj.
func (obj *IRecordInfo) PutFieldAddress() uintptr {
	return obj.VirtualTable.PutField
}

// PutFieldNoCopyAddress returns the PutFieldNoCopy entry point for obj.
func (obj *IRecordInfo) PutFieldNoCopyAddress() uintptr {
	return obj.VirtualTable.PutFieldNoCopy
}

// GetFieldNamesAddress returns the GetFieldNames entry point for obj.
func (obj *IRecordInfo) GetFieldNamesAddress() uintptr {
	return obj.VirtualTable.GetFieldNames
}

// IsMatchingTypeAddress returns the IsMatchingType entry point for obj.
func (obj *IRecordInfo) IsMatchingTypeAddress() uintptr {
	return obj.VirtualTable.IsMatchingType
}

// RecordCreateAddress returns the RecordCreate entry point for obj.
func (obj *IRecordInfo) RecordCreateAddress() uintptr {
	return obj.VirtualTable.RecordCreate
}

// RecordCreateCopyAddress returns the RecordCreateCopy entry point for obj.
func (obj *IRecordInfo) RecordCreateCopyAddress() uintptr {
	return obj.VirtualTable.RecordCreateCopy
}

// RecordDestroyAddress returns the RecordDestroy entry point for obj.
func (obj *IRecordInfo) RecordDestroyAddress() uintptr {
	return obj.VirtualTable.RecordDestroy
}

// AddRef increments the COM reference count for obj.
func (obj *IRecordInfo) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

// Release decrements the COM reference count for obj.
func (obj *IRecordInfo) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

// GetGuid returns the GUID associated with the record type.
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

// GetSize returns the size in bytes of the record type.
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

// RecordInit initializes the record memory referenced by newRecord.
func (obj *IRecordInfo) RecordInit(newRecord uintptr) (err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.RecordInit,
		2,
		uintptr(unsafe.Pointer(obj)),
		newRecord,
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

// RecordClear clears the record memory referenced by existing.
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

// RecordCopy copies the record from existing into newRecord.
func (obj *IRecordInfo) RecordCopy(existing uintptr, newRecord uintptr) (err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.RecordCopy,
		3,
		uintptr(unsafe.Pointer(obj)),
		existing,
		newRecord)

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

// RecordCreate allocates and initializes a new record instance.
func (obj *IRecordInfo) RecordCreate() (ret uintptr, err error) {
	ret, _, err = syscall.Syscall(
		obj.VirtualTable.RecordCreate,
		1,
		uintptr(unsafe.Pointer(obj)),
		0,
		0)
	return
}

// RecordCreateCopy allocates a new record and copies existing into it.
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

// RecordDestroy destroys and frees the record referenced by existing.
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

// IsMatchingType reports whether recordInfo describes the same record type as obj.
func (obj *IRecordInfo) IsMatchingType(recordInfo *IRecordInfo) bool {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.IsMatchingType,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(recordInfo)),
		0)

	return hr != 0
}

// Equals reports whether recordInfo is the same COM record descriptor as obj.
func (obj *IRecordInfo) Equals(recordInfo *IRecordInfo) bool {
	return obj.IsMatchingType(recordInfo)
}
