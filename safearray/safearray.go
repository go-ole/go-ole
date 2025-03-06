//go:build windows

// Package is meant to retrieve and process safe array data returned from COM.

package safearray

import (
	"errors"
	"github.com/go-ole/go-ole"

	"golang.org/x/sys/windows"
)

// Safe Array Feature Flags

const (
	FADF_AUTO        = 0x0001
	FADF_STATIC      = 0x0002
	FADF_EMBEDDED    = 0x0004
	FADF_FIXEDSIZE   = 0x0010
	FADF_RECORD      = 0x0020
	FADF_HAVEIID     = 0x0040
	FADF_HAVEVARTYPE = 0x0080
	FADF_BSTR        = 0x0100
	FADF_UNKNOWN     = 0x0200
	FADF_DISPATCH    = 0x0400
	FADF_VARIANT     = 0x0800
	FADF_RESERVED    = 0xF008
)

// SafeArrayBound defines the SafeArray boundaries.
type SafeArrayBound struct {
	Elements   uint32
	LowerBound int32
}

// SafeArray is how COM handles arrays.
type SafeArray[T any] struct {
	Dimensions   uint16
	FeaturesFlag uint16
	ElementsSize uint32
	LocksAmount  uint32
	Data         uintptr
	Bounds       [16]byte
}

var (
	procSafeArrayAccessData        = modoleaut32.NewProc("SafeArrayAccessData")
	procSafeArrayAllocData         = modoleaut32.NewProc("SafeArrayAllocData")
	procSafeArrayAllocDescriptor   = modoleaut32.NewProc("SafeArrayAllocDescriptor")
	procSafeArrayAllocDescriptorEx = modoleaut32.NewProc("SafeArrayAllocDescriptorEx")
	procSafeArrayCopy              = modoleaut32.NewProc("SafeArrayCopy")
	procSafeArrayCopyData          = modoleaut32.NewProc("SafeArrayCopyData")
	procSafeArrayCreate            = modoleaut32.NewProc("SafeArrayCreate")
	procSafeArrayCreateEx          = modoleaut32.NewProc("SafeArrayCreateEx")
	procSafeArrayCreateVector      = modoleaut32.NewProc("SafeArrayCreateVector")
	procSafeArrayCreateVectorEx    = modoleaut32.NewProc("SafeArrayCreateVectorEx")
	procSafeArrayDestroy           = modoleaut32.NewProc("SafeArrayDestroy")
	procSafeArrayDestroyData       = modoleaut32.NewProc("SafeArrayDestroyData")
	procSafeArrayDestroyDescriptor = modoleaut32.NewProc("SafeArrayDestroyDescriptor")
	procSafeArrayGetDim            = modoleaut32.NewProc("SafeArrayGetDim")
	procSafeArrayGetElement        = modoleaut32.NewProc("SafeArrayGetElement")
	procSafeArrayGetElemsize       = modoleaut32.NewProc("SafeArrayGetElemsize")
	procSafeArrayGetIID            = modoleaut32.NewProc("SafeArrayGetIID")
	procSafeArrayGetLBound         = modoleaut32.NewProc("SafeArrayGetLBound")
	procSafeArrayGetUBound         = modoleaut32.NewProc("SafeArrayGetUBound")
	procSafeArrayGetVartype        = modoleaut32.NewProc("SafeArrayGetVartype")
	procSafeArrayLock              = modoleaut32.NewProc("SafeArrayLock")
	procSafeArrayPtrOfIndex        = modoleaut32.NewProc("SafeArrayPtrOfIndex")
	procSafeArrayUnaccessData      = modoleaut32.NewProc("SafeArrayUnaccessData")
	procSafeArrayUnlock            = modoleaut32.NewProc("SafeArrayUnlock")
	procSafeArrayPutElement        = modoleaut32.NewProc("SafeArrayPutElement")
	procSafeArrayRedim             = modoleaut32.NewProc("SafeArrayRedim")
	procSafeArraySetIID            = modoleaut32.NewProc("SafeArraySetIID")
	procSafeArrayGetRecordInfo     = modoleaut32.NewProc("SafeArrayGetRecordInfo")
	procSafeArraySetRecordInfo     = modoleaut32.NewProc("SafeArraySetRecordInfo")
)

var (
	ArgumentNotSafeArrayError            = errors.New("invalid argument")
	UnableToLockSafeArrayError           = errors.New("safe array could not be locked")
	UnableToUnlockSafeArrayError         = errors.New("safe array could not be unlocked")
	SafeArrayIsLockedError               = errors.New("safe array is locked")
	OutOfMemorySafeArrayError            = errors.New("out of memory")
	BadIndexSafeArrayError               = errors.New("bad index")
	MissingInterfaceIdFlagSafeArrayError = errors.New("missing interface id (FADF_HAVEIID) flag")
	MissingRecordFlagSafeArrayError      = errors.New("missing record (FADF_RECORD) flag")
)

// MarshalSafeArray converts Go array slice to SafeArray for COM/OLE transport.
func MarshalSafeArray[T any]() *SafeArray[T] {
	return &SafeArray[T]{}
}

// UnmarshalSafeArray converts SafeArray to Go array slice from COM/OLE transport.
func UnmarshalSafeArray[T any](sa *SafeArray[T]) T {
	return (T)(uintptr(unsafe.Pointer(sa)))
}

func convertHresultToError(hr uintptr, _ uintptr, _ error) (err error) {
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// AccessData returns converted to Go array slice.
//
// You must call Release() after you are finished.
//
// AKA: SafeArrayAccessData in Windows API.
func (sa *SafeArray[T]) AccessData() (element *[]T, err error) {
	var ptr *uintptr
	hr, _, _ := procSafeArrayAccessData.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(unsafe.Pointer(&ptr)))

	if hr == 0 {
		element = (*[sa.ElementsSize]T)(unsafe.Pointer(ptr))[:]
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.E_UNEXPECTED:
		err = UnableToLockSafeArrayError
	}

	return
}

// UnaccessData releases the memory locked by AccessData() raw array.
//
// AKA: SafeArrayUnaccessData in Windows API.
func (sa *SafeArray[T]) UnaccessData() (err error) {
	hr, _, _ := procSafeArrayUnaccessData.Call(uintptr(unsafe.Pointer(sa)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.E_UNEXPECTED:
		err = UnableToUnlockSafeArrayError
	}

	return
}

// AllocData allocates SafeArray.
//
// AKA: SafeArrayAllocData in Windows API.
func (sa *SafeArray[T]) AllocData() (err error) {
	hr, _, _ := procSafeArrayAllocData.Call(uintptr(unsafe.Pointer(sa)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.E_UNEXPECTED:
		err = UnableToLockSafeArrayError
	}
	return
}

// Clone returns copy of SafeArray.
//
// AKA: SafeArrayCopy in Windows API.
func (sa *SafeArray[T]) Clone() (safearray *SafeArray, err error) {
	hr, _, _ := procSafeArrayCopy.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(unsafe.Pointer(&safearray)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.E_OUTOFMEMORY:
		err = OutOfMemorySafeArrayError
	}
	return
}

// Copy duplicates SafeArray into another SafeArray object.
//
// AKA: SafeArrayCopyData in Windows API.
func (sa *SafeArray[T]) Copy(duplicate *SafeArray) (err error) {
	hr, _, _ := procSafeArrayCopyData.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(unsafe.Pointer(duplicate)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.E_OUTOFMEMORY:
		err = OutOfMemorySafeArrayError
	}
	return
}

// GetDimensions is the amount of dimensions in the SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be multidimensional array.
//
// AKA: SafeArrayGetDim in Windows API.
func (sa *SafeArray[T]) GetDimensions() (dimensions *uint32, err error) {
	l, _, err := procSafeArrayGetDim.Call(uintptr(unsafe.Pointer(sa)))
	dimensions = (*uint32)(unsafe.Pointer(l))
	return
}

// GetElementSize is the element size in bytes.
//
// AKA: SafeArrayGetElemsize in Windows API.
func (sa *SafeArray[T]) GetElementSize() (length *uint32, err error) {
	l, _, err := procSafeArrayGetElemsize.Call(uintptr(unsafe.Pointer(sa)))
	length = (*uint32)(unsafe.Pointer(l))
	return
}

// ReDim changes the right-most (least significant) bound.
//
// AKA: SafeArrayRedim in Windows API.
func (sa *SafeArray[T]) ReDim(bounds *SafeArrayBound) (err error) {
	hr, _, _ := procSafeArrayRedim.Call(uintptr(unsafe.Pointer(&sa)), uintptr(unsafe.Pointer(bounds)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.DISP_E_ARRAYISLOCKED:
		err = SafeArrayIsLockedError
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	}
	return
}

// safeArrayGetElement retrieves element at given index.
func (sa *SafeArray[T]) GetElement(index int32) (element T, err error) {
	var hr uintptr
	switch any(element).(type) {
	case string:
		var ptr *uint16
		hr, _, _ = procSafeArrayGetElement.Call(
			uintptr(unsafe.Pointer(sa)),
			uintptr(unsafe.Pointer(&index)),
			uintptr(unsafe.Pointer(&ptr)))
		if hr == 0 {
			element = windows.UTF16PtrToString(ptr)
			return
		}
	default:
		hr, _, _ = procSafeArrayGetElement.Call(
			uintptr(unsafe.Pointer(sa)),
			uintptr(unsafe.Pointer(&index)),
			uintptr(unsafe.Pointer(&element)))
		if hr == 0 {
			return
		}
	}

	switch windows.Handle(hr) {
	case windows.DISP_E_BADINDEX:
		err = BadIndexSafeArrayError
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.E_OUTOFMEMORY:
		err = OutOfMemorySafeArrayError
	}
	return
}

// GetInterfaceId is the InterfaceID (IID) of the elements in the SafeArray.
//
// AKA: SafeArrayGetIID in Windows API.
func (sa *SafeArray[T]) GetInterfaceId() (guid windows.GUID, err error) {
	hr, _, _ := procSafeArrayGetIID.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(&guid)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = MissingInterfaceIdFlagSafeArrayError
	}

	return
}

// SetInterfaceId is the InterfaceID (IID) of the elements in the SafeArray.
//
// AKA: SafeArrayGetIID in Windows API.
func (sa *SafeArray[T]) SetInterfaceId(interfaceId windows.GUID) (err error) {
	hr, _, _ := procSafeArraySetIID.Call(
		uintptr(unsafe.Pointer(safearray)),
		uintptr(unsafe.Pointer(&guid)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = MissingInterfaceIdFlagSafeArrayError
	}

	return
}

// GetVarType returns data type of SafeArray.
//
// AKA: SafeArrayGetVartype in Windows API.
func (sa *SafeArray[T]) GetVarType() (vt ole.VT, err error) {
	var varType uint16
	hr, _, _ := procSafeArrayGetVartype.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(unsafe.Pointer(&varType)))

	if hr == 0 {
		vt = ole.VT(varType)
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	}
	return
}

// Lock SafeArray for reading to modify SafeArray.
//
// This must be called during some calls to ensure that another process does not
// read or write to the SafeArray during editing.
//
// AKA: SafeArrayLock in Windows API.
func (sa *SafeArray[T]) Lock() (err error) {
	hr, _, _ := procSafeArrayLock.Call(uintptr(unsafe.Pointer(sa)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.E_UNEXPECTED:
		err = UnableToLockSafeArrayError
	}
	return
}

// Unlock SafeArray for reading.
//
// AKA: SafeArrayUnlock in Windows API.
func (sa *SafeArray[T]) Unlock() (err error) {
	hr, _, _ := procSafeArrayUnlock.Call(uintptr(unsafe.Pointer(sa)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.E_UNEXPECTED:
		err = UnableToUnlockSafeArrayError
	}
	return
}

// GetRecordInfo accesses IRecordInfo info for custom types.
//
// AKA: SafeArrayGetRecordInfo in Windows API.
func (sa *SafeArray[T]) GetRecordInfo() (recordInfo ole.IRecordInfo, err error) {
	hr, _, _ := procSafeArrayGetRecordInfo.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(unsafe.Pointer(&recordInfo)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = MissingRecordFlagSafeArrayError
	}
	return
}

// SetRecordInfo mutates IRecordInfo info for custom types.
//
// AKA: SafeArraySetRecordInfo in Windows API.
func (sa *SafeArray[T]) SetRecordInfo(recordInfo ole.IRecordInfo) (err error) {
	hr, _, _ := procSafeArraySetRecordInfo.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(unsafe.Pointer(&recordInfo)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = MissingRecordFlagSafeArrayError
	}
	return
}

// Destroy SafeArray object.
//
// See remarks in https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraydestroy
//
// AKA: SafeArrayDestroy in Windows API.
func (sa *SafeArray[T]) Destroy() (err error) {
	hr, _, _ := procSafeArrayDestroy.Call(uintptr(unsafe.Pointer(sa)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.DISP_E_ARRAYISLOCKED:
		err = SafeArrayIsLockedError
	}
	return
}

// DestroyData destroys SafeArray object.
//
// See remarks in https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraydestroydata
//
// AKA: SafeArrayDestroyData in Windows API.
func (sa *SafeArray[T]) DestroyData() (err error) {
	hr, _, _ := procSafeArrayDestroyData.Call(uintptr(unsafe.Pointer(sa)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.DISP_E_ARRAYISLOCKED:
		err = SafeArrayIsLockedError
	}
	return
}

// DestroyDescriptor destroys SafeArray object.
//
// DestroyData() should be called before this function pointer.
//
// See remarks in https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraydestroydescriptor
//
// AKA: SafeArrayDestroyDescriptor in Windows API.
func (sa *SafeArray[T]) DestroyDescriptor() (err error) {
	hr, _, _ := procSafeArrayDestroyDescriptor.Call(uintptr(unsafe.Pointer(sa)))
	if hr == 0 {
		return
	}

	switch windows.Handle(hr) {
	case windows.E_INVALIDARG:
		err = ArgumentNotSafeArrayError
	case windows.DISP_E_ARRAYISLOCKED:
		err = SafeArrayIsLockedError
	}
	return
}

// safeArrayAllocDescriptor allocates SafeArray.
//
// AKA: SafeArrayAllocDescriptor in Windows API.
func safeArrayAllocDescriptor(dimensions uint32) (safearray *SafeArray, err error) {
	err = convertHresultToError(
		procSafeArrayAllocDescriptor.Call(uintptr(dimensions), uintptr(unsafe.Pointer(&safearray))))
	return
}

// safeArrayAllocDescriptorEx allocates SafeArray.
//
// AKA: SafeArrayAllocDescriptorEx in Windows API.
func safeArrayAllocDescriptorEx(variantType ole.VT, dimensions uint32) (safearray *SafeArray, err error) {
	err = convertHresultToError(
		procSafeArrayAllocDescriptorEx.Call(
			uintptr(variantType),
			uintptr(dimensions),
			uintptr(unsafe.Pointer(&safearray))))
	return
}

// safeArrayCreate creates SafeArray.
//
// AKA: SafeArrayCreate in Windows API.
func safeArrayCreate(variantType ole.VT, dimensions uint32, bounds *SafeArrayBound) (safearray *SafeArray, err error) {
	sa, _, err := procSafeArrayCreate.Call(
		uintptr(variantType),
		uintptr(dimensions),
		uintptr(unsafe.Pointer(bounds)))
	safearray = (*SafeArray)(unsafe.Pointer(&sa))
	return
}

// safeArrayCreateEx creates SafeArray.
//
// AKA: SafeArrayCreateEx in Windows API.
func safeArrayCreateEx(variantType ole.VT, dimensions uint32, bounds *SafeArrayBound, extra uintptr) (safearray *SafeArray, err error) {
	sa, _, err := procSafeArrayCreateEx.Call(
		uintptr(variantType),
		uintptr(dimensions),
		uintptr(unsafe.Pointer(bounds)),
		extra)
	safearray = (*SafeArray)(unsafe.Pointer(sa))
	return
}

// safeArrayCreateVector creates SafeArray.
//
// AKA: SafeArrayCreateVector in Windows API.
func safeArrayCreateVector(variantType ole.VT, lowerBound int32, length uint32) (safearray *SafeArray, err error) {
	sa, _, err := procSafeArrayCreateVector.Call(
		uintptr(variantType),
		uintptr(lowerBound),
		uintptr(length))
	safearray = (*SafeArray)(unsafe.Pointer(sa))
	return
}

// safeArrayCreateVectorEx creates SafeArray.
//
// AKA: SafeArrayCreateVectorEx in Windows API.
func safeArrayCreateVectorEx(variantType ole.VT, lowerBound int32, length uint32, extra uintptr) (safearray *SafeArray, err error) {
	sa, _, err := procSafeArrayCreateVectorEx.Call(
		uintptr(variantType),
		uintptr(lowerBound),
		uintptr(length),
		extra)
	safearray = (*SafeArray)(unsafe.Pointer(sa))
	return
}

// safeArrayGetLBound returns lower bounds of SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetLBound in Windows API.
func safeArrayGetLBound(safearray *SafeArray, dimension uint32) (lowerBound int32, err error) {
	err = convertHresultToError(
		procSafeArrayGetLBound.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(dimension),
			uintptr(unsafe.Pointer(&lowerBound))))
	return
}

// safeArrayGetUBound returns upper bounds of SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetUBound in Windows API.
func safeArrayGetUBound(safearray *SafeArray, dimension uint32) (upperBound int32, err error) {
	err = convertHresultToError(
		procSafeArrayGetUBound.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(dimension),
			uintptr(unsafe.Pointer(&upperBound))))
	return
}

// safeArrayPutElement stores the data element at the specified location in the
// array.
//
// AKA: SafeArrayPutElement in Windows API.
func safeArrayPutElement(safearray *SafeArray, index int64, element uintptr) (err error) {
	err = convertHresultToError(
		procSafeArrayPutElement.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&index)),
			uintptr(unsafe.Pointer(element))))
	return
}
