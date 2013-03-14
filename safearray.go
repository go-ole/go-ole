package ole

import (
	_ "fmt"
	_ "syscall"
	"unsafe"
)

var (
	procSafeArrayAccessData, _        = modoleaut32.FindProc("SafeArrayAccessData")
	procSafeArrayAllocData, _         = modoleaut32.FindProc("SafeArrayAllocData")
	procSafeArrayAllocDescriptor, _   = modoleaut32.FindProc("SafeArrayAllocDescriptor")
	procSafeArrayAllocDescriptorEx, _ = modoleaut32.FindProc("SafeArrayAllocDescriptorEx")
	procSafeArrayCopy, _              = modoleaut32.FindProc("SafeArrayCopy")
	procSafeArrayCopyData, _          = modoleaut32.FindProc("SafeArrayCopyData")
	procSafeArrayCreate, _            = modoleaut32.FindProc("SafeArrayCreate")
	procSafeArrayCreateEx, _          = modoleaut32.FindProc("SafeArrayCreateEx")
	procSafeArrayCreateVector, _      = modoleaut32.FindProc("SafeArrayCreateVector")
	procSafeArrayCreateVectorEx, _    = modoleaut32.FindProc("SafeArrayCreateVectorEx")
	procSafeArrayDestroy, _           = modoleaut32.FindProc("SafeArrayDestroy")
	procSafeArrayDestroyData, _       = modoleaut32.FindProc("SafeArrayDestroyData")
	procSafeArrayDestroyDescriptor, _ = modoleaut32.FindProc("SafeArrayDestroyDescriptor")
	procSafeArrayGetDim, _            = modoleaut32.FindProc("SafeArrayGetDim")
	procSafeArrayGetElement, _        = modoleaut32.FindProc("SafeArrayGetElement")
	procSafeArrayGetElemsize, _       = modoleaut32.FindProc("SafeArrayGetElemsize")
	procSafeArrayGetIID, _            = modoleaut32.FindProc("SafeArrayGetIID")
	procSafeArrayGetLBound, _         = modoleaut32.FindProc("SafeArrayGetLBound")
	procSafeArrayGetRecordInfo, _     = modoleaut32.FindProc("SafeArrayGetRecordInfo")
	procSafeArrayGetUBound, _         = modoleaut32.FindProc("SafeArrayGetUBound")
	procSafeArrayGetVartype, _        = modoleaut32.FindProc("SafeArrayGetVartype")
	procSafeArrayLock, _              = modoleaut32.FindProc("SafeArrayLock")
	procSafeArrayPtrOfIndex, _        = modoleaut32.FindProc("SafeArrayPtrOfIndex")
	procSafeArrayPutElement, _        = modoleaut32.FindProc("SafeArrayPutElement")
	procSafeArrayRedim, _             = modoleaut32.FindProc("SafeArrayRedim")
	procSafeArraySetIID, _            = modoleaut32.FindProc("SafeArraySetIID")
	procSafeArraySetRecordInfo, _     = modoleaut32.FindProc("SafeArraySetRecordInfo")
	procSafeArrayUnaccessData, _      = modoleaut32.FindProc("SafeArrayUnaccessData")
	procSafeArrayUnlock, _            = modoleaut32.FindProc("SafeArrayUnlock")
)

// Returns Raw Array
// Todo: Test
func safeArrayAccessData(safearray *SAFEARRAY) (elem uintptr, err error) {
	err = convertHresultToError(
		procSafeArrayAccessData.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&elem))))
	return
}

func safeArrayAllocData(safearray *SAFEARRAY) (err error) {
	err = convertHresultToError(procSafeArrayAllocData.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

func safeArrayAllocDescriptor(dimensions uint32) (safearray *SAFEARRAY, err error) {
	err = convertHresultToError(
		procSafeArrayAllocDescriptor.Call(uintptr(dimensions), uintptr(unsafe.Pointer(&safearray))))
	return
}

func safeArrayAllocDescriptorEx(variantType uint16, dimensions uint32) (safearray *SAFEARRAY, err error) {
	err = convertHresultToError(
		procSafeArrayAllocDescriptorEx.Call(
			uintptr(variantType),
			uintptr(dimensions),
			uintptr(unsafe.Pointer(&safearray))))
	return
}

func safeArrayCopy(original *SAFEARRAY) (safearray *SAFEARRAY, err error) {
	err = convertHresultToError(
		procSafeArrayCopy.Call(
			uintptr(unsafe.Pointer(original)),
			uintptr(unsafe.Pointer(&safearray))))
	return
}

func safeArrayCopyData(original *SAFEARRAY, duplicate *SAFEARRAY) (err error) {
	err = convertHresultToError(
		procSafeArrayCopyData.Call(
			uintptr(unsafe.Pointer(original)),
			uintptr(unsafe.Pointer(duplicate))))
	return
}

func safeArrayCreate(variantType uint16, dimensions uint32, bounds *SAFEARRAYBOUND) (safearray *SAFEARRAY, err error) {
	sa, _, err = procSafeArrayCreate.Call(
			uintptr(variantType),
			uintptr(dimensions),
			uintptr(unsafe.Pointer(bounds))))
	safearray = (*SAFEARRAY)(unsafe.Pointer(sa))
	return
}

func safeArrayCreateEx(variantType uint16, dimensions uint32, bounds *SAFEARRAYBOUND, extra uintptr) (safearray *SAFEARRAY, err error) {
	sa, _, err = procSafeArrayCreateEx.Call(
			uintptr(variantType),
			uintptr(dimensions),
			uintptr(unsafe.Pointer(bounds)),
			extra))
	safearray = (*SAFEARRAY)(unsafe.Pointer(sa))
	return
}

func safeArrayCreateVector(variantType uint16, lowerBound int32, length uint32) (safearray *SAFEARRAY, err error) {
	sa, _, err = procSafeArrayCreateVector.Call(
		uintptr(variantType),
		uintptr(lowerBound),
		uintptr(length)))
	safearray = (*SAFEARRAY)(unsafe.Pointer(sa))
	return
}

func safeArrayCreateVectorEx(variantType uint16, lowerBound int32, length uint32, extra uintptr) (safearray *SAFEARRAY, err error) {
	sa, _, err = procSafeArrayCreateVectorEx.Call(
		uintptr(variantType),
		uintptr(lowerBound),
		uintptr(length),
		extra)
	safearray = (*SAFEARRAY)(unsafe.Pointer(sa))
	return
}

func safeArrayDestroy(safearray *SAFEARRAY) (err error) {
	err = convertHresultToError(procSafeArrayDestroy.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

func safeArrayDestroyData(safearray *SAFEARRAY) (err error) {
	err = convertHresultToError(procSafeArrayDestroyData.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

func safeArrayDestroyDescriptor(safearray *SAFEARRAY) (err error) {
	err = convertHresultToError(procSafeArrayDestroyDescriptor.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

func safeArrayGetDim(safearray *SAFEARRAY) (dimensions *uint32, err error) {
	l, _, err = procSafeArrayGetDim.Call(uintptr(unsafe.Pointer(safearray)))
	dimensions = (*uint32)(unsafe.Pointer(l))
	return
}

func safeArrayGetElementSize(safearray *SAFEARRAY) (length *uint32, err error) {
	l, _, err = procSafeArrayGetElemsize.Call(uintptr(unsafe.Pointer(safearray)))
	length = (*uint32)(unsafe.Pointer(l))
	return
}

// This is probably wrong.
func safeArrayGetElement(safearray *SAFEARRAY, index int32) (variant *VARIANT, err error) {
	err = convertHresultToError(
		procSafeArrayGetElement.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(index),
			uintptr(unsafe.Pointer(&variant))))
	return
}

func safeArrayGetIID(safearray *SAFEARRAY) (guid *GUID, err error) {
	err = convertHresultToError(
		procSafeArrayGetIID.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&guid))))
	return
}

type SAFEARRAYBOUND struct {
	CElements uint32
	LLbound   int32
}

type SAFEARRAY struct {
	CDims      uint16
	FFeatures  uint16
	CbElements uint32
	CLocks     uint32
	PvData     uint32
	RgsaBound  SAFEARRAYBOUND
}
