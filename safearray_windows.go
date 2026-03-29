//go:build windows

package ole

import (
	"unsafe"

	"golang.org/x/sys/windows"
)

func safeArrayCreateResult(ptr uintptr, err error) (uintptr, error) {
	if ptr == 0 && err != nil && err != windows.ERROR_SUCCESS {
		return 0, err
	}
	return ptr, nil
}

// SafeArrayCreate wraps the Windows SafeArrayCreate API.
func SafeArrayCreate(vt VT, dimensions uint32, bounds unsafe.Pointer) (uintptr, error) {
	ptr, _, err := procSafeArrayCreate.Call(uintptr(vt), uintptr(dimensions), uintptr(bounds))
	return safeArrayCreateResult(ptr, err)
}

// SafeArrayCreateEx wraps the Windows SafeArrayCreateEx API.
func SafeArrayCreateEx(vt VT, dimensions uint32, bounds unsafe.Pointer, extra uintptr) (uintptr, error) {
	ptr, _, err := procSafeArrayCreateEx.Call(uintptr(vt), uintptr(dimensions), uintptr(bounds), extra)
	return safeArrayCreateResult(ptr, err)
}

// SafeArrayCreateVector wraps the Windows SafeArrayCreateVector API.
func SafeArrayCreateVector(vt VT, lowerBound int32, length uint32) (uintptr, error) {
	ptr, _, err := procSafeArrayCreateVector.Call(uintptr(vt), uintptr(lowerBound), uintptr(length))
	return safeArrayCreateResult(ptr, err)
}

// SafeArrayCreateVectorEx wraps the Windows SafeArrayCreateVectorEx API.
func SafeArrayCreateVectorEx(vt VT, lowerBound int32, length uint32, extra uintptr) (uintptr, error) {
	ptr, _, err := procSafeArrayCreateVectorEx.Call(uintptr(vt), uintptr(lowerBound), uintptr(length), extra)
	return safeArrayCreateResult(ptr, err)
}

// SafeArrayAccessData wraps the Windows SafeArrayAccessData API.
func SafeArrayAccessData(array unsafe.Pointer, data unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayAccessData.Call(uintptr(array), uintptr(data))
	return hr
}

// SafeArrayAddRef wraps the Windows SafeArrayAddRef API.
func SafeArrayAddRef(array unsafe.Pointer, dataToRelease unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayAddRef.Call(uintptr(array), uintptr(dataToRelease))
	return hr
}

// SafeArrayAllocData wraps the Windows SafeArrayAllocData API.
func SafeArrayAllocData(array unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayAllocData.Call(uintptr(array))
	return hr
}

// SafeArrayAllocDescriptor wraps the Windows SafeArrayAllocDescriptor API.
func SafeArrayAllocDescriptor(dimensions uint32, array unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayAllocDescriptor.Call(uintptr(dimensions), uintptr(array))
	return hr
}

// SafeArrayAllocDescriptorEx wraps the Windows SafeArrayAllocDescriptorEx API.
func SafeArrayAllocDescriptorEx(vt VT, dimensions uint32, array unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayAllocDescriptorEx.Call(uintptr(vt), uintptr(dimensions), uintptr(array))
	return hr
}

// SafeArrayCopy wraps the Windows SafeArrayCopy API.
func SafeArrayCopy(array unsafe.Pointer, copy unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayCopy.Call(uintptr(array), uintptr(copy))
	return hr
}

// SafeArrayCopyData wraps the Windows SafeArrayCopyData API.
func SafeArrayCopyData(source unsafe.Pointer, target unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayCopyData.Call(uintptr(source), uintptr(target))
	return hr
}

// SafeArrayDestroy wraps the Windows SafeArrayDestroy API.
func SafeArrayDestroy(array unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayDestroy.Call(uintptr(array))
	return hr
}

// SafeArrayDestroyData wraps the Windows SafeArrayDestroyData API.
func SafeArrayDestroyData(array unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayDestroyData.Call(uintptr(array))
	return hr
}

// SafeArrayDestroyDescriptor wraps the Windows SafeArrayDestroyDescriptor API.
func SafeArrayDestroyDescriptor(array unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayDestroyDescriptor.Call(uintptr(array))
	return hr
}

// SafeArrayGetDim wraps the Windows SafeArrayGetDim API.
func SafeArrayGetDim(array unsafe.Pointer) uintptr {
	ret, _, _ := procSafeArrayGetDim.Call(uintptr(array))
	return ret
}

// SafeArrayGetElement wraps the Windows SafeArrayGetElement API.
func SafeArrayGetElement(array unsafe.Pointer, indices unsafe.Pointer, element uintptr) uintptr {
	hr, _, _ := procSafeArrayGetElement.Call(uintptr(array), uintptr(indices), element)
	return hr
}

// SafeArrayGetElemsize wraps the Windows SafeArrayGetElemsize API.
func SafeArrayGetElemsize(array unsafe.Pointer) uintptr {
	ret, _, _ := procSafeArrayGetElemsize.Call(uintptr(array))
	return ret
}

// SafeArrayGetIID wraps the Windows SafeArrayGetIID API.
func SafeArrayGetIID(array unsafe.Pointer, guid unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayGetIID.Call(uintptr(array), uintptr(guid))
	return hr
}

// SafeArrayGetLBound wraps the Windows SafeArrayGetLBound API.
func SafeArrayGetLBound(array unsafe.Pointer, dimension uint32, bound unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayGetLBound.Call(uintptr(array), uintptr(dimension), uintptr(bound))
	return hr
}

// SafeArrayGetRecordInfo wraps the Windows SafeArrayGetRecordInfo API.
func SafeArrayGetRecordInfo(array unsafe.Pointer, recordInfo unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayGetRecordInfo.Call(uintptr(array), uintptr(recordInfo))
	return hr
}

// SafeArrayGetUBound wraps the Windows SafeArrayGetUBound API.
func SafeArrayGetUBound(array unsafe.Pointer, dimension uint32, bound unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayGetUBound.Call(uintptr(array), uintptr(dimension), uintptr(bound))
	return hr
}

// SafeArrayGetVartype wraps the Windows SafeArrayGetVartype API.
func SafeArrayGetVartype(array unsafe.Pointer, vt unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayGetVartype.Call(uintptr(array), uintptr(vt))
	return hr
}

// SafeArrayLock wraps the Windows SafeArrayLock API.
func SafeArrayLock(array unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayLock.Call(uintptr(array))
	return hr
}

// SafeArrayPtrOfIndex wraps the Windows SafeArrayPtrOfIndex API.
func SafeArrayPtrOfIndex(array unsafe.Pointer, indices unsafe.Pointer, element unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayPtrOfIndex.Call(uintptr(array), uintptr(indices), uintptr(element))
	return hr
}

// SafeArrayPutElement wraps the Windows SafeArrayPutElement API.
func SafeArrayPutElement(array unsafe.Pointer, indices unsafe.Pointer, element uintptr) uintptr {
	hr, _, _ := procSafeArrayPutElement.Call(uintptr(array), uintptr(indices), element)
	return hr
}

// SafeArrayRedim wraps the Windows SafeArrayRedim API.
func SafeArrayRedim(array unsafe.Pointer, bound unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayRedim.Call(uintptr(array), uintptr(bound))
	return hr
}

// SafeArrayReleaseData wraps the Windows SafeArrayReleaseData API.
func SafeArrayReleaseData(data uintptr) {
	procSafeArrayReleaseData.Call(data)
}

// SafeArrayReleaseDescriptor wraps the Windows SafeArrayReleaseDescriptor API.
func SafeArrayReleaseDescriptor(array unsafe.Pointer) {
	procSafeArrayReleaseDescriptor.Call(uintptr(array))
}

// SafeArraySetIID wraps the Windows SafeArraySetIID API.
func SafeArraySetIID(array unsafe.Pointer, guid unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArraySetIID.Call(uintptr(array), uintptr(guid))
	return hr
}

// SafeArraySetRecordInfo wraps the Windows SafeArraySetRecordInfo API.
func SafeArraySetRecordInfo(array unsafe.Pointer, recordInfo unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArraySetRecordInfo.Call(uintptr(array), uintptr(recordInfo))
	return hr
}

// SafeArrayUnaccessData wraps the Windows SafeArrayUnaccessData API.
func SafeArrayUnaccessData(array unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayUnaccessData.Call(uintptr(array))
	return hr
}

// SafeArrayUnlock wraps the Windows SafeArrayUnlock API.
func SafeArrayUnlock(array unsafe.Pointer) uintptr {
	hr, _, _ := procSafeArrayUnlock.Call(uintptr(array))
	return hr
}
