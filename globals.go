//go:build windows

package ole

import (
	"errors"
	"golang.org/x/sys/windows"
)

var (
	modcombase  = windows.NewLazySystemDLL("combase.dll")
	modkernel32 = windows.NewLazySystemDLL("kernel32.dll")
	modole32    = windows.NewLazySystemDLL("ole32.dll")
	modoleaut32 = windows.NewLazySystemDLL("oleaut32.dll")
	moduser32   = windows.NewLazySystemDLL("user32.dll")
)

var (
	procSafeArrayAccessData        = modoleaut32.NewProc("SafeArrayAccessData")
	procSafeArrayAddRef            = modoleaut32.NewProc("SafeArrayAddRef")
	procSafeArrayAllocData         = modoleaut32.NewProc("SafeArrayAllocData")
	procSafeArrayAllocDescriptor   = modoleaut32.NewProc("SafeArrayAllocDescriptor")
	procSafeArrayAllocDescriptorEx = modoleaut32.NewProc("SafeArrayAllocDescriptorEx")
	procSafeArrayCreate            = modoleaut32.NewProc("SafeArrayCreate")
	procSafeArrayCreateEx          = modoleaut32.NewProc("SafeArrayCreateEx")
	procSafeArrayCreateVector      = modoleaut32.NewProc("SafeArrayCreateVector")
	procSafeArrayCreateVectorEx    = modoleaut32.NewProc("SafeArrayCreateVectorEx")
	procSafeArrayCopy              = modoleaut32.NewProc("SafeArrayCopy")
	procSafeArrayCopyData          = modoleaut32.NewProc("SafeArrayCopyData")
	procSafeArrayDestroy           = modoleaut32.NewProc("SafeArrayDestroy")
	procSafeArrayDestroyData       = modoleaut32.NewProc("SafeArrayDestroyData")
	procSafeArrayDestroyDescriptor = modoleaut32.NewProc("SafeArrayDestroyDescriptor")
	procSafeArrayGetDim            = modoleaut32.NewProc("SafeArrayGetDim")
	procSafeArrayGetElement        = modoleaut32.NewProc("SafeArrayGetElement")
	procSafeArrayGetElemsize       = modoleaut32.NewProc("SafeArrayGetElemsize")
	procSafeArrayGetIID            = modoleaut32.NewProc("SafeArrayGetIID")
	procSafeArrayGetLBound         = modoleaut32.NewProc("SafeArrayGetLBound")
	procSafeArrayGetRecordInfo     = modoleaut32.NewProc("SafeArrayGetRecordInfo")
	procSafeArrayGetUBound         = modoleaut32.NewProc("SafeArrayGetUBound")
	procSafeArrayGetVartype        = modoleaut32.NewProc("SafeArrayGetVartype")
	procSafeArrayLock              = modoleaut32.NewProc("SafeArrayLock")
	procSafeArrayPtrOfIndex        = modoleaut32.NewProc("SafeArrayPtrOfIndex")
	procSafeArrayPutElement        = modoleaut32.NewProc("SafeArrayPutElement")
	procSafeArrayRedim             = modoleaut32.NewProc("SafeArrayRedim")
	procSafeArrayReleaseData       = modoleaut32.NewProc("SafeArrayReleaseData")
	procSafeArrayReleaseDescriptor = modoleaut32.NewProc("SafeArrayReleaseDescriptor")
	procSafeArraySetIID            = modoleaut32.NewProc("SafeArraySetIID")
	procSafeArraySetRecordInfo     = modoleaut32.NewProc("SafeArraySetRecordInfo")
	procSafeArrayUnaccessData      = modoleaut32.NewProc("SafeArrayUnaccessData")
	procSafeArrayUnlock            = modoleaut32.NewProc("SafeArrayUnlock")
)

var (
	MethodNotImplementedError = errors.New("functionality has not been implemented")
)

// Point is 2D vector type.
type Point struct {
	X int32
	Y int32
}
