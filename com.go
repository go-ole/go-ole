//go:build windows

package ole

import (
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	procCoInitializeSecurity = modole32.NewProc("CoInitializeSecurity")
	procGetActiveObject      = modoleaut32.NewProc("GetActiveObject")
	procGetUserDefaultLCID   = modkernel32.NewProc("GetUserDefaultLCID")
	procCoCreateInstance     = modole32.NewProc("CoCreateInstance")
	procCoGetObject          = modole32.NewProc("CoGetObject")
	procCreateDispTypeInfo   = modoleaut32.NewProc("CreateDispTypeInfo")
	procCreateStdDispatch    = modoleaut32.NewProc("CreateStdDispatch")
	procCopyMemory           = modkernel32.NewProc("RtlMoveMemory")
)

const (
	CLSCTX_INPROC_SERVER   = 1
	CLSCTX_INPROC_HANDLER  = 2
	CLSCTX_LOCAL_SERVER    = 4
	CLSCTX_INPROC_SERVER16 = 8
	CLSCTX_REMOTE_SERVER   = 16
	CLSCTX_ALL             = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER
	CLSCTX_INPROC          = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER
	CLSCTX_SERVER          = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER
)

const (
	CC_FASTCALL = iota
	CC_CDECL
	CC_MSCPASCAL
	CC_PASCAL = CC_MSCPASCAL
	CC_MACPASCAL
	CC_STDCALL
	CC_FPFASTCALL
	CC_SYSCALL
	CC_MPWCDECL
	CC_MPWPASCAL
	CC_MAX = CC_MPWPASCAL
)

const (
	TKIND_ENUM      = 1
	TKIND_RECORD    = 2
	TKIND_MODULE    = 3
	TKIND_INTERFACE = 4
	TKIND_DISPATCH  = 5
	TKIND_COCLASS   = 6
	TKIND_ALIAS     = 7
	TKIND_UNION     = 8
	TKIND_MAX       = 9
)

// The `ConcurrencyModel` aliases the COINIT_* constants so that the `Initialize()` function is type checked and limited
type ConcurrencyModel uint32

const (
	Multithreaded     ConcurrencyModel = windows.COINIT_MULTITHREADED     // Requires concurrency control
	ApartmentThreaded                  = windows.COINIT_APARTMENTTHREADED // Requires COM operations on the same thread as Initialized.
	DisableOle1DDE                     = windows.COINIT_DISABLE_OLE1DDE
	SpeedOverMemory                    = windows.COINIT_SPEED_OVER_MEMORY
)

// The result from calling `CoInitializeEx()` to prevent exposing syscall windows dependency.
type InitializeResult uint32

const (
	UnknownInitializeResult InitializeResult = iota
	SuccessfullyInitialized
	AlreadyInitialized
	IncompatibleConcurrencyModelAlreadyInitialized
)

// This is to enable calling COM Security initialization multiple times
var bSecurityInit bool = false

// Setup COM for the application.
//
// COM Function: CoInitializeEx
func Initialize(model ConcurrencyModel) (InitializeResult, error) {
	err := windows.CoInitializeEx(0, uint32(model))

	if err == nil {
		return SuccessfullyInitialized, nil
	}

	hr := err.(windows.Errno)

	if hr == 1 {
		return AlreadyInitialized, nil
	}

	if uintptr(hr) == uintptr(windows.RPC_E_CHANGED_MODE) {
		return IncompatibleConcurrencyModelAlreadyInitialized, nil
	}

	return UnknownInitializeResult, err
}

// Initialize COM as multithreaded.
//
// Called when you are handling the multithreading for COM.
func InitializeMultithreaded() (InitializeResult, error) {
	return Initialize(Multithreaded)
}

// Initialize COM as apartment-multithreaded.
//
// This is slower, but does not require concurrency control or protections. This assumes the COM code will always run
// on the same thread. This should not be used in a gothread without additional guarantees to ensure the COM code is
// also run on the same thread.
func InitializeApartmentThreaded() (InitializeResult, error) {
	return Initialize(ApartmentThreaded)
}

// You want to call this as a companion for any `Initialize()` call. This ensures that any remaining messages are
// completed that were pending before the application closed.
//
// COM Function: CoUninitialize
func Uninitialize() {
	windows.CoUninitialize()
}

// TaskMemoryFreePointer frees memory owned by COM by Pointer.
//
// COM Function: CoTaskMemFree
func TaskMemoryFreePointer(address unsafe.Pointer) {
	windows.CoTaskMemFree(address)
}

// TaskMemoryFreeAddress frees memory owned by COM by `uintptr`.
//
// COM Function: CoTaskMemFree
func TaskMemoryFreeAddress(address uintptr) {
	p := unsafe.Pointer(&address)
	windows.CoTaskMemFree(p)
}

// CoInitializeSecurity registers security and sets the default security values
// for the process.
func CoInitializeSecurity(cAuthSvc int32,
	dwAuthnLevel uint32,
	dwImpLevel uint32,
	dwCapabilities uint32) (err error) {
	// Check COM Security initialization has done previously
	if !bSecurityInit {
		// https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-coinitializesecurity
		hr, _, _ := procCoInitializeSecurity.Call(
			uintptr(0),              // Allow *all* VSS writers to communicate back!
			uintptr(cAuthSvc),       // Default COM authentication service
			uintptr(0),              // Default COM authorization service
			uintptr(0),              // Reserved parameter
			uintptr(dwAuthnLevel),   // Strongest COM authentication level
			uintptr(dwImpLevel),     // Minimal impersonation abilities
			uintptr(0),              // Default COM authentication settings
			uintptr(dwCapabilities), // Cloaking
			uintptr(0))              // reserved parameter
		if hr != 0 {
			err = windows.Errno(hr)
		} else {
			// COM Security initialization done make global flag true.
			bSecurityInit = true
		}
	}
	return
}

// CreateInstance of single uninitialized object with GUID.
func CreateInstance[T IsIUnknown](clsid windows.GUID, iid windows.GUID) (unk *T, err error) {
	hr, _, _ := procCoCreateInstance.Call(
		uintptr(unsafe.Pointer(&clsid)),
		0,
		CLSCTX_SERVER,
		uintptr(unsafe.Pointer(&iid)),
		uintptr(unsafe.Pointer(&unk)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// GetActiveObject retrieves virtual table to active object.
//
// [T] must be a virtual table structure. This function is unsafe(!!!) and will attempt to populate whatever type you
// pass.
func GetActiveObject[T struct{}](classId windows.GUID, interfaceId windows.GUID) (obj *T, err error) {
	hr, _, _ := procGetActiveObject.Call(
		uintptr(unsafe.Pointer(&classId)),
		uintptr(unsafe.Pointer(&interfaceId)),
		uintptr(unsafe.Pointer(&obj)))
	if hr != 0 {
		return nil, windows.Errno(hr)
	}
	return
}

// GetUserDefaultLCID retrieves current user default locale.
func GetUserDefaultLCID() (lcid uint32) {
	ret, _, _ := procGetUserDefaultLCID.Call()
	lcid = uint32(ret)
	return
}

// GetObject retrieves pointer to active object.
func GetObject[T IsIUnknown](programID string, bindOpts *windows.BIND_OPTS3, interfaceId windows.GUID) (unk *T, err error) {
	if bindOpts != nil {
		bindOpts.CbStruct = uint32(unsafe.Sizeof(windows.BIND_OPTS3{}))
	}
	hr := windows.CoGetObject(
		windows.StringToUTF16Ptr(programID),
		bindOpts,
		&interfaceId,
		&uintptr(unsafe.Pointer(&unk)))
	if hr == nil {
		return
	}
	err = hr
	return
}

// CreateStdDispatch provides default IDispatch implementation for IUnknown.
//
// This handles default IDispatch implementation for objects. It has a few limitations with only supporting one
// language. It will also only return default exception codes.
func CreateStdDispatch(unk *IUnknown, v uintptr, ptinfo *IUnknown) (disp *IDispatch, err error) {
	hr, _, _ := procCreateStdDispatch.Call(
		uintptr(unsafe.Pointer(unk)),
		v,
		uintptr(unsafe.Pointer(ptinfo)),
		uintptr(unsafe.Pointer(&disp)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// CreateDispTypeInfo provides default ITypeInfo implementation for IDispatch.
//
// This will not handle the full implementation of the interface.
func CreateDispTypeInfo(idata *INTERFACEDATA) (pptinfo *IUnknown, err error) {
	hr, _, _ := procCreateDispTypeInfo.Call(
		uintptr(unsafe.Pointer(idata)),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&pptinfo)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// RtlMoveMemory moves location of a block of memory.
func RtlMoveMemory(dest unsafe.Pointer, src unsafe.Pointer, length uint32) {
	procCopyMemory.Call(uintptr(dest), uintptr(src), uintptr(length))
}
