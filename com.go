//go:build windows
// +build windows

package ole

import (
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	procCoInitializeSecurity = modole32.NewProc("CoInitializeSecurity")
	procGetActiveObject      = modoleaut32.NewProc("GetActiveObject")
	procGetUserDefaultLCID   = modkernel32.NewProc("GetUserDefaultLCID")
)

// The `ConcurrencyModel` aliases the COINIT_* constants so that the `Initialize()` function is type checked and limited
type ConcurrencyModel uint32

const (
	// Requires concurrency control
	Multithreaded ConcurrencyModel = windows.COINIT_MULTITHREADED

	// Requires COM operations on the same thread as Initialized.
	ApartmentThreaded = windows.COINIT_APARTMENTTHREADED

	DisableOle1DDE = windows.COINIT_DISABLE_OLE1DDE

	SpeedOverMemory = windows.COINIT_SPEED_OVER_MEMORY
)

// The result from calling `CoInitializeEx()` to prevent exposing syscall windows dependency.
type InitializeResult uint32

const (
	SuccessfullyInitialized InitializeResult = iota << 1
	AlreadyInitialized
	IncompatibleConcurrencyModelAlreadyInitialized
)

// Setup COM for the application.
//
// COM Function: CoInitializeEx
func Initialize(model ConcurrencyModel) (InitializeResult, error) {
	err = windows.CoInitializeEx(0, uint32(model))

	switch err {
	case windows.S_OK:
		return SuccessfullyInitialized, nil
	case windows.S_FALSE:
		return AlreadyInitialized, nil
	case windows.RPC_E_CHANGED_MODE:
		return IncompatibleConcurrencyModelAlreadyInitialized, nil
	default:
		return nil, err
	}
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
func Uninitialize() error {
	windows.CoUninitialize()
}

// Free memory owned by COM by Pointer
//
// COM Function: CoTaskMemFree
func TaskMemoryFreePointer(address unsafe.Pointer) {
	windows.CoTaskMemFree(address)
}

// Free memory owned by COM by `uintptr`
//
// COM Function: CoTaskMemFree
func TaskMemoryFreeAddress(address uintptr) {
	p := unsafe.Pointer(&address)
	windows.CoTaskMemFree(p)
}

// coInitializeSecurity: Registers security and sets the default security values
// for the process.
func coInitializeSecurity(cAuthSvc int32,
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
			err = NewError(hr)
		} else {
			// COM Security initialization done make global flag true.
			bSecurityInit = true
		}
	}
	return
}

// CoInitializeSecurity: Registers security and sets the default security values
// for the process.
func CoInitializeSecurity(cAuthSvc int32,
	dwAuthnLevel uint32,
	dwImpLevel uint32,
	dwCapabilities uint32) (err error) {
	return coInitializeSecurity(cAuthSvc, dwAuthnLevel, dwImpLevel, dwCapabilities)
}

// GetActiveObject retrieves virtual table to active object.
//
// [T] must be a virtual table structure. This function is unsafe(!!!) and will attempt to populate whatever type you
// pass.
func GetActiveObject[T struct{}](classId *windows.GUID, interfaceId *windows.GUID) (obj *T, err error) {
	if interfaceId == nil {
		interfaceId = IID_IUnknown
	}
	hr, _, _ := procGetActiveObject.Call(
		uintptr(unsafe.Pointer(classId)),
		uintptr(unsafe.Pointer(interfaceId)),
		uintptr(unsafe.Pointer(&obj)))
	if hr != windows.S_OK {
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
