//go:build windows

package ole

import (
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	procRoInitialize           = modcombase.NewProc("RoInitialize")
	procRoUninitialize         = modcombase.NewProc("RoUninitialize")
	procRoActivateInstance     = modcombase.NewProc("RoActivateInstance")
	procRoGetActivationFactory = modcombase.NewProc("RoGetActivationFactory")
)

// RoThreading selects the apartment model used by RoInitialize.
type RoThreading uint32

const (
	RoSingleThreaded RoThreading = 0
	RoMultithreaded              = 1
)

// RoInitialize initializes the calling thread for WinRT activation APIs.
//
// Example:
//
//	result, err := ole.RoInitialize(ole.RoSingleThreaded)
//	if err != nil {
//		return err
//	}
//	defer ole.RoUninitialize()
//	_ = result
func RoInitialize(threadType RoThreading) (ret InitializeResult, err error) {
	hr, _, _ := procRoInitialize.Call(uintptr(threadType))

	switch windows.Handle(hr) {
	case windows.S_OK:
		return SuccessfullyInitialized, nil
	case windows.S_FALSE:
		return AlreadyInitialized, nil
	case windows.RPC_E_CHANGED_MODE:
		return IncompatibleConcurrencyModelAlreadyInitialized, nil
	case windows.E_INVALIDARG, windows.E_OUTOFMEMORY, windows.E_UNEXPECTED:
		return UnknownInitializeResult, windows.Errno(hr)
	default:
		return UnknownInitializeResult, windows.Errno(hr)
	}
}

// RoUninitialize balances a successful RoInitialize call.
func RoUninitialize() {
	procRoUninitialize.Call()
}

// RoActivateInstance activates the specified Windows Runtime class.
//
// Please note that the IInspectable may be nil but may also return an IInspectable object. You must check for error.
// If you get windows.E_NOINTERFACE, then the IInspectable interface is not implemented by the specified class.
func RoActivateInstance(classId string) (obj *IInspectable, err error) {
	hClassId, err := NewHString(classId)
	if err != nil {
		return nil, err
	}
	defer DeleteHString(hClassId)

	hr, _, _ := procRoActivateInstance.Call(
		uintptr(unsafe.Pointer(hClassId)),
		uintptr(unsafe.Pointer(&obj)))

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	default:
		err = windows.Errno(hr)
	}

	return
}

// RoGetActivationFactory resolves the activation factory for a WinRT runtime class.
//
// Example:
//
//	factory, err := ole.RoGetActivationFactory("Windows.Foundation.Uri", ole.IID_IActivationFactory)
//	if err != nil {
//		return err
//	}
//	defer factory.Release()
func RoGetActivationFactory(classId string, interfaceId windows.GUID) (obj *IActivationFactory, err error) {
	hClassId, err := NewHString(classId)
	if err != nil {
		return nil, err
	}
	defer DeleteHString(hClassId)

	hr, _, _ := procRoGetActivationFactory.Call(
		uintptr(unsafe.Pointer(hClassId)),
		uintptr(unsafe.Pointer(&interfaceId)),
		uintptr(unsafe.Pointer(&obj)))

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	default:
		err = windows.Errno(hr)
	}

	return
}
