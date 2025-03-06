//go:build windows

package ole

import (
	"errors"
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	procCLSIDFromProgID = modole32.NewProc("CLSIDFromProgID")
	procCLSIDFromString = modole32.NewProc("CLSIDFromString")
	procStringFromCLSID = modole32.NewProc("StringFromCLSID")
	procStringFromIID   = modole32.NewProc("StringFromIID")
	procIIDFromString   = modole32.NewProc("IIDFromString")
)

var (
	UnknownLookupArgument          = errors.New("ole: unknown lookup argument")
	InvalidLookupArgument          = errors.New("ole: invalid lookup argument")
	InvalidClassIdForProgramId     = errors.New("ole: invalid class id from program id")
	ImproperlyFormattedGUID        = errors.New("ole: invalid formatted guid")
	UnableToWriteClassIdToRegistry = errors.New("ole: unable to write class id from program id")
)

// ClassIdFromProgramId retrieves Class Identifier with the given Program Identifier.
//
// The Programmatic Identifier must be registered, because it will be looked up
// in the Windows Registry. The registry entry has the following keys: CLSID,
// Insertable, Protocol and Shell
// (https://msdn.microsoft.com/en-us/library/dd542719(v=vs.85).aspx).
//
// programID identifies the class id with less precision and is not guaranteed
// to be unique. These are usually found in the registry under
// HKEY_LOCAL_MACHINE\SOFTWARE\Classes, usually with the format of
// "Program.Component.Version" with version being optional.
//
// COM function: CLSIDFromProgID
func ClassIdFromProgramId(programId string) (classId windows.GUID, err error) {
	lookup, err := windows.UTF16PtrFromString(programId)
	if err != nil {
		return
	}

	hr, _, _ := procCLSIDFromProgID.Call(uintptr(unsafe.Pointer(lookup)), uintptr(unsafe.Pointer(&classId)))

	switch windows.Handle(hr) {
	case windows.S_OK:
		return
	case windows.CO_E_CLASSSTRING:
		err = InvalidClassIdForProgramId
	case windows.REGDB_E_WRITEREGDB:
		err = UnableToWriteClassIdToRegistry
	default:
		err = UnknownLookupArgument
	}
	return
}

// ClassIdFromGuidString retrieves Class ID from string representation.
//
// This is technically the string version of the GUID and will convert the string to object.
//
// COM function: CLSIDFromString
func ClassIdFromGuidString(guid string) (classId windows.GUID, err error) {
	lookup, err := windows.UTF16PtrFromString(guid)
	if err != nil {
		return
	}

	hr, _, _ := procCLSIDFromString.Call(uintptr(unsafe.Pointer(lookup)), uintptr(unsafe.Pointer(&classId)))

	switch windows.Handle(hr) {
	case windows.Handle(windows.NOERROR):
		return
	case windows.CO_E_CLASSSTRING:
		err = ImproperlyFormattedGUID
	case windows.E_INVALIDARG:
		err = InvalidLookupArgument
	default:
		err = UnknownLookupArgument
	}
	return
}

// ClassIDFromString retrieves class ID whether given is program ID or application string.
//
// Helper that provides check against both Class ID from Program ID and Class ID from string. It is
// faster, if you know which you are using, to use the individual functions, but this will check
// against available functions for you.
func ClassIdFromString(lookup string) (classId windows.GUID, err error) {
	classId, err = ClassIdFromProgramId(lookup)
	if err != nil {
		classId, err = ClassIdFromGuidString(lookup)
		if err != nil {
			return
		}
	}
	return
}

// InterfaceIdFromString returns GUID from value returned by InterfaceIdToString.
//
// COM function: IIDFromString
func InterfaceIdFromString(interfaceId string) (classId windows.GUID, err error) {
	lookup, err := windows.UTF16PtrFromString(interfaceId)
	if err != nil {
		return
	}
	hr, _, _ := procIIDFromString.Call(uintptr(unsafe.Pointer(lookup)), uintptr(unsafe.Pointer(&classId)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// ClassIdToString returns GUID formated string from GUID object.
//
// COM function: StringFromCLSID
func ClassIdToString(classId windows.GUID) (str string, err error) {
	var p *uint16
	hr, _, _ := procStringFromCLSID.Call(uintptr(unsafe.Pointer(&classId)), uintptr(unsafe.Pointer(&p)))
	if hr != 0 {
		err = windows.Errno(hr)
		return
	}
	str = windows.UTF16PtrToString(p)
	return
}

// InterfaceIdToString returns GUID formatted string from GUID object.
//
// COM function: StringFromIID
func InterfaceIdToString(iid windows.GUID) (str string, err error) {
	var p *uint16
	hr, _, _ := procStringFromIID.Call(uintptr(unsafe.Pointer(&iid)), uintptr(unsafe.Pointer(&p)))
	if hr != 0 {
		err = windows.Errno(hr)
		return
	}
	str = windows.UTF16PtrToString(p)
	return
}
