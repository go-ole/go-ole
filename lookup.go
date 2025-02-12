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

// LookupClassIdByProgramId retrieves Class Identifier with the given Program Identifier.
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
func LookupClassIdByProgramId(programId string) (classId *windows.GUID, err error) {
	var guid windows.GUID
	guidPtr := unsafe.Pointer(&guid)
	lookup := windows.UTF16PtrFromString(programId)
	lookupPtr := unsafe.Pointer(lookup)

	hr, _, _ := procCLSIDFromProgID.Call(uintptr(lookupPtr), uintptr(guidPtr))
	switch hr {
	case windows.S_OK:
		return &guid, nil
	case windows.CO_E_CLASSSTRING:
		return nil, InvalidClassIdForProgramId
	case windows.REGDB_E_WRITEREGDB:
		return nil, UnableToWriteClassIdToRegistry
	default:
		return &guid, UnknownLookupArgument
	}
}

// LookupClassIdByGUIDString retrieves Class ID from string representation.
//
// This is technically the string version of the GUID and will convert the string to object.
//
// COM function: CLSIDFromString
func LookupClassIdByGUIDString(guid string) (classId *windows.GUID, err error) {
	var tempGUID windows.GUID
	guidPtr := unsafe.Pointer(&tempGUID)
	lookup := windows.UTF16PtrFromString(guid)
	lookupPtr := unsafe.Pointer(lookup)
	hr, _, _ := procCLSIDFromString.Call(uintptr(lookupPtr), uintptr(guidPtr))
	switch hr {
	case windows.NOERROR:
		return &tempGUID, nil
	case windows.CO_E_CLASSSTRING:
		return nil, ImproperlyFormattedGUID
	case windows.E_INVALIDARG:
		return nil, InvalidLookupArgument
	default:
		return &tempGUID, UnknownLookupArgument
	}
}

// LookupClassId retrieves class ID whether given is program ID or application string.
//
// Helper that provides check against both Class ID from Program ID and Class ID from string. It is
// faster, if you know which you are using, to use the individual functions, but this will check
// against available functions for you.
func LookupClassId(lookup string) (classId *windows.GUID, err error) {
	classId, err = LookupClassIdByProgramId(lookup)
	if err != nil {
		classId, err = LookupClassIdByGUIDString(lookup)
		if err != nil {
			return
		}
	}
	return
}

// ClassIDFrom retrieves class ID whether given is program ID or application string.
func ClassIDFrom(lookup string) (classID *windows.GUID, err error) {
	return LookupClassId(lookup)
}

// InterfaceIdFromString returns GUID from value returned by StringFromInterfaceId.
//
// COM function: IIDFromString
func InterfaceIdFromString(interfaceId string) (classId *windows.GUID, err error) {
	var guid windows.GUID

	lpsz := uintptr(unsafe.Pointer(windows.StringToUTF16Ptr(progId)))
	hr, _, _ := procIIDFromString.Call(lpsz, uintptr(unsafe.Pointer(&guid)))
	if hr != 0 {
		err = NewError(hr)
	}
	clsid = &guid
	return
}

// StringFromClassId returns GUID formated string from GUID object.
//
// COM function: StringFromCLSID
func StringFromClassId(classId *windows.GUID) (str string, err error) {
	var p *uint16
	ptr := unsafe.Pointer(&p)
	classIdPtr := unsafe.Pointer(classId)
	hr, _, _ := procStringFromCLSID.Call(uintptr(classIdPtr), uintptr(ptr))
	if hr != 0 {
		err = NewError(hr)
	}
	str = LpOleStrToString(p)
	return
}

// StringFromInterfaceId returns GUID formatted string from GUID object.
//
// COM function: StringFromIID
func StringFromInterfaceId(iid *windows.GUID) (str string, err error) {
	var p *uint16
	hr, _, _ := procStringFromIID.Call(uintptr(unsafe.Pointer(iid)), uintptr(unsafe.Pointer(&p)))
	if hr != 0 {
		err = NewError(hr)
	}
	str = LpOleStrToString(p)
	return
}
