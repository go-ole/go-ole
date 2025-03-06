//go:build windows

package ole

import (
	"errors"
	"golang.org/x/sys/windows"
	"syscall"
	"unsafe"
)

const (
	DISPATCH_METHOD         int16 = 1
	DISPATCH_PROPERTYGET          = 2
	DISPATCH_PROPERTYPUT          = 4
	DISPATCH_PROPERTYPUTREF       = 8
)

const (
	DISPID_UNKNOWN     int32 = -1
	DISPID_VALUE             = 0
	DISPID_PROPERTYPUT       = -3
	DISPID_NEWENUM           = -4
	DISPID_EVALUATE          = -5
	DISPID_CONSTRUCTOR       = -6
	DISPID_DESTRUCTOR        = -7
	DISPID_COLLECT           = -8
)

// DISPPARAMS are the arguments that passed to methods or property.
type DISPPARAMS struct {
	rgvarg            uintptr
	rgdispidNamedArgs uintptr
	cArgs             uint32
	cNamedArgs        uint32
}

// PARAMDATA defines parameter data type.
type PARAMDATA struct {
	Name *int16
	Vt   uint16
}

// METHODDATA defines method info.
type METHODDATA struct {
	Name     *uint16
	Data     *PARAMDATA
	Dispid   int32
	Meth     uint32
	CC       int32
	CArgs    uint32
	Flags    uint16
	VtReturn uint32
}

// INTERFACEDATA defines interface info.
type INTERFACEDATA struct {
	MethodData *METHODDATA
	CMembers   uint32
}

// TYPEDESC defines data type.
type TYPEDESC struct {
	Hreftype uint32
	VT       uint16
}

// IDLDESC defines IDL info.
type IDLDESC struct {
	DwReserved uint32
	WIDLFlags  uint16
}

// TYPEATTR defines type info.
type TYPEATTR struct {
	Guid             windows.GUID
	Lcid             uint32
	dwReserved       uint32
	MemidConstructor int32
	MemidDestructor  int32
	LpstrSchema      *uint16
	CbSizeInstance   uint32
	Typekind         int32
	CFuncs           uint16
	CVars            uint16
	CImplTypes       uint16
	CbSizeVft        uint16
	CbAlignment      uint16
	WTypeFlags       uint16
	WMajorVerNum     uint16
	WMinorVerNum     uint16
	TdescAlias       TYPEDESC
	IdldescType      IDLDESC
}

type IDispatchAddresses interface {
	IsIUnknown
	GetTypeInfoCountAddress() uintptr
	GetTypeInfoAddress() uintptr
	GetIDsOfNamesAddress() uintptr
	InvokeAddress() uintptr
}

type IDispatch struct {
	VirtualTable *IDispatchVirtualTable
}

type IDispatchVirtualTable struct {
	// IUnknown
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
	// IDispatch
	GetTypeInfoCount uintptr
	GetTypeInfo      uintptr
	GetIDsOfNames    uintptr
	Invoke           uintptr
}

func (obj *IDispatch) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *IDispatch) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

func (obj *IDispatch) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

func (obj *IDispatch) GetTypeInfoCountAddress() uintptr {
	return obj.VirtualTable.GetTypeInfoCount
}

func (obj *IDispatch) GetTypeInfoAddress() uintptr {
	return obj.VirtualTable.GetTypeInfo
}

func (obj *IDispatch) GetIDsOfNamesAddress() uintptr {
	return obj.VirtualTable.GetIDsOfNames
}

func (obj *IDispatch) InvokeAddress() uintptr {
	return obj.VirtualTable.Invoke
}

func (obj *IDispatch) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IDispatch) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IDispatch) HasTypeInfo() bool {
	var ret uint
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetTypeInfoCount,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&ret)),
		0)

	if windows.Handle(hr) == windows.E_NOTIMPL {
		return false
	}

	return ret == 1
}

func (obj *IDispatch) GetTypeInfo() (ret *ITypeInfo) {
	hr, _, _ := syscall.Syscall6(
		obj.VirtualTable.GetTypeInfo,
		4,
		uintptr(unsafe.Pointer(obj)),
		uintptr(0),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&ret)),
		0,
		0)

	if windows.Handle(hr) == windows.DISP_E_BADINDEX {
		return nil
	}

	return
}

func (obj *IDispatch) GetIDsOfNames(names []string) (ret map[string]int32, err error) {
	wNames := make([]*uint16, len(names))
	for i := 0; i < len(names); i++ {
		wNames[i], _ = windows.UTF16PtrFromString(names[i])
	}
	dispid := make([]int32, len(names))
	namelen := uint32(len(names))
	hr, _, _ := syscall.Syscall6(
		obj.VirtualTable.GetIDsOfNames,
		6,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&IID_NULL)),
		uintptr(unsafe.Pointer(&wNames[0])),
		uintptr(namelen),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&dispid[0])))

	if hr != 0 {
		err = windows.Errno(hr)
		return
	}

	ret = make(map[string]int32, len(names))
	for i := 0; i < int(namelen); i++ {
		ret[names[i]] = dispid[i]
	}

	return
}

// GetSingleIDOfName is a helper that returns single display ID for IDispatch name.
//
// This replaces the common pattern of attempting to get a single name from the list of available IDs. It gives the
// first ID, if it is available.
func (obj *IDispatch) GetSingleIDOfName(name string) (displayID int32, err error) {
	displayIDs, err := obj.GetIDsOfNames([]string{name})
	if err != nil {
		return
	}
	displayID = displayIDs[name]
	return
}

func (obj *IDispatch) Invoke(name string, dispatch int16, params ...*VARIANT) (result *VARIANT, err error) {
	displayID, err := obj.GetSingleIDOfName(name)
	if err != nil {
		return
	}
	return InvokeOnIDispatch(obj, displayID, dispatch, params...)
}

// CallMethod invokes named function with arguments on object.
func (obj *IDispatch) CallMethod(name string, params ...*VARIANT) (*VARIANT, error) {
	return obj.Invoke(name, DISPATCH_METHOD, params...)
}

// MustCallMethod calls method on IDispatch with parameters or panics.
func (obj *IDispatch) MustCallMethod(name string, params ...*VARIANT) (result *VARIANT) {
	result, err := obj.CallMethod(name, params...)
	if err != nil {
		panic(err.Error())
	}
	return
}

// GetProperty retrieves the property with the name with the ability to pass arguments.
//
// Most of the time you will not need to pass arguments as most objects do not allow for this
// feature. Or at least, should not allow for this feature. Some servers don't follow best practices
// and this is provided for those edge cases.
func (obj *IDispatch) GetProperty(name string, params ...*VARIANT) (*VARIANT, error) {
	return obj.Invoke(name, DISPATCH_PROPERTYGET, params...)
}

// MustGetProperty retrieves property from IDispatch or panics.
func (obj *IDispatch) MustGetProperty(name string, params ...*VARIANT) (result *VARIANT) {
	result, err := obj.GetProperty(name, params...)
	if err != nil {
		panic(err.Error())
	}
	return
}

// PutProperty attempts to mutate a property in the object.
func (obj *IDispatch) PutProperty(name string, params ...*VARIANT) (*VARIANT, error) {
	return obj.Invoke(name, DISPATCH_PROPERTYPUT, params...)
}

// MustPutProperty mutates property or panics.
func (obj *IDispatch) MustPutProperty(name string, params ...*VARIANT) (result *VARIANT) {
	result, err := obj.PutProperty(name, params...)
	if err != nil {
		panic(err.Error())
	}
	return
}

// PutPropertyRef mutates property reference.
func (obj *IDispatch) PutPropertyRef(name string, params ...*VARIANT) (result *VARIANT, err error) {
	return obj.Invoke(name, DISPATCH_PROPERTYPUTREF, params...)
}

// MustPutPropertyRef mutates property reference or panics.
func (obj *IDispatch) MustPutPropertyRef(name string, params ...*VARIANT) (result *VARIANT) {
	result, err := obj.PutPropertyRef(name, params...)
	if err != nil {
		panic(err.Error())
	}
	return
}

func QueryIDispatchFromIUnknown(unknown IsIUnknown) (dispatch *IDispatch, err error) {
	if unknown == nil {
		return nil, ComInterfaceIsNilPointer
	}

	dispatch, err = QueryInterfaceOnIUnknown[IDispatch](unknown, IID_IDispatch)
	if err != nil {
		return nil, err
	}
	return
}

func InvokeOnIDispatch(obj IDispatchAddresses, displayId int32, dispatch int16, params ...*VARIANT) (result *VARIANT, err error) {
	dispParams := MakeDisplayParams(dispatch, params...)
	result = new(VARIANT)
	var excepInfo EXCEPINFO
	VariantInit(result)
	hr, _, _ := syscall.Syscall9(
		obj.InvokeAddress(),
		9,
		uintptr(unsafe.Pointer(&obj)),
		uintptr(displayId),
		uintptr(unsafe.Pointer(&IID_NULL)),
		uintptr(GetUserDefaultLCID()),
		uintptr(dispatch),
		uintptr(unsafe.Pointer(&dispParams)),
		uintptr(unsafe.Pointer(&result)),
		uintptr(unsafe.Pointer(&excepInfo)),
		0)
	if hr != 0 {
		excepInfo.renderStrings()
		excepInfo.Clear()
		err = errors.Join(windows.Errno(hr), errors.New(excepInfo.description))
	}
	return
}

func MakeDisplayParams(dispatch int16, params ...*VARIANT) DISPPARAMS {
	var dispparams DISPPARAMS

	if dispatch&DISPATCH_PROPERTYPUT != 0 {
		dispnames := [1]int32{DISPID_PROPERTYPUT}
		dispparams.rgdispidNamedArgs = uintptr(unsafe.Pointer(&dispnames[0]))
		dispparams.cNamedArgs = 1
	} else if dispatch&DISPATCH_PROPERTYPUTREF != 0 {
		dispnames := [1]int32{DISPID_PROPERTYPUT}
		dispparams.rgdispidNamedArgs = uintptr(unsafe.Pointer(&dispnames[0]))
		dispparams.cNamedArgs = 1
	}
	if len(params) > 0 {
		dispparams.rgvarg = uintptr(unsafe.Pointer(&params[0]))
		dispparams.cArgs = uint32(len(params))
	}

	return dispparams
}
