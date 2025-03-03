//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
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
	// IUnknown
	QueryInterface uintptr
	addRef         uintptr
	release        uintptr
	// IDispatch
	getTypeInfoCount uintptr
	getTypeInfo      uintptr
	getIDsOfNames    uintptr
	invoke           uintptr
}

func (v *IDispatch) QueryInterfaceAddress() uintptr {
	return v.QueryInterface
}

func (v *IDispatch) AddRefAddress() uintptr {
	return v.addRef
}

func (v *IDispatch) ReleaseAddress() uintptr {
	return v.release
}

func (v *IDispatch) GetTypeInfoCountAddress() uintptr {
	return v.getTypeInfoCount
}

func (v *IDispatch) GetTypeInfoAddress() uintptr {
	return v.getTypeInfo
}

func (v *IDispatch) GetIDsOfNamesAddress() uintptr {
	return v.getIDsOfNames
}

func (v *IDispatch) InvokeAddress() uintptr {
	return v.invoke
}

func (obj *IDispatch) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *IDispatch) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *IDispatch) HasTypeInfo() bool {
	var ret uint
	hr, _, _ := windows.Syscall(
		obj.getTypeInfoCount,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&ret)),
		0)

	if hr == windows.E_NOTIMPL {
		return false
	}

	return ret == 1
}

func (obj *IDispatch) GetTypeInfo() (ret *ITypeInfo) {
	var ret uint
	hr, _, _ := windows.Syscall6(
		obj.getTypeInfo,
		4,
		uintptr(unsafe.Pointer(obj)),
		uintptr(0),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&ret)),
		0,
		0)

	if hr == windows.DISP_E_BADINDEX {
		return nil
	}

	return
}

func (obj *IDispatch) GetIDsOfNames(names []string) (ret map[string]int32, err error) {
	wNames := make([]*uint16, len(names))
	for i := 0; i < len(names); i++ {
		wNames[i] = windows.UTF16PtrFromString(names[i])
	}
	dispid = make([]int32, len(names))
	namelen := uint32(len(names))
	hr, _, _ := windows.Syscall6(
		obj.getIDsOfNames,
		6,
		uintptr(unsafe.Pointer(dispatch)),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(unsafe.Pointer(&wNames[0])),
		uintptr(namelen),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&dispid[0])))

	if hr != windows.S_OK {
		err = hr
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

func (obj *IDispatch) Invoke(name string, dispatch int16) (result *VARIANT, err error) {

}

// InvokeWithOptionalArgs accepts arguments as an array, works like Invoke.
//
// Accepts name and will attempt to retrieve Display ID to pass to Invoke.
//
// Passing params as an array is a workaround that could be fixed in later versions of Go that
// prevent passing empty params. During testing it was discovered that this is an acceptable way of
// getting around not being able to pass params normally.
func (obj *IDispatch) InvokeWithOptionalArgs(name string, dispatch int16, params []interface{}) (result *VARIANT, err error) {
	displayID, err := v.GetSingleIDOfName(name)
	if err != nil {
		return
	}

	if len(params) < 1 {
		result, err = obj.Invoke(displayID, dispatch)
	} else {
		result, err = obj.Invoke(displayID, dispatch, params...)
	}

	return
}

// CallMethod invokes named function with arguments on object.
func (v *IDispatch) CallMethod(name string, params ...interface{}) (*VARIANT, error) {
	return v.InvokeWithOptionalArgs(name, DISPATCH_METHOD, params)
}

// GetProperty retrieves the property with the name with the ability to pass arguments.
//
// Most of the time you will not need to pass arguments as most objects do not allow for this
// feature. Or at least, should not allow for this feature. Some servers don't follow best practices
// and this is provided for those edge cases.
func (v *IDispatch) GetProperty(name string, params ...interface{}) (*VARIANT, error) {
	return v.InvokeWithOptionalArgs(name, DISPATCH_PROPERTYGET, params)
}

// PutProperty attempts to mutate a property in the object.
func (v *IDispatch) PutProperty(name string, params ...interface{}) (*VARIANT, error) {
	return v.InvokeWithOptionalArgs(name, DISPATCH_PROPERTYPUT, params)
}

func QueryIDispatchFromIUnknown(unknown *IsIUnknown) (dispatch *IDispatch, err error) {
	if unknown == nil {
		return nil, ComInterfaceIsNilPointer
	}

	dispatch, err = QueryInterfaceOnIUnknown[IDispatch](unknown, IID_IDispatch)
	if err != nil {
		return nil, err
	}
	return
}

func InvokeOnIDispatch(obj *IDispatchAddresses, displayId int32, dispatch int16, params ...interface{}) (result *VARIANT, err error) {
	result, err = invoke(obj, displayId, dispatch, params...)
	return
}

func getIDsOfName(disp *IDispatch, names []string) (dispid []int32, err error) {
	return
}

func MakeDisplayParams(dispatch int16, params ...interface{}) DISPPARAMS {
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

	var vargs []VARIANT
	if len(params) > 0 {
		vargs = make([]VARIANT, len(params))
		for i, v := range params {
			//n := len(params)-i-1
			n := len(params) - i - 1
			VariantInit(&vargs[n])
			switch vv := v.(type) {
			case bool:
				if vv {
					vargs[n] = NewVariant(VT_BOOL, 0xffff)
				} else {
					vargs[n] = NewVariant(VT_BOOL, 0)
				}
			case *bool:
				vargs[n] = NewVariant(VT_BOOL|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*bool)))))
			case uint8:
				vargs[n] = NewVariant(VT_UI1, int64(v.(uint8)))
			case *uint8:
				vargs[n] = NewVariant(VT_UI1|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint8)))))
			case int8:
				vargs[n] = NewVariant(VT_I1, int64(v.(int8)))
			case *int8:
				vargs[n] = NewVariant(VT_I1|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int8)))))
			case int16:
				vargs[n] = NewVariant(VT_I2, int64(v.(int16)))
			case *int16:
				vargs[n] = NewVariant(VT_I2|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int16)))))
			case uint16:
				vargs[n] = NewVariant(VT_UI2, int64(v.(uint16)))
			case *uint16:
				vargs[n] = NewVariant(VT_UI2|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint16)))))
			case int32:
				vargs[n] = NewVariant(VT_I4, int64(v.(int32)))
			case *int32:
				vargs[n] = NewVariant(VT_I4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int32)))))
			case uint32:
				vargs[n] = NewVariant(VT_UI4, int64(v.(uint32)))
			case *uint32:
				vargs[n] = NewVariant(VT_UI4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint32)))))
			case int64:
				vargs[n] = NewVariant(VT_I8, int64(v.(int64)))
			case *int64:
				vargs[n] = NewVariant(VT_I8|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int64)))))
			case uint64:
				vargs[n] = NewVariant(VT_UI8, int64(uintptr(v.(uint64))))
			case *uint64:
				vargs[n] = NewVariant(VT_UI8|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint64)))))
			case int:
				vargs[n] = NewVariant(VT_I4, int64(v.(int)))
			case *int:
				vargs[n] = NewVariant(VT_I4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int)))))
			case uint:
				vargs[n] = NewVariant(VT_UI4, int64(v.(uint)))
			case *uint:
				vargs[n] = NewVariant(VT_UI4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint)))))
			case float32:
				vargs[n] = NewVariant(VT_R4, *(*int64)(unsafe.Pointer(&vv)))
			case *float32:
				vargs[n] = NewVariant(VT_R4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*float32)))))
			case float64:
				vargs[n] = NewVariant(VT_R8, *(*int64)(unsafe.Pointer(&vv)))
			case *float64:
				vargs[n] = NewVariant(VT_R8|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*float64)))))
			case *big.Int:
				vargs[n] = NewVariant(VT_DECIMAL, v.(*big.Int).Int64())
			case string:
				vargs[n] = NewVariant(VT_BSTR, int64(uintptr(unsafe.Pointer(SysAllocStringLen(v.(string))))))
			case *string:
				vargs[n] = NewVariant(VT_BSTR|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*string)))))
			case time.Time:
				s := vv.Format("2006-01-02 15:04:05")
				vargs[n] = NewVariant(VT_BSTR, int64(uintptr(unsafe.Pointer(SysAllocStringLen(s)))))
			case *time.Time:
				s := vv.Format("2006-01-02 15:04:05")
				vargs[n] = NewVariant(VT_BSTR|VT_BYREF, int64(uintptr(unsafe.Pointer(&s))))
			case *ole.IDispatch:
				vargs[n] = NewVariant(VT_DISPATCH, int64(uintptr(unsafe.Pointer(v.(*ole.IDispatch)))))
			case **ole.IDispatch:
				vargs[n] = NewVariant(VT_DISPATCH|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(**ole.IDispatch)))))
			case nil:
				vargs[n] = NewVariant(VT_NULL, 0)
			case *VARIANT:
				vargs[n] = NewVariant(VT_VARIANT|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*VARIANT)))))
			case []byte:
				safeByteArray := safeArrayFromByteSlice(v.([]byte))
				vargs[n] = NewVariant(VT_ARRAY|VT_UI1, int64(uintptr(unsafe.Pointer(safeByteArray))))
				defer VariantClear(&vargs[n])
			case []string:
				safeByteArray := safeArrayFromStringSlice(v.([]string))
				vargs[n] = NewVariant(VT_ARRAY|VT_BSTR, int64(uintptr(unsafe.Pointer(safeByteArray))))
				defer VariantClear(&vargs[n])
			default:
				panic("unknown type")
			}
		}
		dispparams.rgvarg = uintptr(unsafe.Pointer(&vargs[0]))
		dispparams.cArgs = uint32(len(params))
	}

	return dispparams
}

func invoke(disp *IDispatch, dispid int32, dispatch int16, params ...interface{}) (result *VARIANT, err error) {
	dispParams := MakeDisplayParams(dispatch, params...)
	result = new(VARIANT)
	var excepInfo EXCEPINFO
	VariantInit(result)
	hr, _, _ := windows.Syscall9(
		disp.invoke,
		9,
		uintptr(unsafe.Pointer(disp)),
		uintptr(dispid),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(GetUserDefaultLCID()),
		uintptr(dispatch),
		uintptr(unsafe.Pointer(&dispParams)),
		uintptr(unsafe.Pointer(result)),
		uintptr(unsafe.Pointer(&excepInfo)),
		0)
	if hr != 0 {
		excepInfo.renderStrings()
		excepInfo.Clear()
		err = NewErrorWithSubError(hr, excepInfo.description, excepInfo)
	}
	for i, varg := range vargs {
		n := len(params) - i - 1
		if varg.VT == VT_BSTR && varg.Val != 0 {
			SysFreeString(((*int16)(unsafe.Pointer(uintptr(varg.Val)))))
		}
		if varg.VT == (VT_BSTR|VT_BYREF) && varg.Val != 0 {
			*(params[n].(*string)) = windows.UTF16PtrToString(*(**uint16)(unsafe.Pointer(uintptr(varg.Val))))
		}
	}
	return
}
