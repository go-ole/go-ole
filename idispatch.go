package ole

import (
	"github.com/go-ole/go-ole/legacy"
	"math/big"
	"syscall"
	"time"
	"unsafe"
)

type IsIDispatch interface {
	IsIUnknown
	GetTypeInfoCountAddress() uintptr
	GetTypeInfoAddress() uintptr
	GetIDsOfNamesAddress() uintptr
	InvokeAddress() uintptr
}

type IDispatch struct {
	QueryInterface   uintptr
	AddRef           uintptr
	Release          uintptr
	GetTypeInfoCount uintptr
	GetTypeInfo      uintptr
	GetIDsOfNames    uintptr
	Invoke           uintptr
}

func (v *IDispatch) QueryInterfaceAddress() uintptr {
	return v.QueryInterface
}

func (v *IDispatch) AddRefAddress() uintptr {
	return v.QueryInterface
}

func (v *IDispatch) ReleaseAddress() uintptr {
	return v.QueryInterface
}

func (v *IDispatch) GetTypeInfoCountAddress() uintptr {
	return v.GetTypeInfoCount
}

func (v *IDispatch) GetTypeInfoAddress() uintptr {
	return v.GetTypeInfo
}

func (v *IDispatch) GetIDsOfNamesAddress() uintptr {
	return v.GetIDsOfNames
}

func (v *IDispatch) InvokeAddress() uintptr {
	return v.Invoke
}

func GetIDsOfNameOnIDispatch(dispatch *IsIDispatch, names []string) (dispid []int32, err error) {
	wnames := make([]*uint16, len(names))
	for i := 0; i < len(names); i++ {
		wnames[i] = windows.UTF16PtrFromString(names[i])
	}
	dispid = make([]int32, len(names))
	namelen := uint32(len(names))
	hr, _, _ := windows.Syscall6(
		dispatch.GetIDsOfNamesAddress(),
		6,
		uintptr(unsafe.Pointer(dispatch)),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(unsafe.Pointer(&wnames[0])),
		uintptr(namelen),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&dispid[0])))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func InvokeOnIDispatch(dispatch *IsIDispatch, dispid int32, dispatch int16, params ...interface{}) (result *VARIANT, err error) {
	result, err = invoke(v, dispid, dispatch, params...)
	return
}

func GetTypeInfoCountOnIDispatch(dispatch *IsIDispatch) (c uint32, err error) {
	c, err = getTypeInfoCount(v)
	return
}

func GetTypeInfoOnIDispatch(dispatch *IsIDispatch) (tinfo *legacy.ITypeInfo, err error) {
	tinfo, err = getTypeInfo(v)
	return
}

// GetSingleIDOfName is a helper that returns single display ID for IDispatch name.
//
// This replaces the common pattern of attempting to get a single name from the list of available
// IDs. It gives the first ID, if it is available.
func (v *IDispatch) GetSingleIDOfName(name string) (displayID int32, err error) {
	var displayIDs []int32
	displayIDs, err = v.GetIDsOfName([]string{name})
	if err != nil {
		return
	}
	displayID = displayIDs[0]
	return
}

// InvokeWithOptionalArgs accepts arguments as an array, works like Invoke.
//
// Accepts name and will attempt to retrieve Display ID to pass to Invoke.
//
// Passing params as an array is a workaround that could be fixed in later versions of Go that
// prevent passing empty params. During testing it was discovered that this is an acceptable way of
// getting around not being able to pass params normally.
func (v *IDispatch) InvokeWithOptionalArgs(name string, dispatch int16, params []interface{}) (result *VARIANT, err error) {
	displayID, err := v.GetSingleIDOfName(name)
	if err != nil {
		return
	}

	if len(params) < 1 {
		result, err = v.Invoke(displayID, dispatch)
	} else {
		result, err = v.Invoke(displayID, dispatch, params...)
	}

	return
}

// CallMethod invokes named function with arguments on object.
func (v *IDispatch) CallMethod(name string, params ...interface{}) (*VARIANT, error) {
	return v.InvokeWithOptionalArgs(name, legacy.DISPATCH_METHOD, params)
}

// GetProperty retrieves the property with the name with the ability to pass arguments.
//
// Most of the time you will not need to pass arguments as most objects do not allow for this
// feature. Or at least, should not allow for this feature. Some servers don't follow best practices
// and this is provided for those edge cases.
func (v *IDispatch) GetProperty(name string, params ...interface{}) (*VARIANT, error) {
	return v.InvokeWithOptionalArgs(name, legacy.DISPATCH_PROPERTYGET, params)
}

// PutProperty attempts to mutate a property in the object.
func (v *IDispatch) PutProperty(name string, params ...interface{}) (*VARIANT, error) {
	return v.InvokeWithOptionalArgs(name, legacy.DISPATCH_PROPERTYPUT, params)
}

func getIDsOfName(disp *IDispatch, names []string) (dispid []int32, err error) {
	return
}

func getTypeInfoCount(disp *ole.IDispatch) (c uint32, err error) {
	hr, _, _ := syscall.Syscall(
		disp.VTable().GetTypeInfoCount,
		2,
		uintptr(unsafe.Pointer(disp)),
		uintptr(unsafe.Pointer(&c)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func getTypeInfo(disp *ole.IDispatch) (tinfo *ITypeInfo, err error) {
	hr, _, _ := syscall.Syscall(
		disp.VTable().GetTypeInfo,
		3,
		uintptr(unsafe.Pointer(disp)),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&tinfo)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func invoke(disp *ole.IDispatch, dispid int32, dispatch int16, params ...interface{}) (result *VARIANT, err error) {
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

	result = new(VARIANT)
	var excepInfo EXCEPINFO
	VariantInit(result)
	hr, _, _ := syscall.Syscall9(
		disp.VTable().Invoke,
		9,
		uintptr(unsafe.Pointer(disp)),
		uintptr(dispid),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(GetUserDefaultLCID()),
		uintptr(dispatch),
		uintptr(unsafe.Pointer(&dispparams)),
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
			*(params[n].(*string)) = LpOleStrToString(*(**uint16)(unsafe.Pointer(uintptr(varg.Val))))
		}
	}
	return
}
