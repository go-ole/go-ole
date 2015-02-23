package ole

import (
	"syscall"
	"unsafe"
)

type IDispatch struct {
	lpVtbl *pIDispatchVtbl
}

type pIDispatchVtbl struct {
	pQueryInterface   uintptr
	pAddRef           uintptr
	pRelease          uintptr
	pGetTypeInfoCount uintptr
	pGetTypeInfo      uintptr
	pGetIDsOfNames    uintptr
	pInvoke           uintptr
}

func (v *IDispatch) QueryInterface(iid *GUID) (disp *IDispatch, err error) {
	disp, err = queryInterface((*IUnknown)(unsafe.Pointer(v)), iid)
	return
}

func (v *IDispatch) MustQueryInterface(iid *GUID) (disp *IDispatch) {
	disp, _ = queryInterface((*IUnknown)(unsafe.Pointer(v)), iid)
	return
}

func (v *IDispatch) AddRef() int32 {
	return addRef((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IDispatch) Release() int32 {
	return release((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IDispatch) GetIDsOfName(names []string) (dispid []int32, err error) {
	dispid, err = getIDsOfName(v, names)
	return
}

func (v *IDispatch) Invoke(dispid int32, dispatch int16, params ...interface{}) (result *VARIANT, err error) {
	result, err = invoke(v, dispid, dispatch, params...)
	return
}

func (v *IDispatch) GetTypeInfoCount() (c uint32, err error) {
	c, err = getTypeInfoCount(v)
	return
}

func (v *IDispatch) GetTypeInfo() (tinfo *ITypeInfo, err error) {
	tinfo, err = getTypeInfo(v)
	return
}

func getIDsOfName(disp *IDispatch, names []string) (dispid []int32, err error) {
	wnames := make([]*uint16, len(names))
	for i := 0; i < len(names); i++ {
		wnames[i] = syscall.StringToUTF16Ptr(names[i])
	}
	dispid = make([]int32, len(names))
	hr, _, _ := syscall.Syscall6(
		disp.lpVtbl.pGetIDsOfNames,
		6,
		uintptr(unsafe.Pointer(disp)),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(unsafe.Pointer(&wnames[0])),
		uintptr(len(names)),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&dispid[0])))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func getTypeInfoCount(disp *IDispatch) (c uint32, err error) {
	hr, _, _ := syscall.Syscall(
		disp.lpVtbl.pGetTypeInfoCount,
		2,
		uintptr(unsafe.Pointer(disp)),
		uintptr(unsafe.Pointer(&c)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func getTypeInfo(disp *IDispatch) (tinfo *ITypeInfo, err error) {
	hr, _, _ := syscall.Syscall(
		disp.lpVtbl.pGetTypeInfo,
		3,
		uintptr(unsafe.Pointer(disp)),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&tinfo)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func invoke(disp *IDispatch, dispid int32, dispatch int16, params ...interface{}) (result *VARIANT, err error) {
	var dispparams DISPPARAMS

	if dispatch&DISPATCH_PROPERTYPUT != 0 {
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
			switch v.(type) {
			case bool:
				if v.(bool) {
					vargs[n] = VARIANT{VT_BOOL, 0, 0, 0, 0xffff}
				} else {
					vargs[n] = VARIANT{VT_BOOL, 0, 0, 0, 0}
				}
			case *bool:
				vargs[n] = VARIANT{VT_BOOL | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*bool))))}
			case byte:
				vargs[n] = VARIANT{VT_I1, 0, 0, 0, int64(v.(byte))}
			case *byte:
				vargs[n] = VARIANT{VT_I1 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*byte))))}
			case int16:
				vargs[n] = VARIANT{VT_I2, 0, 0, 0, int64(v.(int16))}
			case *int16:
				vargs[n] = VARIANT{VT_I2 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*int16))))}
			case uint16:
				vargs[n] = VARIANT{VT_UI2, 0, 0, 0, int64(v.(int16))}
			case *uint16:
				vargs[n] = VARIANT{VT_UI2 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*uint16))))}
			case int, int32:
				vargs[n] = VARIANT{VT_UI4, 0, 0, 0, int64(v.(int))}
			case *int, *int32:
				vargs[n] = VARIANT{VT_I4 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*int))))}
			case uint, uint32:
				vargs[n] = VARIANT{VT_UI4, 0, 0, 0, int64(v.(uint))}
			case *uint, *uint32:
				vargs[n] = VARIANT{VT_UI4 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*uint))))}
			case int64:
				vargs[n] = VARIANT{VT_I8, 0, 0, 0, int64(v.(int64))}
			case *int64:
				vargs[n] = VARIANT{VT_I8 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*int64))))}
			case uint64:
				vargs[n] = VARIANT{VT_UI8, 0, 0, 0, v.(int64)}
			case *uint64:
				vargs[n] = VARIANT{VT_UI8 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*uint64))))}
			case float32:
				vargs[n] = VARIANT{VT_R4, 0, 0, 0, int64(v.(float32))}
			case *float32:
				vargs[n] = VARIANT{VT_R4 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*float32))))}
			case float64:
				vargs[n] = VARIANT{VT_R8, 0, 0, 0, int64(v.(float64))}
			case *float64:
				vargs[n] = VARIANT{VT_R8 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*float64))))}
			case string:
				vargs[n] = VARIANT{VT_BSTR, 0, 0, 0, int64(uintptr(unsafe.Pointer(SysAllocString(v.(string)))))}
			case *string:
				vargs[n] = VARIANT{VT_BSTR | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*string))))}
			case *IDispatch:
				vargs[n] = VARIANT{VT_DISPATCH, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*IDispatch))))}
			case **IDispatch:
				vargs[n] = VARIANT{VT_DISPATCH | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(**IDispatch))))}
			case nil:
				vargs[n] = VARIANT{VT_NULL, 0, 0, 0, 0}
			case *VARIANT:
				vargs[n] = VARIANT{VT_VARIANT | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*VARIANT))))}
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
		disp.lpVtbl.pInvoke,
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
		err = NewError(hr)
		if excepInfo.bstrDescription != nil {
			bs := UTF16PtrToString(excepInfo.bstrDescription)
			println(bs)
		}
	}
	for _, varg := range vargs {
		if varg.VT == VT_BSTR && varg.Val != 0 {
			SysFreeString(((*int16)(unsafe.Pointer(uintptr(varg.Val)))))
		}
		/*
			if varg.VT == (VT_BSTR|VT_BYREF) && varg.Val != 0 {
				*(params[n].(*string)) = UTF16PtrToString((*uint16)(unsafe.Pointer(uintptr(varg.Val))))
				println(*(params[n].(*string)))
			}
		*/
	}
	return
}
