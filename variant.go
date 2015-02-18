// +build windows

package ole

import "unsafe"

func NewVariant(vt VT, val int64) VARIANT {
	return VARIANT{VT: vt, Val: val}
}

func (v *VARIANT) ToIUnknown() *IUnknown {
	if v.VT != VT_UNKNOWN {
		return nil
	}
	return (*IUnknown)(unsafe.Pointer(uintptr(v.Val)))
}

func (v *VARIANT) ToIDispatch() *IDispatch {
	if v.VT != VT_DISPATCH {
		return nil
	}
	return (*IDispatch)(unsafe.Pointer(uintptr(v.Val)))
}

func (v *VARIANT) ToArray() *SafeArrayConversion {
	if v.VT != VT_SAFEARRAY {
		return nil
	}
	var safeArray *SafeArray = (*SafeArray)(unsafe.Pointer(uintptr(v.Val)))
	return &SafeArrayConversion{safeArray}
}

func (v *VARIANT) ToString() string {
	if v.VT != VT_BSTR {
		return ""
	}
	return BstrToString(*(**uint16)(unsafe.Pointer(&v.Val)))
}

func (v *VARIANT) Clear() error {
	return VariantClear(v)
}

// Returns v's value based on its VALTYPE.
// Currently supported types: 2- and 4-byte integers, strings, bools.
// Note that 64-bit integers, datetimes, and other types are stored as strings
// and will be returned as strings.
func (v *VARIANT) Value() interface{} {
	switch v.VT {
	case VT_I2, VT_I4:
		return v.Val
	case VT_BSTR:
		return v.ToString()
	case VT_BOOL:
		return v.Val != 0
	}
	return nil
}
