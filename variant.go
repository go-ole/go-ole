//go:build windows

package ole

import (
	"errors"
	"golang.org/x/sys/windows"
	"reflect"
	"sync"
	"time"
	"unsafe"
)

const (
	VariantTypeTrue  int32 = 0xffff
	VariantTypeFalse       = 0
)

var (
	procVariantInit  = modoleaut32.NewProc("VariantInit")
	procVariantClear = modoleaut32.NewProc("VariantClear")
)

func (v *VARIANT) Clear() error {
	return VariantClear(v)
}

// VariantInit initializes variant.
func VariantInit(v *VARIANT) (err error) {
	hr, _, _ := procVariantInit.Call(uintptr(unsafe.Pointer(v)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// VariantClear clears value in Variant settings to VT_EMPTY.
func VariantClear(v *VARIANT) (err error) {
	hr, _, _ := procVariantClear.Call(uintptr(unsafe.Pointer(v)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

type ToVariantCallback func(any) *VARIANT
type FromVariantCallback func(*VARIANT) any

// variantTypeContainer keeps track of mapping for type conversions.
type variantTypeContainer struct {
	lock sync.RWMutex
	to   map[string]ToVariantCallback
	from map[VT]FromVariantCallback
}

// variantTypeContainer stores lock and allows for defining transitions from VARIANT to native and native to VARIANT.
var conversions *variantTypeContainer = &variantTypeContainer{to: make(map[string]ToVariantCallback), from: make(map[VT]FromVariantCallback)}

var (
	VariantTypeMismatch    = errors.New("variant type mismatch")
	UnsupportedNativeType  = errors.New("native type is unsupported")
	UnsupportedVariantType = errors.New("variant type is unsupported")
)

// WrapVariant allows supporting any native type to a Variant.
//
// If the native type is unsupported, then you may need to call RegisterVariantConverter or RegisterToVariantConverter
// in order to handle the conversion to a VARIANT.
func WrapVariant[T any](val T) (*VARIANT, error) {
	conversions.lock.RLock()
	defer conversions.lock.RUnlock()
	callback, ok := conversions.to[reflect.TypeFor[T]().Name()]
	if !ok {
		return nil, UnsupportedNativeType
	}
	return callback(val), nil
}

//func MarshalDispatchParams(params ...interface{}) ([]VARIANT, error) {
//	var vargs []VARIANT
//	if len(params) > 0 {
//		vargs = make([]VARIANT, len(params))
//		for i, v := range params {
//			//n := len(params)-i-1
//			n := len(params) - i - 1
//			VariantInit(&vargs[n])
//			switch vv := v.(type) {
//			case bool:
//				if vv {
//					vargs[n] = NewVariant(VT_BOOL, 0xffff)
//				} else {
//					vargs[n] = NewVariant(VT_BOOL, 0)
//				}
//			case *bool:
//				vargs[n] = NewVariant(VT_BOOL|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*bool)))))
//			case uint8:
//				vargs[n] = NewVariant(VT_UI1, int64(v.(uint8)))
//			case *uint8:
//				vargs[n] = NewVariant(VT_UI1|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint8)))))
//			case int8:
//				vargs[n] = NewVariant(VT_I1, int64(v.(int8)))
//			case *int8:
//				vargs[n] = NewVariant(VT_I1|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int8)))))
//			case int16:
//				vargs[n] = NewVariant(VT_I2, int64(v.(int16)))
//			case *int16:
//				vargs[n] = NewVariant(VT_I2|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int16)))))
//			case uint16:
//				vargs[n] = NewVariant(VT_UI2, int64(v.(uint16)))
//			case *uint16:
//				vargs[n] = NewVariant(VT_UI2|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint16)))))
//			case int32:
//				vargs[n] = NewVariant(VT_I4, int64(v.(int32)))
//			case *int32:
//				vargs[n] = NewVariant(VT_I4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int32)))))
//			case uint32:
//				vargs[n] = NewVariant(VT_UI4, int64(v.(uint32)))
//			case *uint32:
//				vargs[n] = NewVariant(VT_UI4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint32)))))
//			case int64:
//				vargs[n] = NewVariant(VT_I8, int64(v.(int64)))
//			case *int64:
//				vargs[n] = NewVariant(VT_I8|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int64)))))
//			case uint64:
//				vargs[n] = NewVariant(VT_UI8, int64(uintptr(v.(uint64))))
//			case *uint64:
//				vargs[n] = NewVariant(VT_UI8|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint64)))))
//			case int:
//				vargs[n] = NewVariant(VT_I4, int64(v.(int)))
//			case *int:
//				vargs[n] = NewVariant(VT_I4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int)))))
//			case uint:
//				vargs[n] = NewVariant(VT_UI4, int64(v.(uint)))
//			case *uint:
//				vargs[n] = NewVariant(VT_UI4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint)))))
//			case float32:
//				vargs[n] = NewVariant(VT_R4, *(*int64)(unsafe.Pointer(&vv)))
//			case *float32:
//				vargs[n] = NewVariant(VT_R4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*float32)))))
//			case float64:
//				vargs[n] = NewVariant(VT_R8, *(*int64)(unsafe.Pointer(&vv)))
//			case *float64:
//				vargs[n] = NewVariant(VT_R8|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*float64)))))
//			case *big.Int:
//				vargs[n] = NewVariant(VT_DECIMAL, v.(*big.Int).Int64())
//			case string:
//				vargs[n] = NewVariant(VT_BSTR, int64(uintptr(unsafe.Pointer(SysAllocStringLen(v.(string))))))
//			case *string:
//				vargs[n] = NewVariant(VT_BSTR|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*string)))))
//			case time.Time:
//				s := vv.Format("2006-01-02 15:04:05")
//				vargs[n] = NewVariant(VT_BSTR, int64(uintptr(unsafe.Pointer(SysAllocStringLen(s)))))
//			case *time.Time:
//				s := vv.Format("2006-01-02 15:04:05")
//				vargs[n] = NewVariant(VT_BSTR|VT_BYREF, int64(uintptr(unsafe.Pointer(&s))))
//			case *ole.IDispatch:
//				vargs[n] = NewVariant(VT_DISPATCH, int64(uintptr(unsafe.Pointer(v.(*ole.IDispatch)))))
//			case **ole.IDispatch:
//				vargs[n] = NewVariant(VT_DISPATCH|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(**ole.IDispatch)))))
//			case nil:
//				vargs[n] = NewVariant(VT_NULL, 0)
//			case *VARIANT:
//				vargs[n] = NewVariant(VT_VARIANT|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*VARIANT)))))
//			case []byte:
//				safeByteArray := safeArrayFromByteSlice(v.([]byte))
//				vargs[n] = NewVariant(VT_ARRAY|VT_UI1, int64(uintptr(unsafe.Pointer(safeByteArray))))
//				defer VariantClear(&vargs[n])
//			case []string:
//				safeByteArray := safeArrayFromStringSlice(v.([]string))
//				vargs[n] = NewVariant(VT_ARRAY|VT_BSTR, int64(uintptr(unsafe.Pointer(safeByteArray))))
//				defer VariantClear(&vargs[n])
//			default:
//				panic("unknown type")
//			}
//		}
//	}
//	return vargs, nil
//}

// WrapParametersWithVariant converts arbitrary values to an array of *VARIANT.
//
// If the native type is unsupported, then you may need to call RegisterVariantConverter or RegisterToVariantConverter
// in order to handle the conversion to a VARIANT.
func WrapParametersWithVariant(params ...any) (args []*VARIANT) {
	conversions.lock.RLock()
	defer conversions.lock.RUnlock()

	if len(params) == 0 {
		return
	}

	args = make([]*VARIANT, len(params))

	for i, v := range params {
		n := len(params) - i - 1
		VariantInit(args[n])

		// Attempt nil since we can't automate that
		if v == nil {
			args[n] = MakeNullVariant()
			continue
		}

		typeName := reflect.TypeOf(v).Name()
		callback, ok := conversions.to[typeName]
		if !ok {
			panic(errors.New(typeName + " is not registered for conversion to the *VARIANT type"))
		}
		args[n] = callback(v)

		//switch vv := v.(type) {
		//case int:
		//	vargs[n] = NewVariant(VT_INT, int64(v.(int)))
		//case *int:
		//	vargs[n] = NewVariant(VT_INT|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int)))))
		//case uint:
		//	vargs[n] = NewVariant(VT_UINT, int64(v.(uint)))
		//case *uint:
		//	vargs[n] = NewVariant(VT_UINT|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint)))))
		//case float32:
		//	vargs[n] = NewVariant(VT_R4, *(*int64)(unsafe.Pointer(&vv)))
		//case *float32:
		//	vargs[n] = NewVariant(VT_R4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*float32)))))
		//case float64:
		//	vargs[n] = NewVariant(VT_R8, *(*int64)(unsafe.Pointer(&vv)))
		//case *float64:
		//	vargs[n] = NewVariant(VT_R8|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*float64)))))
		//case *big.Int:
		//	vargs[n] = NewVariant(VT_DECIMAL, v.(*big.Int).Int64())
		//case string:
		//	vargs[n] = NewVariant(VT_BSTR, int64(uintptr(unsafe.Pointer(SysAllocStringLen(v.(string))))))
		//case *string:
		//	vargs[n] = NewVariant(VT_BSTR|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*string)))))
		//case *ole.IDispatch:
		//	vargs[n] = NewVariant(VT_DISPATCH, int64(uintptr(unsafe.Pointer(v.(*ole.IDispatch)))))
		//case **ole.IDispatch:
		//	vargs[n] = NewVariant(VT_DISPATCH|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(**ole.IDispatch)))))
		//case *VARIANT:
		//	vargs[n] = NewVariant(VT_VARIANT|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*VARIANT)))))
		//case []byte:
		//	safeByteArray := safeArrayFromByteSlice(v.([]byte))
		//	vargs[n] = NewVariant(VT_ARRAY|VT_UI1, int64(uintptr(unsafe.Pointer(safeByteArray))))
		//	defer VariantClear(&vargs[n])
		//case []string:
		//	safeByteArray := safeArrayFromStringSlice(v.([]string))
		//	vargs[n] = NewVariant(VT_ARRAY|VT_BSTR, int64(uintptr(unsafe.Pointer(safeByteArray))))
		//	defer VariantClear(&vargs[n])
		//default:
		//	panic("unknown type")
		//}
	}

	return
}

// UnwrapVariant allows supporting converting from a VARIANT type to a native Go type.
//
// This is done by callback and registered by RegisterVariantConverter or RegisterFromVariantConverter.
func UnwrapVariant[T any](variant *VARIANT) T {
	conversions.lock.RLock()
	defer conversions.lock.RUnlock()
	callback, ok := conversions.from[variant.VT]
	if !ok {
		panic(UnsupportedNativeType)
	}
	return callback(variant).(T)
}

// RegisterVariantConverter registers both conversion for to and from VARIANT.
//
// This function is an attempt to support handling unsupported VARIANT and native types. Also allow for providing a shim
// to fix bugs in the library (please submit a bug report with the corrected implementation).
func RegisterVariantConverter[T any](vt VT, to ToVariantCallback, from FromVariantCallback) error {
	conversions.lock.Lock()
	conversions.from[vt] = from
	conversions.to[reflect.TypeFor[T]().Name()] = to
	conversions.lock.Unlock()
	return nil
}

// RegisterFromVariantConverter registers both conversion from VARIANT.
//
// This function is an attempt to support handling unsupported VARIANT and native types. Also allow for providing a shim
// to fix bugs in the library (please submit a bug report with the corrected implementation).
func RegisterFromVariantConverter(vt VT, from FromVariantCallback) error {
	conversions.lock.Lock()
	conversions.from[vt] = from
	conversions.lock.Unlock()
	return nil
}

// RegisterToVariantConverter registers both conversion to VARIANT.
//
// This function is an attempt to support handling unsupported VARIANT and native types. Also allow for providing a shim
// to fix bugs in the library (please submit a bug report with the corrected implementation).
func RegisterToVariantConverter[T any](to ToVariantCallback) error {
	conversions.lock.Lock()
	conversions.to[reflect.TypeFor[T]().Name()] = to
	conversions.lock.Unlock()
	return nil
}

// DeregisterVariantConverter removes registered type conversion for both conversions to and from VARIANT.
//
// This function is an attempt to support handling unsupported VARIANT and native types. Also allow for providing a shim
// to fix bugs in the library (please submit a bug report with the corrected implementation).
func DeregisterVariantConverter[T any](vt VT) error {
	conversions.lock.Lock()
	delete(conversions.from, vt)
	delete(conversions.to, reflect.TypeFor[T]().Name())
	conversions.lock.Unlock()
	return nil
}

// DeregisterFromVariantConverter removes registered type conversion for converting VARIANT to a native go type.
//
// This function is an attempt to support handling unsupported VARIANT and native types. Also allow for providing a shim
// to fix bugs in the library (please submit a bug report with the corrected implementation).
func DeregisterFromVariantConverter(vt VT) error {
	conversions.lock.Lock()
	delete(conversions.from, vt)
	conversions.lock.Unlock()
	return nil
}

// DeregisterToVariantConverter removes registered type conversion for converting native Go type to a VARIANT.
//
// This function is an attempt to support handling unsupported VARIANT and native types. Also allow for providing a shim
// to fix bugs in the library (please submit a bug report with the corrected implementation).
func DeregisterToVariantConverter[T any]() error {
	conversions.lock.Lock()
	delete(conversions.to, reflect.TypeFor[T]().Name())
	conversions.lock.Unlock()
	return nil
}

// RegisterVariantConverters registers all type conversions that are supported by the library.
//
// You must call this function before calling out to IDispatch.
func RegisterVariantConverters() {
	conversions.lock.Lock()
	defer conversions.lock.Unlock()
	conversions.from[VT_NULL] = VariantToNull
	conversions.from[VT_EMPTY] = VariantToEmpty

	conversions.from[VT_HRESULT] = VariantToHandle
	conversions.from[VT_ERROR] = VariantToError
	conversions.to[reflect.TypeFor[windows.Handle]().Name()] = HResultToVariant

	conversions.from[VT_UNKNOWN] = VariantToComObject[*IUnknown]
	conversions.from[VT_UNKNOWN|VT_BYREF] = VariantToComObject[**IUnknown]
	conversions.to[reflect.TypeFor[*IsIUnknown]().Name()] = IUnknownToVariant
	conversions.to[reflect.TypeFor[*IUnknown]().Name()] = IUnknownToVariant

	conversions.from[VT_DISPATCH] = VariantToComObject[*IDispatch]
	conversions.from[VT_DISPATCH|VT_BYREF] = VariantToComObject[**IDispatch]
	conversions.to[reflect.TypeFor[*IsIDispatch]().Name()] = IDispatchToVariant
	conversions.to[reflect.TypeFor[*IDispatch]().Name()] = IDispatchToVariant

	conversions.from[VT_BOOL] = VariantToBool
	conversions.from[VT_BOOL|VT_BYREF] = VariantToBoolPtr
	conversions.to[reflect.TypeFor[bool]().Name()] = BoolToVariant
	conversions.to[reflect.TypeFor[*bool]().Name()] = BoolPtrToVariant

	conversions.from[VT_VARIANT] = VariantToGoVariant
	conversions.to[reflect.TypeFor[*VARIANT]().Name()] = GoVariantToVariant

	conversions.from[VT_CY] = VariantToInt64
	conversions.to[reflect.TypeFor[Currency]().Name()] = CurrencyToVariant

	conversions.from[VT_DATE] = VariantToTime
	conversions.from[VT_FILETIME] = VariantToFileTime
	conversions.to[reflect.TypeFor[time.Time]().Name()] = TimeToVariant

	conversions.from[VT_I1] = VariantToInt8
	conversions.from[VT_I1|VT_BYREF] = VariantToInt8Ptr
	conversions.to[reflect.TypeFor[int8]().Name()] = Int8ToVariant
	conversions.to[reflect.TypeFor[*int8]().Name()] = Int8PtrToVariant

	conversions.from[VT_UI1] = VariantToUInt8
	conversions.from[VT_UI1|VT_BYREF] = VariantToUInt8Ptr
	conversions.to[reflect.TypeFor[uint8]().Name()] = UInt8ToVariant
	conversions.to[reflect.TypeFor[*uint8]().Name()] = UInt8PtrToVariant

	conversions.from[VT_I2] = VariantToInt16
	conversions.from[VT_I2|VT_BYREF] = VariantToInt16Ptr
	conversions.to[reflect.TypeFor[int16]().Name()] = Int16ToVariant
	conversions.to[reflect.TypeFor[*int16]().Name()] = Int16PtrToVariant

	conversions.from[VT_UI2] = VariantToUInt16
	conversions.from[VT_UI2|VT_BYREF] = VariantToUInt16Ptr
	conversions.to[reflect.TypeFor[uint16]().Name()] = UInt16ToVariant
	conversions.to[reflect.TypeFor[*uint16]().Name()] = UInt16PtrToVariant

	conversions.from[VT_I4] = VariantToInt32
	conversions.from[VT_INT] = VariantToInt32
	conversions.from[VT_I4|VT_BYREF] = VariantToInt32Ptr
	conversions.from[VT_INT|VT_BYREF] = VariantToInt32Ptr
	conversions.to[reflect.TypeFor[int32]().Name()] = Int32ToVariant
	conversions.to[reflect.TypeFor[*int32]().Name()] = Int32PtrToVariant

	conversions.from[VT_UI4] = VariantToUInt32
	conversions.from[VT_UINT] = VariantToUInt
	conversions.from[VT_UI4|VT_BYREF] = VariantToUInt32Ptr
	conversions.from[VT_UINT|VT_BYREF] = VariantToUInt32Ptr
	conversions.to[reflect.TypeFor[uint32]().Name()] = UInt32ToVariant
	conversions.to[reflect.TypeFor[*uint32]().Name()] = UInt32PtrToVariant

	conversions.from[VT_I8] = VariantToInt64
	conversions.from[VT_I8|VT_BYREF] = VariantToInt64Ptr
	conversions.to[reflect.TypeFor[int64]().Name()] = Int64ToVariant
	conversions.to[reflect.TypeFor[*int64]().Name()] = Int64PtrToVariant

	conversions.from[VT_UI8] = VariantToUInt64
	conversions.from[VT_UI8|VT_BYREF] = VariantToUInt64Ptr
	conversions.to[reflect.TypeFor[uint64]().Name()] = UInt64ToVariant
	conversions.to[reflect.TypeFor[*uint64]().Name()] = UInt64PtrToVariant

	conversions.from[VT_INT_PTR] = VariantToIntPtr
	conversions.to[reflect.TypeFor[int]().Name()] = IntPtrToVariant // This depends on the platform, it should be either 4 bytes on 32-bit and 8 bytes on 64-bit
	conversions.to[reflect.TypeFor[*int]().Name()] = IntPtrToVariant

	conversions.from[VT_UINT_PTR] = VariantToUIntPtr
	conversions.from[VT_PTR] = VariantToUIntPtr
	conversions.to[reflect.TypeFor[*uint]().Name()] = UIntPtrToVariant // This depends on the platform, it should be either 4 bytes on 32-bit and 8 bytes on 64-bit
	conversions.to[reflect.TypeFor[uint]().Name()] = UIntPtrToVariant
	conversions.to[reflect.TypeFor[uintptr]().Name()] = UIntPtrToVariant

	conversions.from[VT_R4] = VariantToFloat32
	conversions.to[reflect.TypeFor[float32]().Name()] = Float32ToVariant
	conversions.from[VT_R4|VT_BYREF] = VariantToFloat32Ptr
	conversions.to[reflect.TypeFor[*float32]().Name()] = Float32PtrToVariant

	conversions.from[VT_R8] = VariantToFloat64
	conversions.to[reflect.TypeFor[float64]().Name()] = Float64ToVariant
	conversions.from[VT_R8|VT_BYREF] = VariantToFloat64Ptr
	conversions.to[reflect.TypeFor[*float64]().Name()] = Float64PtrToVariant

	conversions.from[VT_BSTR] = VariantBStrToString
	conversions.from[VT_BSTR|VT_BYREF] = VariantBStrToString
	conversions.to[reflect.TypeFor[string]().Name()] = StringToBStrVariant
}

// MakeNullVariant is for creating an empty VARIANT with a null value.
//
// Nil does not allow for automatic conversion through the map.
func MakeNullVariant() *VARIANT {
	return &VARIANT{VT: VT_NULL}
}

// VariantToNull converts VARIANT to nil.
func VariantToNull(variant *VARIANT) any {
	return nil
}

// MakeEmptyVariant is for creating an empty VARIANT.
//
// Empty does not allow for automatic conversion through the map.
func MakeEmptyVariant() *VARIANT {
	return &VARIANT{VT: VT_EMPTY}
}

// VariantToEmpty converts VARIANT to an empty string.
func VariantToEmpty(variant *VARIANT) any {
	return ""
}

// HResultToVariant will take a windows.Handle and convert it to VT_HRESULT VARIANT.
//
// This should be automatically converted when you invoke using windows.Handle.
func HResultToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_HRESULT, Val: int64(i.(windows.Handle))}
}

// ErrorToVariant will take a windows.Handle and convert it to VT_ERROR VARIANT.
//
// If you need to send VT_ERROR, then you will want to call this manually. windows.Handle will not automatically be
// converted to this type.
func ErrorToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_ERROR, Val: int64(i.(windows.Handle))}
}

func VariantToError(variant *VARIANT) any {
	return windows.Handle(uintptr(variant.Val))
}

func VariantToHandle(variant *VARIANT) any {
	return windows.Handle(uintptr(variant.Val))
}

func HandleToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_HRESULT, Val: int64(i.(windows.Handle))}
}

func IUnknownToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_UNKNOWN, Val: int64(uintptr(unsafe.Pointer(i.(*IUnknown))))}
}

func IDispatchToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_DISPATCH, Val: int64(uintptr(unsafe.Pointer(i.(*IDispatch))))}
}

func VariantToComObject[T IsIUnknown](variant *VARIANT) any {
	return (*T)(unsafe.Pointer(uintptr(variant.Val)))
}

// VariantToGoVariant converts a *VARIANT to a *VARIANT of VT_VARIANT VT type.
//
// There is no automatic conversion for this, you must call this manually.
func VariantToGoVariant(variant *VARIANT) *VARIANT {
	return (*VARIANT)(unsafe.Pointer(uintptr(variant.Val)))
}

// GoVariantToVariant converts a *VARIANT to *VARIANT.
//
// There is no automatic conversion for this, you must call this manually.
func GoVariantToVariant(variant *VARIANT) *VARIANT {
	return &VARIANT{VT: VT_VARIANT | VT_BYREF, Val: int64(uintptr(unsafe.Pointer(variant)))}
}

// ClassIdToVariant converts a windows.GUID or *windows.GUID to *VARIANT.
//
// There is no automatic conversion for this, you must call this manually.
func ClassIdToVariant(i any) any {
	switch i.(type) {
	case windows.GUID:
		return &VARIANT{VT: VT_CLSID, Val: int64(uintptr(unsafe.Pointer(i.(windows.GUID))))}
	case *windows.GUID:
		return &VARIANT{VT: VT_CLSID | VT_BYREF, Val: int64(uintptr(unsafe.Pointer(i.(*windows.GUID))))}
	}
}

// VariantToClassId converts *VARIANT to windows.GUID
func VariantToClassId(variant *VARIANT) any {
	if variant.VT == VT_CLSID|VT_BYREF {
		return (*windows.GUID)(unsafe.Pointer(uintptr(variant.Val)))
	}
	if variant.VT == VT_CLSID {
		return (windows.GUID)(unsafe.Pointer(uintptr(variant.Val)))
	}
}

// VoidToVariant converts a C void pointer to *VARIANT.
//
// There is no automatic conversion for this, you must call this manually.
func VoidToVariant(i any) any {
	return &VARIANT{VT: VT_VOID, Val: int64(uintptr(unsafe.Pointer(i)))}
}

// VariantToVoid converts *VARIANT to C void pointer.
//
// There is no automatic conversion for this, you must call this manually.
func VariantToVoid[T any](variant *VARIANT) any {
	return (*T)(unsafe.Pointer(uintptr(variant.Val)))
}

// MakeCurrencyVariant will create a currency VARIANT type.
//
// You have to call this manually, there is not an automatic conversion
func MakeCurrencyVariant(i int64) *VARIANT {
	return &VARIANT{VT: VT_CY, Val: i}
}

// CurrencyToVariant will create a VT_CY VT type *VARIANT.
//
// You should take cents and multiply by 10,000 and convert to int64 type.
func CurrencyToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_CY, Val: i.(int64)}
}

// VariantToFileTime converts *VARIANT to time.Time.
func VariantToFileTime(variant *VARIANT) any {
	nanoseconds := variant.Val
	return FileTimeEpoch.Add(time.Microsecond * int64(nanoseconds/10))
}

// FileTimeToVariant returns the VT_FILETIME *VARIANT of int64.
//
// The epoch for FileTime is the year of 1601 and time.Duration has a resolution of 290 years. It is not possible to Sub
// two time.Time to get the Val for VT_FILETIME. Therefore, you have to calculate this yourself.
// FILETIME is an int64 with internal of 100 nanoseconds from midnight January 1st, 1601.
//
// You have to call this manually, there is not an automatic conversion
func FileTimeToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_FILETIME, Val: i.(int64)}
}

// VariantToTime converts *VARIANT to time.Time.
func VariantToTime(variant *VARIANT) any {
	timestamp := (float64)(unsafe.Pointer(uintptr(variant.Val)))
	days := int64(timestamp)
	remainder := timestamp - float64(days)
	hours := remainder / float64(DateSingleHour)
	remainder = remainder - hours
	minutes := remainder / float64(DateSingleMinute)
	remainder = remainder - minutes
	seconds := remainder / float64(DateSingleSecond)
	remainder = remainder - seconds
	milliseconds := remainder / float64(DateSingleMilliSecond)

	date := DateEpoch.Add(days * 24 * time.Hour)
	date = date.Add(hours * time.Hour)
	date = date.Add(minutes * time.Minute)
	date = date.Add(seconds * time.Second)
	date = date.Add(milliseconds * time.Millisecond)
	return date
}

// TimeToVariant returns the VT_DATE VARIANT of time.Time.
//
// This subs the time passed with DateEpoch, which is December 30th, 1899. time.Duration has a limit of 290 years and
// is currently limited with Go 1.23 to dates up to 2189. Everyone alive currently does not need to worry about system
// time, but if you are passing other dates, then please be aware that you may want to handle arbitrary dates far into
// the future. It is assumed that COM supports dates before 1899, but you might want to passed strings if outside these
// ranges.
func TimeToVariant(i any) *VARIANT {
	date := i.(time.Time)
	duration := date.Sub(DateEpoch)
	rawHours := duration.Hours()
	days := int64(rawHours / 24)
	hours := int64(rawHours - (days * 24))
	minutes := duration.Minutes()
	seconds := duration.Seconds()
	milliseconds := duration.Milliseconds()

	winDate := float64(days) + float64(hours*DateSingleHour) + float64(minutes*DateSingleMinute) + float64(seconds*DateSingleSecond) + float64(milliseconds*time.Millisecond)

	return &VARIANT{VT: VT_DATE, Val: int64(uintptr(unsafe.Pointer(winDate)))}
}

func VariantToBool(variant *VARIANT) any {
	return variant.Val != int64(VariantTypeFalse)
}

// BoolToVariant converts go type to Variant VT_BOOL type
func BoolToVariant(i any) *VARIANT {
	val := i.(bool)
	if val {
		return &VARIANT{VT: VT_BOOL, Val: int64(VariantTypeTrue)}
	}
	return &VARIANT{VT: VT_BOOL, Val: int64(VariantTypeFalse)}
}

func VariantToBoolPtr(variant *VARIANT) any {
	val := (*int16)(unsafe.Pointer(uintptr(variant.Val)))
	return val != VariantTypeFalse
}

// BoolToVariant converts go type to Variant VT_BOOL type
func BoolPtrToVariant(i any) *VARIANT {
	val := i.(*bool)
	var ptr int16
	if val {
		ptr = int16(VariantTypeTrue)
	} else {
		ptr = int16(VariantTypeFalse)
	}
	return &VARIANT{VT: VT_BOOL, Val: int64(uintptr(unsafe.Pointer(ptr)))}
}

// Int8ToVariant converts go type to Variant VT_I1 type
func Int8ToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_I1, Val: int64(i.(int8))}
}

func VariantToInt8(variant *VARIANT) any {
	return int8(variant.Val)
}

// Int8PtrToVariant converts go type to Variant VT_I1|VT_BYREF type
func Int8PtrToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_I1 | VT_BYREF, Val: int64(uintptr(unsafe.Pointer(i.(*int8))))}
}

func VariantToInt8Ptr(variant *VARIANT) any {
	return (*int8)(unsafe.Pointer(uintptr(variant.Val)))
}

func UInt8ToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_UI1, Val: int64(i.(uint8))}
}

func VariantToUInt8(variant *VARIANT) any {
	return uint8(variant.Val)
}

// UInt8PtrToVariant converts go type to Variant VT_UI1|VT_BYREF type
func UInt8PtrToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_UI1 | VT_BYREF, Val: int64(uintptr(unsafe.Pointer(i.(*uint8))))}
}

func VariantToUInt8Ptr(variant *VARIANT) any {
	return (*uint8)(unsafe.Pointer(uintptr(variant.Val)))
}

func Int16ToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_I2, Val: int64(i.(int16))}
}

func VariantToInt16(variant *VARIANT) any {
	return int16(variant.Val)
}

// Int16PtrToVariant converts go type to Variant VT_I2|VT_BYREF type
func Int16PtrToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_I2 | VT_BYREF, Val: int64(uintptr(unsafe.Pointer(i.(*int16))))}
}

func VariantToInt16Ptr(variant *VARIANT) any {
	return (*int16)(unsafe.Pointer(uintptr(variant.Val)))
}

func UInt16ToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_UI2, Val: int64(i.(uint16))}
}

func VariantToUInt16(variant *VARIANT) any {
	return uint16(variant.Val)
}

// UInt16PtrToVariant converts go type to Variant VT_UI2|VT_BYREF type
func UInt16PtrToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_UI2 | VT_BYREF, Val: int64(uintptr(unsafe.Pointer(i.(*uint16))))}
}

func VariantToUInt16Ptr(variant *VARIANT) any {
	return (*uint16)(unsafe.Pointer(uintptr(variant.Val)))
}

func Int32ToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_I4, Val: int64(i.(int32))}
}

func VariantToInt32(variant *VARIANT) any {
	return int32(variant.Val)
}

// Int32PtrToVariant converts go type to Variant VT_I4|VT_BYREF type
func Int32PtrToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_I4 | VT_BYREF, Val: int64(uintptr(unsafe.Pointer(i.(*int32))))}
}

func VariantToInt32Ptr(variant *VARIANT) any {
	return (*int32)(unsafe.Pointer(uintptr(variant.Val)))
}

func UInt32ToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_UI4, Val: int64(i.(uint32))}
}

func VariantToUInt32(variant *VARIANT) any {
	return uint32(variant.Val)
}

// UInt32PtrToVariant converts go type to Variant VT_UI4|VT_BYREF type
func UInt32PtrToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_UI4 | VT_BYREF, Val: int64(uintptr(unsafe.Pointer(i.(*uint16))))}
}

func VariantToUInt32Ptr(variant *VARIANT) any {
	return (*uint32)(unsafe.Pointer(uintptr(variant.Val)))
}

func Int64ToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_I8, Val: i.(int64)}
}

func VariantToInt64(variant *VARIANT) any {
	return int64(variant.Val)
}

// Int64PtrToVariant converts go type to Variant VT_I8|VT_BYREF type
func Int64PtrToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_I8 | VT_BYREF, Val: int64(uintptr(unsafe.Pointer(i.(*int64))))}
}

func VariantToInt64Ptr(variant *VARIANT) any {
	return (*int64)(unsafe.Pointer(uintptr(variant.Val)))
}

func UInt64ToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_UI8, Val: int64(i.(uint64))}
}

func VariantToUInt64(variant *VARIANT) any {
	return uint64(variant.Val)
}

// UInt64PtrToVariant converts go type to Variant VT_UI8|VT_BYREF type
func UInt64PtrToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_UI8 | VT_BYREF, Val: int64(uintptr(unsafe.Pointer(i.(*uint64))))}
}

func VariantToUInt64Ptr(variant *VARIANT) any {
	return (*uint64)(unsafe.Pointer(uintptr(variant.Val)))
}

func IntToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_INT, Val: int64(i.(int))}
}

func VariantToInt(variant *VARIANT) any {
	return int(variant.Val)
}

func UIntToVariant(i any) *VARIANT {
	return &VARIANT{VT: VT_UINT, Val: int64(i.(uint))}
}

func VariantToUInt(variant *VARIANT) any {
	return uint(variant.Val)
}

func VariantToFloat32(variant *VARIANT) any {
	return (float32)(unsafe.Pointer(uintptr(variant.Val)))
}

func Float32ToVariant(i any) *VARIANT {
	number := i.(float32)
	address := &number
	return &VARIANT{VT: VT_R4, Val: int64(uintptr(unsafe.Pointer(address)))}
}

func VariantToFloat64(variant *VARIANT) any {
	return (float64)(unsafe.Pointer(uintptr(variant.Val)))
}

func Float64ToVariant(i any) *VARIANT {
	number := i.(float64)
	address := &number
	return &VARIANT{VT: VT_R4, Val: int64(uintptr(unsafe.Pointer(address)))}
}

func VariantBStrToString(variant *VARIANT) any {
	return windows.UTF16PtrToString((*uint16)(unsafe.Pointer(uintptr(variant.Val))))
}

func StringToBStrVariant(i any) *VARIANT {
	str := i.(string)
	ptr := windows.UTF16PtrFromString(str)
	return &VARIANT{VT: VT_BSTR, Val: int64(uintptr(unsafe.Pointer(&ptr)))}
}
