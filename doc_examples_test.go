//go:build windows

package ole

import (
	"fmt"
	"unsafe"

	"golang.org/x/sys/windows"
)

// ---- Initialization ----

func ExampleInitialize() {
	result, err := Initialize(Multithreaded)
	if err != nil {
		return
	}
	if result == SuccessfullyInitialized || result == AlreadyInitialized {
		defer Uninitialize()
	}
	fmt.Println(result == SuccessfullyInitialized || result == AlreadyInitialized)
	// Output:
	// true
}

func ExampleInitializeMultithreaded() {
	result, err := InitializeMultithreaded()
	if err != nil {
		return
	}
	if result == SuccessfullyInitialized || result == AlreadyInitialized {
		defer Uninitialize()
	}
	fmt.Println(result == SuccessfullyInitialized || result == AlreadyInitialized)
	// Output:
	// true
}

func ExampleInitializeApartmentThreaded() {
	result, err := InitializeApartmentThreaded()
	if err != nil {
		return
	}
	if result == SuccessfullyInitialized || result == AlreadyInitialized {
		defer Uninitialize()
	}
}

// ---- Class/Interface ID Lookup ----

func ExampleClassIdFromString() {
	classID, err := ClassIdFromString("{00000000-0000-0000-C000-000000000046}")
	if err != nil {
		return
	}
	fmt.Println(classID == IID_IUnknown)
	// Output:
	// true
}

func ExampleClassIdFromString_programId() {
	// ClassIdFromString also accepts Program IDs like "Excel.Application".
	// It tries ClassIdFromProgramId first, then falls back to ClassIdFromGuidString.
	_, err := ClassIdFromString("Excel.Application")
	if err != nil {
		// Excel may not be installed; this is expected in some environments.
		return
	}
}

func ExampleClassIdFromString_guidFormat() {
	classID, err := ClassIdFromString("{00020400-0000-0000-C000-000000000046}")
	if err != nil {
		return
	}
	fmt.Println(classID == IID_IDispatch)
	// Output:
	// true
}

func ExampleClassIdFromGuidString() {
	classID, err := ClassIdFromGuidString("{00000000-0000-0000-C000-000000000046}")
	if err != nil {
		return
	}
	fmt.Println(classID == IID_IUnknown)
	// Output:
	// true
}

func ExampleClassIdToString() {
	text, err := ClassIdToString(IID_IUnknown)
	if err != nil {
		return
	}
	fmt.Println(text)
	// Output:
	// {00000000-0000-0000-C000-000000000046}
}

func ExampleClassIdToString_dispatch() {
	text, err := ClassIdToString(IID_IDispatch)
	if err != nil {
		return
	}
	fmt.Println(text)
	// Output:
	// {00020400-0000-0000-C000-000000000046}
}

func ExampleClassIdToString_roundtrip() {
	original := IID_IConnectionPointContainer
	text, err := ClassIdToString(original)
	if err != nil {
		return
	}
	recovered, err := ClassIdFromGuidString(text)
	if err != nil {
		return
	}
	fmt.Println(original == recovered)
	// Output:
	// true
}

func ExampleInterfaceIdToString() {
	text, err := InterfaceIdToString(IID_IDispatch)
	if err != nil {
		return
	}
	fmt.Println(text)
	// Output:
	// {00020400-0000-0000-C000-000000000046}
}

func ExampleInterfaceIdFromString() {
	iid, err := InterfaceIdFromString("{00020400-0000-0000-C000-000000000046}")
	if err != nil {
		return
	}
	fmt.Println(iid == IID_IDispatch)
	// Output:
	// true
}

func ExampleInterfaceIdFromString_unknown() {
	iid, err := InterfaceIdFromString("{00000000-0000-0000-C000-000000000046}")
	if err != nil {
		return
	}
	fmt.Println(iid == IID_IUnknown)
	// Output:
	// true
}

// ---- String Functions ----

func ExampleSysAllocString() {
	bstr := SysAllocString("hello")
	defer SysFreeString(bstr)
	fmt.Println(bstr != nil)
	// Output:
	// true
}

func ExampleSysAllocString_empty() {
	bstr := SysAllocString("")
	defer SysFreeString(bstr)
	fmt.Println(bstr != nil)
	// Output:
	// true
}

func ExampleSysAllocString_unicode() {
	bstr := SysAllocString("こんにちは")
	defer SysFreeString(bstr)
	fmt.Println(bstr != nil)
	// Output:
	// true
}

func ExampleSysAllocStringLen() {
	bstr := SysAllocStringLen("hello")
	defer SysFreeString(bstr)
	fmt.Println(bstr != nil)
	// Output:
	// true
}

func ExampleSysAllocStringLen_length() {
	bstr := SysAllocStringLen("hello")
	defer SysFreeString(bstr)
	length := SysStringLen((*int16)(unsafe.Pointer(bstr)))
	fmt.Println(length)
	// Output:
	// 5
}

func ExampleSysAllocStringLen_unicode() {
	bstr := SysAllocStringLen("日本語")
	defer SysFreeString(bstr)
	length := SysStringLen((*int16)(unsafe.Pointer(bstr)))
	fmt.Println(length)
	// Output:
	// 3
}

func ExampleSysFreeString() {
	bstr := SysAllocString("temporary")
	err := SysFreeString(bstr)
	fmt.Println(err)
	// Output:
	// <nil>
}

func ExampleSysStringLen() {
	bstr := SysAllocStringLen("test")
	defer SysFreeString(bstr)
	length := SysStringLen((*int16)(unsafe.Pointer(bstr)))
	fmt.Println(length)
	// Output:
	// 4
}

func ExampleNewHString() {
	value, err := NewHString("go-ole")
	if err != nil {
		return
	}
	defer DeleteHString(value)

	fmt.Println(value.String())
	// Output:
	// go-ole
}

func ExampleNewHString_unicode() {
	value, err := NewHString("日本語テスト")
	if err != nil {
		return
	}
	defer DeleteHString(value)

	fmt.Println(value.String())
	// Output:
	// 日本語テスト
}

func ExampleNewHString_empty() {
	value, err := NewHString("")
	if err != nil {
		return
	}
	defer DeleteHString(value)

	fmt.Println(value.String())
	// Output:
	//
}

func ExampleDeleteHString() {
	value, err := NewHString("temporary")
	if err != nil {
		return
	}
	err = DeleteHString(value)
	fmt.Println(err)
	// Output:
	// <nil>
}

// ---- Variant Creation/Conversion ----

func ExampleRegisterVariantConverters() {
	RegisterVariantConverters()
	// After registration, WrapVariant and UnwrapVariant work with standard types.
	variant, err := WrapVariant(int32(42))
	if err != nil {
		return
	}
	fmt.Println(variant.VT == VT_I4)
	// Output:
	// true
}

func ExampleVariantInit() {
	var v VARIANT
	err := VariantInit(&v)
	if err != nil {
		return
	}
	fmt.Println(v.VT)
	// Output:
	// 0
}

func ExampleVariantClear() {
	v := BoolToVariant(true)
	err := VariantClear(v)
	if err != nil {
		return
	}
	fmt.Println(v.VT)
	// Output:
	// 0
}

func ExampleVARIANT_Clear() {
	v := Int32ToVariant(int32(100))
	err := v.Clear()
	if err != nil {
		return
	}
	fmt.Println(v.VT)
	// Output:
	// 0
}

func ExampleWrapVariant() {
	variant, err := WrapVariant(int32(42))
	if err != nil {
		return
	}
	defer variant.Clear()

	fmt.Println(variant.VT, variant.Val)
	// Output:
	// 3 42
}

func ExampleWrapVariant_bool() {
	RegisterVariantConverters()
	variant, err := WrapVariant(true)
	if err != nil {
		return
	}
	fmt.Println(variant.VT == VT_BOOL)
	// Output:
	// true
}

func ExampleWrapVariant_string() {
	RegisterVariantConverters()
	variant, err := WrapVariant("hello")
	if err != nil {
		return
	}
	defer variant.Clear()
	fmt.Println(variant.VT == VT_BSTR)
	// Output:
	// true
}

func ExampleWrapVariant_float64() {
	RegisterVariantConverters()
	variant, err := WrapVariant(float64(3.14))
	if err != nil {
		return
	}
	fmt.Println(variant.VT == VT_R8)
	// Output:
	// true
}

func ExampleWrapVariant_int64() {
	RegisterVariantConverters()
	variant, err := WrapVariant(int64(9999))
	if err != nil {
		return
	}
	fmt.Println(variant.VT == VT_I8, variant.Val)
	// Output:
	// true 9999
}

func ExampleUnwrapVariant_int32() {
	RegisterVariantConverters()
	variant := Int32ToVariant(int32(42))
	value := UnwrapVariant[int32](variant)
	fmt.Println(value)
	// Output:
	// 42
}

func ExampleUnwrapVariant_bool() {
	RegisterVariantConverters()
	variant := BoolToVariant(true)
	value := UnwrapVariant[bool](variant)
	fmt.Println(value)
	// Output:
	// true
}

func ExampleUnwrapVariant_string() {
	RegisterVariantConverters()
	variant := StringToBStrVariant("hello world")
	defer variant.Clear()
	value := UnwrapVariant[string](variant)
	fmt.Println(value)
	// Output:
	// hello world
}

func ExampleUnwrapVariant_float64() {
	RegisterVariantConverters()
	variant := Float64ToVariant(float64(3.14))
	value := UnwrapVariant[float64](variant)
	fmt.Println(value)
	// Output:
	// 3.14
}

func ExampleUnwrapVariant_int64() {
	RegisterVariantConverters()
	variant := Int64ToVariant(int64(1000000))
	value := UnwrapVariant[int64](variant)
	fmt.Println(value)
	// Output:
	// 1000000
}

func ExampleBoolToVariant() {
	variant := BoolToVariant(true)
	fmt.Println(variant.VT == VT_BOOL)
	fmt.Println(VariantToBool(variant))
	// Output:
	// true
	// true
}

func ExampleBoolToVariant_false() {
	variant := BoolToVariant(false)
	fmt.Println(variant.VT == VT_BOOL)
	fmt.Println(VariantToBool(variant))
	// Output:
	// true
	// false
}

func ExampleBoolToVariant_roundtrip() {
	RegisterVariantConverters()
	variant := BoolToVariant(true)
	value := UnwrapVariant[bool](variant)
	fmt.Println(value)
	// Output:
	// true
}

func ExampleVariantToBool() {
	variant := BoolToVariant(true)
	value := VariantToBool(variant)
	fmt.Println(value)
	// Output:
	// true
}

func ExampleVariantToBool_false() {
	variant := BoolToVariant(false)
	value := VariantToBool(variant)
	fmt.Println(value)
	// Output:
	// false
}

func ExampleInt8ToVariant() {
	variant := Int8ToVariant(int8(127))
	fmt.Println(variant.VT == VT_I1, variant.Val)
	// Output:
	// true 127
}

func ExampleVariantToInt8() {
	variant := Int8ToVariant(int8(-42))
	value := VariantToInt8(variant)
	fmt.Println(value)
	// Output:
	// -42
}

func ExampleUInt8ToVariant() {
	variant := UInt8ToVariant(uint8(255))
	fmt.Println(variant.VT == VT_UI1, variant.Val)
	// Output:
	// true 255
}

func ExampleVariantToUInt8() {
	variant := UInt8ToVariant(uint8(200))
	value := VariantToUInt8(variant)
	fmt.Println(value)
	// Output:
	// 200
}

func ExampleInt16ToVariant() {
	variant := Int16ToVariant(int16(32000))
	fmt.Println(variant.VT == VT_I2, variant.Val)
	// Output:
	// true 32000
}

func ExampleInt16ToVariant_negative() {
	variant := Int16ToVariant(int16(-100))
	fmt.Println(variant.VT == VT_I2)
	fmt.Println(VariantToInt16(variant))
	// Output:
	// true
	// -100
}

func ExampleInt16ToVariant_roundtrip() {
	RegisterVariantConverters()
	variant := Int16ToVariant(int16(1234))
	value := UnwrapVariant[int16](variant)
	fmt.Println(value)
	// Output:
	// 1234
}

func ExampleVariantToInt16() {
	variant := Int16ToVariant(int16(-500))
	value := VariantToInt16(variant)
	fmt.Println(value)
	// Output:
	// -500
}

func ExampleUInt16ToVariant() {
	variant := UInt16ToVariant(uint16(65535))
	fmt.Println(variant.VT == VT_UI2, variant.Val)
	// Output:
	// true 65535
}

func ExampleVariantToUInt16() {
	variant := UInt16ToVariant(uint16(1000))
	value := VariantToUInt16(variant)
	fmt.Println(value)
	// Output:
	// 1000
}

func ExampleInt32ToVariant() {
	variant := Int32ToVariant(int32(42))
	fmt.Println(variant.VT == VT_I4, variant.Val)
	// Output:
	// true 42
}

func ExampleInt32ToVariant_negative() {
	variant := Int32ToVariant(int32(-1))
	fmt.Println(variant.VT == VT_I4)
	fmt.Println(VariantToInt32(variant))
	// Output:
	// true
	// -1
}

func ExampleInt32ToVariant_roundtrip() {
	RegisterVariantConverters()
	variant := Int32ToVariant(int32(99999))
	value := UnwrapVariant[int32](variant)
	fmt.Println(value)
	// Output:
	// 99999
}

func ExampleInt32ToVariant_zero() {
	variant := Int32ToVariant(int32(0))
	fmt.Println(variant.VT == VT_I4, variant.Val)
	// Output:
	// true 0
}

func ExampleInt32ToVariant_maxValue() {
	variant := Int32ToVariant(int32(2147483647))
	fmt.Println(VariantToInt32(variant))
	// Output:
	// 2147483647
}

func ExampleVariantToInt32() {
	variant := Int32ToVariant(int32(12345))
	value := VariantToInt32(variant)
	fmt.Println(value)
	// Output:
	// 12345
}

func ExampleUInt32ToVariant() {
	variant := UInt32ToVariant(uint32(4294967295))
	fmt.Println(variant.VT == VT_UI4, variant.Val)
	// Output:
	// true 4294967295
}

func ExampleVariantToUInt32() {
	variant := UInt32ToVariant(uint32(100))
	value := VariantToUInt32(variant)
	fmt.Println(value)
	// Output:
	// 100
}

func ExampleInt64ToVariant() {
	variant := Int64ToVariant(int64(9999999999))
	fmt.Println(variant.VT == VT_I8, variant.Val)
	// Output:
	// true 9999999999
}

func ExampleInt64ToVariant_negative() {
	variant := Int64ToVariant(int64(-1))
	fmt.Println(VariantToInt64(variant))
	// Output:
	// -1
}

func ExampleInt64ToVariant_roundtrip() {
	RegisterVariantConverters()
	variant := Int64ToVariant(int64(123456789))
	value := UnwrapVariant[int64](variant)
	fmt.Println(value)
	// Output:
	// 123456789
}

func ExampleVariantToInt64() {
	variant := Int64ToVariant(int64(1000000))
	value := VariantToInt64(variant)
	fmt.Println(value)
	// Output:
	// 1000000
}

func ExampleUInt64ToVariant() {
	variant := UInt64ToVariant(uint64(18446744073709551615))
	fmt.Println(variant.VT == VT_UI8)
	// Output:
	// true
}

func ExampleVariantToUInt64() {
	variant := UInt64ToVariant(uint64(999))
	value := VariantToUInt64(variant)
	fmt.Println(value)
	// Output:
	// 999
}

func ExampleIntToVariant() {
	variant := IntToVariant(42)
	fmt.Println(variant.VT == VT_INT, variant.Val)
	// Output:
	// true 42
}

func ExampleUIntToVariant() {
	variant := UIntToVariant(uint(100))
	fmt.Println(variant.VT == VT_UINT, variant.Val)
	// Output:
	// true 100
}

func ExampleFloat32ToVariant() {
	variant := Float32ToVariant(float32(3.14))
	fmt.Println(variant.VT == VT_R4)
	value := VariantToFloat32(variant).(float32)
	fmt.Printf("%.2f\n", value)
	// Output:
	// true
	// 3.14
}

func ExampleFloat32ToVariant_roundtrip() {
	RegisterVariantConverters()
	variant := Float32ToVariant(float32(1.5))
	value := UnwrapVariant[float32](variant)
	fmt.Println(value)
	// Output:
	// 1.5
}

func ExampleFloat32ToVariant_zero() {
	variant := Float32ToVariant(float32(0))
	fmt.Println(variant.VT == VT_R4)
	// Output:
	// true
}

func ExampleVariantToFloat32() {
	variant := Float32ToVariant(float32(2.5))
	value := VariantToFloat32(variant)
	fmt.Println(value)
	// Output:
	// 2.5
}

func ExampleFloat64ToVariant() {
	variant := Float64ToVariant(float64(3.14159))
	fmt.Println(variant.VT == VT_R8)
	value := VariantToFloat64(variant).(float64)
	fmt.Println(value)
	// Output:
	// true
	// 3.14159
}

func ExampleFloat64ToVariant_negative() {
	variant := Float64ToVariant(float64(-273.15))
	value := VariantToFloat64(variant).(float64)
	fmt.Println(value)
	// Output:
	// -273.15
}

func ExampleFloat64ToVariant_roundtrip() {
	RegisterVariantConverters()
	variant := Float64ToVariant(float64(2.71828))
	value := UnwrapVariant[float64](variant)
	fmt.Println(value)
	// Output:
	// 2.71828
}

func ExampleFloat64ToVariant_zero() {
	variant := Float64ToVariant(float64(0))
	fmt.Println(variant.VT == VT_R8, variant.Val)
	// Output:
	// true 0
}

func ExampleFloat64ToVariant_large() {
	variant := Float64ToVariant(float64(1e15))
	value := VariantToFloat64(variant).(float64)
	fmt.Println(value)
	// Output:
	// 1e+15
}

func ExampleVariantToFloat64() {
	variant := Float64ToVariant(float64(99.99))
	value := VariantToFloat64(variant)
	fmt.Println(value)
	// Output:
	// 99.99
}

func ExampleStringToBStrVariant() {
	variant := StringToBStrVariant("hello")
	defer variant.Clear()
	fmt.Println(variant.VT == VT_BSTR)
	// Output:
	// true
}

func ExampleStringToBStrVariant_unicode() {
	variant := StringToBStrVariant("日本語")
	defer variant.Clear()
	fmt.Println(variant.VT == VT_BSTR)
	value := VariantBStrToString(variant)
	fmt.Println(value)
	// Output:
	// true
	// 日本語
}

func ExampleStringToBStrVariant_roundtrip() {
	RegisterVariantConverters()
	variant := StringToBStrVariant("go-ole test")
	defer variant.Clear()
	value := UnwrapVariant[string](variant)
	fmt.Println(value)
	// Output:
	// go-ole test
}

func ExampleStringToBStrVariant_empty() {
	variant := StringToBStrVariant("")
	defer variant.Clear()
	fmt.Println(variant.VT == VT_BSTR)
	value := VariantBStrToString(variant)
	fmt.Println(value)
	// Output:
	// true
	//
}

func ExampleStringToBStrVariant_special() {
	variant := StringToBStrVariant("line1\tline2")
	defer variant.Clear()
	value := VariantBStrToString(variant)
	fmt.Println(value)
	// Output:
	// line1	line2
}

func ExampleVariantBStrToString() {
	variant := StringToBStrVariant("hello world")
	defer variant.Clear()
	value := VariantBStrToString(variant)
	fmt.Println(value)
	// Output:
	// hello world
}

func ExampleMakeNullVariant() {
	v := MakeNullVariant()
	fmt.Println(v.VT == VT_NULL)
	// Output:
	// true
}

func ExampleMakeEmptyVariant() {
	v := MakeEmptyVariant()
	fmt.Println(v.VT == VT_EMPTY)
	// Output:
	// true
}

func ExampleHResultToVariant() {
	variant := HResultToVariant(windows.Handle(0))
	fmt.Println(variant.VT == VT_HRESULT, variant.Val)
	// Output:
	// true 0
}

func ExampleErrorToVariant() {
	variant := ErrorToVariant(windows.Handle(0x80004005))
	fmt.Println(variant.VT == VT_ERROR)
	// Output:
	// true
}

func ExampleMakeCurrencyVariant() {
	// Currency values are stored as int64 with 4 implied decimal places.
	// 12.3456 is stored as 123456.
	variant := MakeCurrencyVariant(123456)
	fmt.Println(variant.VT == VT_CY, variant.Val)
	// Output:
	// true 123456
}

func ExampleGoVariantToVariant() {
	inner := Int32ToVariant(int32(42))
	outer := GoVariantToVariant(inner)
	fmt.Println(outer.VT == VT_VARIANT|VT_BYREF)
	// Output:
	// true
}

func ExampleIUnknownToVariant() {
	unknown := &IUnknown{}
	variant := IUnknownToVariant(unknown)
	fmt.Println(variant.VT == VT_UNKNOWN)
	// Output:
	// true
}

func ExampleIDispatchToVariant() {
	dispatch := &IDispatch{}
	variant := IDispatchToVariant(dispatch)
	fmt.Println(variant.VT == VT_DISPATCH)
	// Output:
	// true
}

func ExampleWrapParametersWithVariant() {
	RegisterVariantConverters()
	params := WrapParametersWithVariant(int32(1), true, "hello")
	fmt.Println(len(params))
	fmt.Println(params[0].VT == VT_BSTR) // reversed order
	fmt.Println(params[1].VT == VT_BOOL)
	fmt.Println(params[2].VT == VT_I4)
	// Output:
	// 3
	// true
	// true
	// true
}

// ---- DISPPARAMS ----

func ExampleMakeDisplayParams() {
	visible := BoolToVariant(true)
	params := MakeDisplayParams(DISPATCH_PROPERTYPUT, visible)

	fmt.Println(params.cArgs, params.cNamedArgs)
	// Output:
	// 1 1
}

func ExampleMakeDisplayParams_method() {
	arg := Int32ToVariant(int32(10))
	params := MakeDisplayParams(DISPATCH_METHOD, arg)

	fmt.Println(params.cArgs, params.cNamedArgs)
	// Output:
	// 1 0
}

// ---- SafeArray ----

func ExampleFromSlice() {
	array, err := FromSlice([][]int32{{1, 2}, {3, 4}})
	if err != nil {
		return
	}
	defer array.Destroy()

	bounds, err := array.BoundsInfo()
	if err != nil {
		return
	}

	fmt.Println(len(bounds), bounds[0].Elements, bounds[1].Elements)
	// Output:
	// 2 2 2
}

func ExampleFromSlice_strings() {
	array, err := FromSlice([]string{"apple", "banana", "cherry"})
	if err != nil {
		return
	}
	defer array.Destroy()

	bounds, err := array.BoundsInfo()
	if err != nil {
		return
	}
	fmt.Println(bounds[0].Elements)
	// Output:
	// 3
}

func ExampleFromSlice_singleDimension() {
	array, err := FromSlice([]int32{10, 20, 30})
	if err != nil {
		return
	}
	defer array.Destroy()

	dims, err := array.GetDimensions()
	if err != nil {
		return
	}
	fmt.Println(dims)
	// Output:
	// 1
}

func ExampleToSlice() {
	array, err := FromSlice([]string{"red", "blue"})
	if err != nil {
		return
	}
	defer array.Destroy()

	values, err := ToSlice[[]string](array)
	if err != nil {
		return
	}

	fmt.Println(values[0], values[1])
	// Output:
	// red blue
}

func ExampleToSlice_int32() {
	array, err := FromSlice([]int32{10, 20, 30})
	if err != nil {
		return
	}
	defer array.Destroy()

	values, err := ToSlice[[]int32](array)
	if err != nil {
		return
	}
	fmt.Println(values[0], values[1], values[2])
	// Output:
	// 10 20 30
}

func ExampleToSlice_roundtrip() {
	original := []float64{1.1, 2.2, 3.3}
	array, err := FromSlice(original)
	if err != nil {
		return
	}
	defer array.Destroy()

	recovered, err := ToSlice[[]float64](array)
	if err != nil {
		return
	}
	fmt.Println(recovered[0], recovered[1], recovered[2])
	// Output:
	// 1.1 2.2 3.3
}

func ExampleMarshalSafeArray() {
	array, err := MarshalSafeArray([]int32{1, 2, 3})
	if err != nil {
		return
	}
	defer array.Destroy()

	bounds, _ := array.BoundsInfo()
	fmt.Println(bounds[0].Elements)
	// Output:
	// 3
}

func ExampleUnmarshalSafeArray() {
	array, err := FromSlice([]int32{100, 200, 300})
	if err != nil {
		return
	}
	defer array.Destroy()

	values, err := UnmarshalSafeArray[[]int32](array)
	if err != nil {
		return
	}
	fmt.Println(values[0], values[1], values[2])
	// Output:
	// 100 200 300
}

func ExampleWrapSliceAsVariant() {
	variant, err := WrapSliceAsVariant([]int32{1, 2, 3})
	if err != nil {
		return
	}
	fmt.Println(variant.VT == VT_ARRAY|VT_I4)
	// Output:
	// true
}

func ExampleSafeArray_Destroy() {
	array, err := FromSlice([]int32{1, 2, 3})
	if err != nil {
		return
	}
	err = array.Destroy()
	fmt.Println(err)
	// Output:
	// <nil>
}

func ExampleSafeArray_GetDimensions() {
	array, err := FromSlice([][]int32{{1, 2}, {3, 4}})
	if err != nil {
		return
	}
	defer array.Destroy()

	dims, err := array.GetDimensions()
	if err != nil {
		return
	}
	fmt.Println(dims)
	// Output:
	// 2
}

func ExampleSafeArray_GetElementSize() {
	array, err := FromSlice([]int32{1, 2, 3})
	if err != nil {
		return
	}
	defer array.Destroy()

	size, err := array.GetElementSize()
	if err != nil {
		return
	}
	fmt.Println(size)
	// Output:
	// 4
}

func ExampleSafeArray_BoundsInfo() {
	array, err := FromSlice([]int32{10, 20, 30, 40, 50})
	if err != nil {
		return
	}
	defer array.Destroy()

	bounds, err := array.BoundsInfo()
	if err != nil {
		return
	}
	fmt.Println(bounds[0].Elements, bounds[0].LowerBound)
	// Output:
	// 5 0
}

func ExampleSafeArray_Copy() {
	original, err := FromSlice([]int32{1, 2, 3})
	if err != nil {
		return
	}
	defer original.Destroy()

	clone, err := original.Copy()
	if err != nil {
		return
	}
	defer clone.Destroy()

	values, _ := ToSlice[[]int32](clone)
	fmt.Println(values[0], values[1], values[2])
	// Output:
	// 1 2 3
}

// ---- SafeArray: Multi-Dimensional ----

// FromSlice creates a 2D SafeArray from a nested Go slice. Each inner slice
// must have the same length (rectangular shape).
func ExampleFromSlice_multiDimensional2D() {
	array, err := FromSlice([][]int32{
		{10, 20, 30},
		{40, 50, 60},
	})
	if err != nil {
		return
	}
	defer array.Destroy()

	dims, _ := array.GetDimensions()
	bounds, _ := array.BoundsInfo()
	fmt.Println("dimensions:", dims)
	fmt.Println("rows:", bounds[0].Elements, "cols:", bounds[1].Elements)

	values, _ := ToSlice[[][]int32](array)
	fmt.Println(values[0])
	fmt.Println(values[1])
	// Output:
	// dimensions: 2
	// rows: 2 cols: 3
	// [10 20 30]
	// [40 50 60]
}

// FromSlice supports 3D nested slices as well, as long as the shape is
// rectangular at every level.
func ExampleFromSlice_multiDimensional3D() {
	array, err := FromSlice([][][]int32{
		{{1, 2}, {3, 4}},
		{{5, 6}, {7, 8}},
		{{9, 10}, {11, 12}},
	})
	if err != nil {
		return
	}
	defer array.Destroy()

	dims, _ := array.GetDimensions()
	bounds, _ := array.BoundsInfo()
	fmt.Println("dimensions:", dims)
	fmt.Println("dim0:", bounds[0].Elements, "dim1:", bounds[1].Elements, "dim2:", bounds[2].Elements)

	values, _ := ToSlice[[][][]int32](array)
	fmt.Println(values[0][0], values[0][1])
	fmt.Println(values[2][0], values[2][1])
	// Output:
	// dimensions: 3
	// dim0: 3 dim1: 2 dim2: 2
	// [1 2] [3 4]
	// [9 10] [11 12]
}

// Create builds a multi-dimensional SafeArray from explicit bounds. Use
// PutElement to populate individual cells by index.
func ExampleCreate_multiDimensional() {
	// Create a 3x4 array of int32 (3 rows, 4 columns).
	array, err := Create(VT_I4, []SafeArrayBound{
		{Elements: 4, LowerBound: 0}, // columns (innermost dimension in COM)
		{Elements: 3, LowerBound: 0}, // rows (outermost dimension in COM)
	})
	if err != nil {
		return
	}
	defer array.Destroy()

	// PutElement indices are in Go slice order: [row, col].
	array.PutElement([]int32{0, 0}, int32(1))
	array.PutElement([]int32{0, 1}, int32(2))
	array.PutElement([]int32{1, 0}, int32(3))
	array.PutElement([]int32{2, 3}, int32(99))

	values, _ := ToSlice[[][]int32](array)
	fmt.Println(values[0][0], values[0][1])
	fmt.Println(values[1][0])
	fmt.Println(values[2][3])
	// Output:
	// 1 2
	// 3
	// 99
}

// ---- WinRT Initialization ----

func ExampleRoInitialize() {
	result, err := RoInitialize(RoSingleThreaded)
	if err != nil {
		return
	}
	if result == SuccessfullyInitialized || result == AlreadyInitialized {
		defer RoUninitialize()
	}
}
