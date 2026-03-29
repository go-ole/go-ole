//go:build windows

package ole

import (
	"errors"
	"testing"

	"golang.org/x/sys/windows"
)

func requirePanicError(t *testing.T, want error) func() {
	t.Helper()

	return func() {
		t.Helper()

		recovered := recover()
		if recovered == nil {
			t.Fatalf("expected panic %v", want)
		}

		err, ok := recovered.(error)
		if !ok {
			t.Fatalf("panic = %T, want error %v", recovered, want)
		}

		if !errors.Is(err, want) {
			t.Fatalf("panic = %v, want %v", err, want)
		}
	}
}

func TestMakeNullVariant(t *testing.T) {
	variant := MakeNullVariant()
	if variant.VT != VT_NULL {
		t.Error("Null variant should be of type VT_NULL")
	}

	if variant.Val != 0 {
		t.Error("VT_NULL variant should not have a value set")
	}
}

func TestVariantToNull(t *testing.T) {
	variant := VariantToNull(MakeNullVariant())
	if variant != nil {
		t.Error("Identity guarantee for VariantToNull failed with VT_NULL Variant")
	}
}

func TestMakeEmptyVariant(t *testing.T) {
	variant := MakeEmptyVariant()
	if variant.VT != VT_EMPTY {
		t.Error("Empty variant should be of type VT_EMPTY")
	}

	if variant.Val != 0 {
		t.Error("VT_EMPTY variant should not have a value set")
	}
}

func TestVariantToEmpty(t *testing.T) {
	variant := VariantToEmpty(MakeEmptyVariant())
	if variant != "" {
		t.Error("Identity guarantee for VariantToEmpty failed with VT_EMPTY Variant")
	}
}

func TestErrorToVariant(t *testing.T) {
	variant := ErrorToVariant(windows.Handle(uintptr(100)))
	if variant.VT != VT_ERROR {
		t.Error("Error variant should be of type VT_ERROR")
	}

	if variant.Val != int64(100) {
		t.Error("VT_ERROR does not match")
	}
}

func TestVariantToError(t *testing.T) {
	variant := VariantToError(ErrorToVariant(windows.Handle(uintptr(100)))).(windows.Handle)
	if variant != windows.Handle(uintptr(100)) {
		t.Error("Identity guarantee for VariantToError failed with VT_ERROR Variant")
	}
}

func TestHandleToVariant(t *testing.T) {
	variant := HandleToVariant(windows.Handle(uintptr(100)))
	if variant.VT != VT_HRESULT {
		t.Error("Error variant should be of type VT_HRESULT")
	}

	if variant.Val != int64(100) {
		t.Error("VT_HRESULT does not match")
	}
}

func TestVariantToHandle(t *testing.T) {
	variant := VariantToHandle(HandleToVariant(windows.Handle(uintptr(100)))).(windows.Handle)
	if variant != windows.Handle(uintptr(100)) {
		t.Error("Identity guarantee for VariantToError failed with VT_HRESULT Variant")
	}
}

func TestBoolToVariant(t *testing.T) {
	variant := BoolToVariant(true)
	if variant.VT != VT_BOOL {
		t.Error("Bool variant should be of type VT_BOOL")
	}
	if variant.Val != int64(VariantTypeTrue) {
		t.Error("Bool value should be VariantTypeTrue")
	}

	variant = BoolToVariant(false)
	if variant.Val != int64(VariantTypeFalse) {
		t.Error("Bool value should be VariantTypeFalse")
	}
}

func TestVariantToBool(t *testing.T) {
	variant := &VARIANT{VT: VT_BOOL, Val: int64(VariantTypeTrue)}
	if VariantToBool(variant).(bool) != true {
		t.Error("Identity guarantee for VariantToBool failed")
	}

	variant.Val = int64(VariantTypeFalse)
	if VariantToBool(variant).(bool) != false {
		t.Error("Identity guarantee for VariantToBool failed")
	}
}

func TestInt32ToVariant(t *testing.T) {
	variant := Int32ToVariant(int32(12345))
	if variant.VT != VT_I4 {
		t.Error("Int32 variant should be of type VT_I4")
	}
	if variant.Val != 12345 {
		t.Error("Int32 value mismatch")
	}
}

func TestVariantToInt32(t *testing.T) {
	variant := &VARIANT{VT: VT_I4, Val: 12345}
	if VariantToInt32(variant).(int32) != 12345 {
		t.Error("Identity guarantee for VariantToInt32 failed")
	}
}

func TestFloat64ToVariant(t *testing.T) {
	val := 123.456
	variant := Float64ToVariant(val)
	if variant.VT != VT_R8 {
		t.Error("Float64 variant should be of type VT_R8")
	}
}

func TestVariantToFloat64(t *testing.T) {
	val := 123.456
	variant := Float64ToVariant(val)
	if VariantToFloat64(variant).(float64) != val {
		t.Error("Identity guarantee for VariantToFloat64 failed")
	}
}

func TestRegisterVariantConverter(t *testing.T) {
	type CustomType struct {
		Value int
	}

	vt := VT(0x9999)
	to := func(i any) *VARIANT {
		return &VARIANT{VT: vt, Val: int64(i.(CustomType).Value)}
	}
	from := func(v *VARIANT) any {
		return CustomType{Value: int(v.Val)}
	}

	RegisterVariantConverter[CustomType](vt, to, from)
	defer DeregisterVariantConverter[CustomType](vt)

	val := CustomType{Value: 42}
	variant, err := WrapVariant(val)
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	if variant.VT != vt || variant.Val != 42 {
		t.Fatalf("WrapVariant result mismatch: %v", variant)
	}

	unwrapped := UnwrapVariant[CustomType](variant)
	if unwrapped.Value != 42 {
		t.Fatalf("UnwrapVariant result mismatch: %v", unwrapped)
	}
}

func TestWrapVariantUnsupportedType(t *testing.T) {
	type UnsupportedType struct{}

	_, err := WrapVariant(UnsupportedType{})
	if !errors.Is(err, UnsupportedNativeType) {
		t.Fatalf("WrapVariant error = %v, want %v", err, UnsupportedNativeType)
	}
}

func TestRegisterToVariantConverter(t *testing.T) {
	type ToOnlyType struct {
		Value int
	}

	RegisterToVariantConverter[ToOnlyType](func(i any) *VARIANT {
		return &VARIANT{VT: VT_I4, Val: int64(i.(ToOnlyType).Value)}
	})
	defer DeregisterToVariantConverter[ToOnlyType]()

	variant, err := WrapVariant(ToOnlyType{Value: 77})
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}

	if variant.VT != VT_I4 || variant.Val != 77 {
		t.Fatalf("WrapVariant result mismatch: %v", variant)
	}
}

func TestDeregisterToVariantConverter(t *testing.T) {
	type ToOnlyType struct {
		Value int
	}

	RegisterToVariantConverter[ToOnlyType](func(i any) *VARIANT {
		return &VARIANT{VT: VT_I4, Val: int64(i.(ToOnlyType).Value)}
	})
	DeregisterToVariantConverter[ToOnlyType]()

	_, err := WrapVariant(ToOnlyType{Value: 88})
	if !errors.Is(err, UnsupportedNativeType) {
		t.Fatalf("WrapVariant error = %v, want %v", err, UnsupportedNativeType)
	}
}

func TestRegisterFromVariantConverter(t *testing.T) {
	vt := VT(0x9998)

	RegisterFromVariantConverter(vt, func(v *VARIANT) any {
		return int(v.Val)
	})
	defer DeregisterFromVariantConverter(vt)

	value := UnwrapVariant[int](&VARIANT{VT: vt, Val: 64})
	if value != 64 {
		t.Fatalf("UnwrapVariant result mismatch: %v", value)
	}
}

func TestDeregisterFromVariantConverter(t *testing.T) {
	vt := VT(0x9997)

	RegisterFromVariantConverter(vt, func(v *VARIANT) any {
		return int(v.Val)
	})
	DeregisterFromVariantConverter(vt)

	defer requirePanicError(t, UnsupportedNativeType)()

	_ = UnwrapVariant[int](&VARIANT{VT: vt, Val: 64})
}

func TestUnwrapVariantUnsupportedType(t *testing.T) {
	defer requirePanicError(t, UnsupportedNativeType)()

	_ = UnwrapVariant[int](&VARIANT{VT: VT(0x9996)})
}

func TestWrapParametersWithVariant(t *testing.T) {
	RegisterVariantConverters()

	args := WrapParametersWithVariant(int32(12), nil, true)
	if len(args) != 3 {
		t.Fatalf("WrapParametersWithVariant length = %d, want 3", len(args))
	}

	if args[0].VT != VT_BOOL || VariantToBool(args[0]).(bool) != true {
		t.Fatalf("args[0] = %v, want VT_BOOL true", args[0])
	}

	if args[1].VT != VT_NULL || args[1].Val != 0 {
		t.Fatalf("args[1] = %v, want VT_NULL", args[1])
	}

	if args[2].VT != VT_I4 || VariantToInt32(args[2]).(int32) != 12 {
		t.Fatalf("args[2] = %v, want VT_I4 12", args[2])
	}
}
