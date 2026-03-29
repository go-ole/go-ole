//go:build windows

package ole

import "testing"
import "golang.org/x/sys/windows"

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
