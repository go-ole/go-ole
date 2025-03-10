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
