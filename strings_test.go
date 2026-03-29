//go:build windows

package ole

import (
	"strings"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestSysAllocStringRoundTrip(t *testing.T) {
	bstr := SysAllocString("Hello, 世界")
	if bstr == nil {
		t.Fatal("SysAllocString returned nil")
	}
	defer SysFreeString(bstr)

	if got := windows.UTF16PtrToString(bstr); got != "Hello, 世界" {
		t.Fatalf("UTF16PtrToString() = %q, want %q", got, "Hello, 世界")
	}

	if got := SysStringLen((*int16)(unsafe.Pointer(bstr))); got != uint32(len([]rune("Hello, 世界"))) {
		t.Fatalf("SysStringLen() = %d, want %d", got, len([]rune("Hello, 世界")))
	}
}

func TestSysAllocStringEmpty(t *testing.T) {
	bstr := SysAllocString("")
	if bstr == nil {
		t.Fatal("SysAllocString returned nil for empty string")
	}
	defer SysFreeString(bstr)

	if got := windows.UTF16PtrToString(bstr); got != "" {
		t.Fatalf("UTF16PtrToString() = %q, want empty string", got)
	}

	if got := SysStringLen((*int16)(unsafe.Pointer(bstr))); got != 0 {
		t.Fatalf("SysStringLen() = %d, want 0", got)
	}
}

func TestSysAllocStringLenCountsEmbeddedNull(t *testing.T) {
	value := "a\x00b"
	bstr := SysAllocStringLen(value)
	if bstr == nil {
		t.Fatal("SysAllocStringLen returned nil")
	}
	defer SysFreeString(bstr)

	if got := SysStringLen((*int16)(unsafe.Pointer(bstr))); got != 3 {
		t.Fatalf("SysStringLen() = %d, want 3", got)
	}

	if got := windows.UTF16PtrToString(bstr); got != "a" {
		t.Fatalf("UTF16PtrToString() = %q, want %q", got, "a")
	}
}

func TestGetErrorDescription(t *testing.T) {
	description := GetErrorDescription(2)
	if description == "" {
		t.Fatal("GetErrorDescription returned empty string")
	}

	if strings.HasSuffix(description, "\n") || strings.HasSuffix(description, "\r") {
		t.Fatalf("GetErrorDescription() = %q, should not end with newline", description)
	}
}
