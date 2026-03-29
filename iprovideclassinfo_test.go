//go:build windows

package ole

import (
	"errors"
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestIProvideClassInfoAddressMethods(t *testing.T) {
	virtualTable := &IProvideClassInfoVirtualTable{
		QueryInterface: 1,
		AddRef:         2,
		Release:        3,
		GetClassInfo:   4,
	}
	provideClassInfo := &IProvideClassInfo{VirtualTable: virtualTable}

	if provideClassInfo.QueryInterfaceAddress() != 1 {
		t.Fatalf("QueryInterfaceAddress() = %d, want 1", provideClassInfo.QueryInterfaceAddress())
	}
	if provideClassInfo.AddRefAddress() != 2 {
		t.Fatalf("AddRefAddress() = %d, want 2", provideClassInfo.AddRefAddress())
	}
	if provideClassInfo.ReleaseAddress() != 3 {
		t.Fatalf("ReleaseAddress() = %d, want 3", provideClassInfo.ReleaseAddress())
	}
	if provideClassInfo.GetClassInfoAddress() != 4 {
		t.Fatalf("GetClassInfoAddress() = %d, want 4", provideClassInfo.GetClassInfoAddress())
	}
}

func TestIProvideClassInfoGetClassInfo(t *testing.T) {
	t.Run("success", func(t *testing.T) {
		want := &ITypeInfo{}
		virtualTable := &IProvideClassInfoVirtualTable{
			GetClassInfo: syscall.NewCallback(func(this uintptr, info uintptr) uintptr {
				*(**ITypeInfo)(unsafe.Pointer(info)) = want
				return uintptr(windows.S_OK)
			}),
		}
		provideClassInfo := &IProvideClassInfo{VirtualTable: virtualTable}

		got, err := provideClassInfo.GetClassInfo()
		if err != nil {
			t.Fatalf("GetClassInfo failed: %v", err)
		}
		if got != want {
			t.Fatalf("GetClassInfo() = %p, want %p", got, want)
		}
	})

	t.Run("error", func(t *testing.T) {
		virtualTable := &IProvideClassInfoVirtualTable{
			GetClassInfo: syscall.NewCallback(func(this uintptr, info uintptr) uintptr {
				return uintptr(windows.E_POINTER)
			}),
		}
		provideClassInfo := &IProvideClassInfo{VirtualTable: virtualTable}

		got, err := provideClassInfo.GetClassInfo()
		if got != nil {
			t.Fatalf("GetClassInfo() = %p, want nil", got)
		}
		if !errors.Is(err, windows.Errno(windows.E_POINTER)) {
			t.Fatalf("GetClassInfo() error = %v, want %v", err, windows.Errno(windows.E_POINTER))
		}
	})
}
