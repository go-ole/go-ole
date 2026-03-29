//go:build windows

package ole

import (
	"testing"
)

func TestIUnknown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Error(r)
		}
	}()

	var err error

	err = Initialize(0)
	if err != nil {
		t.Fatal(err)
	}

	defer Uninitialize()

	var unknown *IUnknown

	unknown, err = CreateInstance(CLSID_COMEchoTestObject, IID_IUnknown)
	if err == nil {
		defer unknown.Release()
	}
}

func TestIUnknown_AddRefRelease(t *testing.T) {
	var refCount uint32
	vt := &IUnknownVirtualTable{
		AddRef: syscall.NewCallback(func(this uintptr) uintptr {
			refCount++
			return uintptr(refCount)
		}),
		Release: syscall.NewCallback(func(this uintptr) uintptr {
			refCount--
			return uintptr(refCount)
		}),
	}
	unknown := &IUnknown{VirtualTable: vt}

	if count := unknown.AddRef(); count != 1 {
		t.Errorf("AddRef() = %d, want 1", count)
	}
	if count := unknown.Release(); count != 0 {
		t.Errorf("Release() = %d, want 0", count)
	}
}

func TestIUnknown_QueryInterface(t *testing.T) {
	t.Run("success", func(t *testing.T) {
		want := &IUnknownVirtualTable{}
		vt := &IUnknownVirtualTable{
			QueryInterface: syscall.NewCallback(func(this uintptr, riid uintptr, ppv uintptr) uintptr {
				*(*uintptr)(unsafe.Pointer(ppv)) = uintptr(unsafe.Pointer(want))
				return uintptr(windows.S_OK)
			}),
		}
		unknown := &IUnknown{VirtualTable: vt}

		got, err := QueryInterfaceOnIUnknown[IUnknownVirtualTable](unknown, IID_IUnknown)
		if err != nil {
			t.Fatalf("QueryInterface failed: %v", err)
		}
		if got != want {
			t.Fatalf("QueryInterface() = %p, want %p", got, want)
		}
	})

	t.Run("no interface", func(t *testing.T) {
		vt := &IUnknownVirtualTable{
			QueryInterface: syscall.NewCallback(func(this uintptr, riid uintptr, ppv uintptr) uintptr {
				return uintptr(windows.E_NOINTERFACE)
			}),
		}
		unknown := &IUnknown{VirtualTable: vt}

		_, err := QueryInterfaceOnIUnknown[IUnknownVirtualTable](unknown, IID_IUnknown)
		if err != ComInterfaceNotImplementedError {
			t.Fatalf("QueryInterface error = %v, want ComInterfaceNotImplementedError", err)
		}
	})
}
