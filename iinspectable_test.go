//go:build windows

package ole

import (
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestIInspectableAddressMethods(t *testing.T) {
	virtualTable := &IInspectableVirtualTable{
		QueryInterface:      1,
		AddRef:              2,
		Release:             3,
		GetIIds:             4,
		GetRuntimeClassName: 5,
		GetTrustLevel:       6,
	}
	inspectable := &IInspectable{VirtualTable: virtualTable}

	if inspectable.QueryInterfaceAddress() != 1 {
		t.Fatalf("QueryInterfaceAddress() = %d, want 1", inspectable.QueryInterfaceAddress())
	}
	if inspectable.AddRefAddress() != 2 {
		t.Fatalf("AddRefAddress() = %d, want 2", inspectable.AddRefAddress())
	}
	if inspectable.ReleaseAddress() != 3 {
		t.Fatalf("ReleaseAddress() = %d, want 3", inspectable.ReleaseAddress())
	}
	if inspectable.GetInterfaceIdsAddress() != 4 {
		t.Fatalf("GetInterfaceIdsAddress() = %d, want 4", inspectable.GetInterfaceIdsAddress())
	}
	if inspectable.GetRuntimeClassNameAddress() != 5 {
		t.Fatalf("GetRuntimeClassNameAddress() = %d, want 5", inspectable.GetRuntimeClassNameAddress())
	}
	if inspectable.GetTrustLevelAddress() != 6 {
		t.Fatalf("GetTrustLevelAddress() = %d, want 6", inspectable.GetTrustLevelAddress())
	}
}

func TestIInspectableGetInterfaceIds(t *testing.T) {
	want := []windows.GUID{IID_IInspectable, iidIActivationFactoryTest}
	raw := windows.CoTaskMemAlloc(uintptr(len(want)) * unsafe.Sizeof(want[0]))
	if raw == nil {
		t.Fatal("CoTaskMemAlloc returned nil")
	}
	buffer := unsafe.Slice((*windows.GUID)(raw), len(want))
	copy(buffer, want)

	virtualTable := &IInspectableVirtualTable{
		GetIIds: syscall.NewCallback(func(this uintptr, count uintptr, ids uintptr) uintptr {
			*(*uint32)(unsafe.Pointer(count)) = uint32(len(want))
			*(**windows.GUID)(unsafe.Pointer(ids)) = (*windows.GUID)(raw)
			return uintptr(windows.S_OK)
		}),
	}
	inspectable := &IInspectable{VirtualTable: virtualTable}

	got, err := inspectable.GetInterfaceIds()
	if err != nil {
		t.Fatalf("GetInterfaceIds failed: %v", err)
	}
	if len(got) != len(want) {
		t.Fatalf("GetInterfaceIds len = %d, want %d", len(got), len(want))
	}
	for index := range want {
		if got[index] != want[index] {
			t.Fatalf("GetInterfaceIds[%d] = %v, want %v", index, got[index], want[index])
		}
	}
}

func TestIInspectableGetInterfaceIdsEmpty(t *testing.T) {
	virtualTable := &IInspectableVirtualTable{
		GetIIds: syscall.NewCallback(func(this uintptr, count uintptr, ids uintptr) uintptr {
			*(*uint32)(unsafe.Pointer(count)) = 0
			*(**windows.GUID)(unsafe.Pointer(ids)) = nil
			return uintptr(windows.S_OK)
		}),
	}
	inspectable := &IInspectable{VirtualTable: virtualTable}

	got, err := inspectable.GetInterfaceIds()
	if err != nil {
		t.Fatalf("GetInterfaceIds failed: %v", err)
	}
	if got != nil {
		t.Fatalf("GetInterfaceIds() = %#v, want nil", got)
	}
}

func TestIInspectableGetRuntimeClassName(t *testing.T) {
	hstring, err := NewHString("Windows.Foundation.Uri")
	if err != nil {
		t.Fatalf("NewHString failed: %v", err)
	}

	virtualTable := &IInspectableVirtualTable{
		GetRuntimeClassName: syscall.NewCallback(func(this uintptr, className uintptr) uintptr {
			*(*HString)(unsafe.Pointer(className)) = hstring
			return uintptr(windows.S_OK)
		}),
	}
	inspectable := &IInspectable{VirtualTable: virtualTable}

	got, err := inspectable.GetRuntimeClassName()
	if err != nil {
		t.Fatalf("GetRuntimeClassName failed: %v", err)
	}
	if got != "Windows.Foundation.Uri" {
		t.Fatalf("GetRuntimeClassName() = %q, want %q", got, "Windows.Foundation.Uri")
	}
}

func TestIInspectableGetRuntimeClassNameError(t *testing.T) {
	virtualTable := &IInspectableVirtualTable{
		GetRuntimeClassName: syscall.NewCallback(func(this uintptr, className uintptr) uintptr {
			return uintptr(windows.E_POINTER)
		}),
	}
	inspectable := &IInspectable{VirtualTable: virtualTable}

	got, err := inspectable.GetRuntimeClassName()
	if got != "" {
		t.Fatalf("GetRuntimeClassName() = %q, want empty string", got)
	}
	if err != windows.Errno(windows.E_POINTER) {
		t.Fatalf("GetRuntimeClassName() error = %v, want %v", err, windows.Errno(windows.E_POINTER))
	}
}

func TestIInspectableGetTrustLevel(t *testing.T) {
	virtualTable := &IInspectableVirtualTable{
		GetTrustLevel: syscall.NewCallback(func(this uintptr, level uintptr) uintptr {
			*(*uint32)(unsafe.Pointer(level)) = uint32(PartialTrust)
			return uintptr(windows.S_OK)
		}),
	}
	inspectable := &IInspectable{VirtualTable: virtualTable}

	if got := inspectable.GetTrustLevel(); got != PartialTrust {
		t.Fatalf("GetTrustLevel() = %d, want %d", got, PartialTrust)
	}
}
