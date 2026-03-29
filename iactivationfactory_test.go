//go:build windows

package ole

import (
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestIActivationFactoryAddressMethods(t *testing.T) {
	virtualTable := &IActivationFactoryVirtualTable{
		QueryInterface:      1,
		addRef:              2,
		release:             3,
		getIIds:             4,
		getRuntimeClassName: 5,
		getTrustLevel:       6,
		activateInstance:    7,
	}
	factory := &IActivationFactory{VirtualTable: virtualTable}

	if factory.QueryInterfaceAddress() != 1 {
		t.Fatalf("QueryInterfaceAddress() = %d, want 1", factory.QueryInterfaceAddress())
	}
	if factory.AddRefAddress() != 2 {
		t.Fatalf("AddRefAddress() = %d, want 2", factory.AddRefAddress())
	}
	if factory.ReleaseAddress() != 3 {
		t.Fatalf("ReleaseAddress() = %d, want 3", factory.ReleaseAddress())
	}
	if factory.GetInterfaceIdsAddress() != 4 {
		t.Fatalf("GetInterfaceIdsAddress() = %d, want 4", factory.GetInterfaceIdsAddress())
	}
	if factory.GetRuntimeClassNameAddress() != 5 {
		t.Fatalf("GetRuntimeClassNameAddress() = %d, want 5", factory.GetRuntimeClassNameAddress())
	}
	if factory.GetTrustLevelAddress() != 6 {
		t.Fatalf("GetTrustLevelAddress() = %d, want 6", factory.GetTrustLevelAddress())
	}
	if factory.ActivateInstanceAddress() != 7 {
		t.Fatalf("ActivateInstanceAddress() = %d, want 7", factory.ActivateInstanceAddress())
	}
}

func TestIActivationFactoryGetRuntimeClassName(t *testing.T) {
	hstring, err := NewHString("Windows.Foundation.Uri")
	if err != nil {
		t.Fatalf("NewHString failed: %v", err)
	}

	virtualTable := &IActivationFactoryVirtualTable{
		getRuntimeClassName: syscall.NewCallback(func(this uintptr, className uintptr) uintptr {
			*(*HString)(unsafe.Pointer(className)) = hstring
			return uintptr(windows.S_OK)
		}),
	}
	factory := &IActivationFactory{VirtualTable: virtualTable}

	got, err := factory.GetRuntimeClassName()
	if err != nil {
		t.Fatalf("GetRuntimeClassName failed: %v", err)
	}
	if got != "Windows.Foundation.Uri" {
		t.Fatalf("GetRuntimeClassName() = %q, want %q", got, "Windows.Foundation.Uri")
	}
}

func TestIActivationFactoryGetTrustLevel(t *testing.T) {
	virtualTable := &IActivationFactoryVirtualTable{
		getTrustLevel: syscall.NewCallback(func(this uintptr, level uintptr) uintptr {
			*(*uint32)(unsafe.Pointer(level)) = uint32(FullTrust)
			return uintptr(windows.S_OK)
		}),
	}
	factory := &IActivationFactory{VirtualTable: virtualTable}

	if got := factory.GetTrustLevel(); got != FullTrust {
		t.Fatalf("GetTrustLevel() = %d, want %d", got, FullTrust)
	}
}

func TestIActivationFactoryGetInterfaceIds(t *testing.T) {
	want := []windows.GUID{IID_IInspectable, iidIActivationFactoryTest}
	buffer := make([]windows.GUID, len(want))
	copy(buffer, want)
	raw := unsafe.Pointer(&buffer[0])

	virtualTable := &IActivationFactoryVirtualTable{
		getIIds: syscall.NewCallback(func(this uintptr, count uintptr, ids uintptr) uintptr {
			*(*uint32)(unsafe.Pointer(count)) = uint32(len(want))
			*(**windows.GUID)(unsafe.Pointer(ids)) = (*windows.GUID)(raw)
			return uintptr(windows.S_OK)
		}),
	}
	factory := &IActivationFactory{VirtualTable: virtualTable}

	got, err := factory.GetInterfaceIds()
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

func TestIActivationFactoryActivateInstance(t *testing.T) {
	t.Run("success", func(t *testing.T) {
		want := &IInspectable{}
		virtualTable := &IActivationFactoryVirtualTable{
			activateInstance: syscall.NewCallback(func(this uintptr, instance uintptr) uintptr {
				*(**IInspectable)(unsafe.Pointer(instance)) = want
				return uintptr(windows.S_OK)
			}),
		}
		factory := &IActivationFactory{VirtualTable: virtualTable}

		got, err := factory.ActivateInstance()
		if err != nil {
			t.Fatalf("ActivateInstance failed: %v", err)
		}
		if got != want {
			t.Fatalf("ActivateInstance() = %p, want %p", got, want)
		}
	})

	t.Run("error", func(t *testing.T) {
		virtualTable := &IActivationFactoryVirtualTable{
			activateInstance: syscall.NewCallback(func(this uintptr, instance uintptr) uintptr {
				return uintptr(windows.E_POINTER)
			}),
		}
		factory := &IActivationFactory{VirtualTable: virtualTable}

		got, err := factory.ActivateInstance()
		if got != nil {
			t.Fatalf("ActivateInstance() = %p, want nil", got)
		}
		if err != windows.Errno(windows.E_POINTER) {
			t.Fatalf("ActivateInstance() error = %v, want %v", err, windows.Errno(windows.E_POINTER))
		}
	})
}
