//go:build windows

package ole

import (
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	iidIActivationFactoryTest, _      = windows.GUIDFromString("{00000035-0000-0000-C000-000000000046}")
	iidIUriRuntimeClassFactoryTest, _ = windows.GUIDFromString("{44A9796F-723E-4FDF-A218-033E75B0C084}")
)

type testIUriRuntimeClassFactory struct {
	VirtualTable *testIUriRuntimeClassFactoryVirtualTable
}

type testIUriRuntimeClassFactoryVirtualTable struct {
	QueryInterface      uintptr
	AddRef              uintptr
	Release             uintptr
	GetIIds             uintptr
	GetRuntimeClassName uintptr
	GetTrustLevel       uintptr
	CreateUri           uintptr
	CreateWithRelative  uintptr
}

func (obj *testIUriRuntimeClassFactory) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *testIUriRuntimeClassFactory) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

func (obj *testIUriRuntimeClassFactory) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

func (obj *testIUriRuntimeClassFactory) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *testIUriRuntimeClassFactory) CreateUri(rawURI HString) (uri *IInspectable, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.CreateUri,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(rawURI),
		uintptr(unsafe.Pointer(&uri)),
	)

	if windows.Handle(hr) != windows.S_OK {
		err = windows.Errno(hr)
	}

	return
}

func TestWinRT_RoGetActivationFactory_URI(t *testing.T) {
	result, err := RoInitialize(RoMultithreaded)
	if err != nil {
		t.Fatalf("RoInitialize failed: %v", err)
	}
	if result == IncompatibleConcurrencyModelAlreadyInitialized {
		t.Skip("WinRT already initialized with an incompatible concurrency model")
	}
	defer RoUninitialize()

	factory, err := RoGetActivationFactory("Windows.Foundation.Uri", iidIActivationFactoryTest)
	if err != nil {
		t.Fatalf("RoGetActivationFactory failed: %v", err)
	}
	defer factory.Release()

	className, err := factory.GetRuntimeClassName()
	if err != nil {
		t.Fatalf("IActivationFactory.GetRuntimeClassName failed: %v", err)
	}
	if className != "Windows.Foundation.Uri" {
		t.Fatalf("unexpected activation factory class name: got %q", className)
	}

	uriFactory, err := QueryInterfaceOnIUnknown[testIUriRuntimeClassFactory](factory, iidIUriRuntimeClassFactoryTest)
	if err != nil {
		t.Fatalf("querying IUriRuntimeClassFactory failed: %v", err)
	}
	defer uriFactory.Release()

	rawURI, err := NewHString("https://example.com/")
	if err != nil {
		t.Fatalf("NewHString failed: %v", err)
	}
	defer DeleteHString(rawURI)

	uri, err := uriFactory.CreateUri(rawURI)
	if err != nil {
		t.Fatalf("CreateUri failed: %v", err)
	}
	defer uri.Release()

	runtimeClassName, err := uri.GetRuntimeClassName()
	if err != nil {
		t.Fatalf("created URI GetRuntimeClassName failed: %v", err)
	}
	if runtimeClassName != "Windows.Foundation.Uri" {
		t.Fatalf("unexpected created runtime class name: got %q", runtimeClassName)
	}
}
