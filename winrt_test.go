//go:build windows

package ole

import (
	"testing"

	"golang.org/x/sys/windows"
)

func TestWinRT_XMLDocument(t *testing.T) {
	// IXmlDocumentIO is ABI.Windows.Data.Xml.Dom.IXmlDocumentIO
	IXmlDocumentIO, err := windows.GUIDFromString("{6cd0e74e-ee65-4489-9ebf-ca43e87ba637}")
	if err != nil {
		t.Error(err)
		return
	}

	RoInitialize(RoMultithreaded)
	defer RoUninitialize()

	inspectable, err := RoActivateInstance("Windows.Data.Xml.Dom.XmlDocument")
	if err != nil {
		t.Error(err)
		return
	}
	defer inspectable.Release()

	xmldoc, err := QueryInterfaceOnIUnknown[IDispatch](inspectable, IXmlDocumentIO)
	if err != nil {
		t.Error(err)
		return
	}
	defer xmldoc.Release()

	hString, err := NewHString("<test></test>")
	if err != nil {
		t.Error(err)
		return
	}
	defer DeleteHString(hString)

	// panics with "unknown type"
	xmldoc.CallMethod("LoadXml", HStringToVariant(hString))
}

func TestWinRT_RoInitializeSingleThreaded(t *testing.T) {
	result, err := RoInitialize(RoSingleThreaded)
	if err != nil {
		t.Fatalf("RoInitialize failed: %v", err)
	}
	if result == IncompatibleConcurrencyModelAlreadyInitialized {
		t.Skip("WinRT already initialized with an incompatible concurrency model")
	}
	defer RoUninitialize()

	if result != SuccessfullyInitialized && result != AlreadyInitialized {
		t.Fatalf("RoInitialize result = %v, want initialized state", result)
	}
}

func TestWinRT_RoActivateInstanceInvalidClass(t *testing.T) {
	result, err := RoInitialize(RoMultithreaded)
	if err != nil {
		t.Fatalf("RoInitialize failed: %v", err)
	}
	if result == IncompatibleConcurrencyModelAlreadyInitialized {
		t.Skip("WinRT already initialized with an incompatible concurrency model")
	}
	defer RoUninitialize()

	obj, err := RoActivateInstance("Windows.Invalid.Class")
	if err == nil {
		if obj != nil {
			obj.Release()
		}
		t.Fatal("RoActivateInstance should fail for invalid class")
	}
	if obj != nil {
		obj.Release()
		t.Fatal("RoActivateInstance returned non-nil object for invalid class")
	}
}

func TestWinRT_RoGetActivationFactoryInvalidClass(t *testing.T) {
	result, err := RoInitialize(RoMultithreaded)
	if err != nil {
		t.Fatalf("RoInitialize failed: %v", err)
	}
	if result == IncompatibleConcurrencyModelAlreadyInitialized {
		t.Skip("WinRT already initialized with an incompatible concurrency model")
	}
	defer RoUninitialize()

	factory, err := RoGetActivationFactory("Windows.Invalid.Class", IID_IActivationFactory)
	if err == nil {
		if factory != nil {
			factory.Release()
		}
		t.Fatal("RoGetActivationFactory should fail for invalid class")
	}
	if factory != nil {
		factory.Release()
		t.Fatal("RoGetActivationFactory returned non-nil factory for invalid class")
	}
}

func TestWinRT_RoActivateInstanceRuntimeClassName(t *testing.T) {
	result, err := RoInitialize(RoMultithreaded)
	if err != nil {
		t.Fatalf("RoInitialize failed: %v", err)
	}
	if result == IncompatibleConcurrencyModelAlreadyInitialized {
		t.Skip("WinRT already initialized with an incompatible concurrency model")
	}
	defer RoUninitialize()

	inspectable, err := RoActivateInstance("Windows.Data.Xml.Dom.XmlDocument")
	if err != nil {
		t.Fatalf("RoActivateInstance failed: %v", err)
	}
	defer inspectable.Release()

	className, err := inspectable.GetRuntimeClassName()
	if err != nil {
		t.Fatalf("GetRuntimeClassName failed: %v", err)
	}
	if className != "Windows.Data.Xml.Dom.XmlDocument" {
		t.Fatalf("runtime class name = %q, want %q", className, "Windows.Data.Xml.Dom.XmlDocument")
	}
}
