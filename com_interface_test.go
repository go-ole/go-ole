//go:build windows
// +build windows

package ole

import (
	"golang.org/x/sys/windows"
	"testing"
)

func TestInitialize(t *testing.T) {
	result, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	if result != SuccessfullyInitialized && result != AlreadyInitialized {
		t.Errorf("Expected SuccessfullyInitialized or AlreadyInitialized, got %v", result)
	}
}

func TestInitializeApartmentThreaded(t *testing.T) {
	result, err := InitializeApartmentThreaded()
	if err != nil {
		t.Fatalf("InitializeApartmentThreaded failed: %v", err)
	}
	defer Uninitialize()

	if result != SuccessfullyInitialized && result != AlreadyInitialized && result != IncompatibleConcurrencyModelAlreadyInitialized {
		t.Errorf("Expected valid result, got %v", result)
	}
}

func TestLookupClassId(t *testing.T) {
	_, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	// Use a common CLSID that should be present on Windows
	clsid, err := LookupClassId("InternetExplorer.Application")
	if err != nil {
		t.Skip("InternetExplorer.Application not found, skipping lookup test")
	}

	if clsid == (windows.GUID{}) {
		t.Error("Expected non-empty CLSID for InternetExplorer.Application")
	}
}

func TestCreateInstance(t *testing.T) {
	_, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	clsid, err := LookupClassId("InternetExplorer.Application")
	if err != nil {
		t.Skip("InternetExplorer.Application not found, skipping instance creation test")
	}

	unknown, err := CreateInstance[IUnknown](clsid, IID_IUnknown)
	if err != nil {
		t.Fatalf("CreateInstance failed: %v", err)
	}
	if unknown == nil {
		t.Fatal("Expected unknown to be non-nil")
	}
	unknown.Release()
}

func TestGetActiveObject(t *testing.T) {
	_, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	clsid, err := LookupClassId("InternetExplorer.Application")
	if err != nil {
		t.Skip("InternetExplorer.Application not found, skipping active object test")
	}

	// This might fail if IE is not running, which is fine
	obj, err := GetActiveObject[IDispatch](clsid, IID_IDispatch)
	if err != nil {
		t.Logf("GetActiveObject failed (expected if not running): %v", err)
		return
	}
	if obj != nil {
		obj.Release()
	}
}
