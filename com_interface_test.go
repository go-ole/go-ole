//go:build windows
// +build windows

package ole

import (
	"golang.org/x/sys/windows"
	"testing"
)

const testProgramID = "Shell.Application"

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

	clsid, err := ClassIdFromString(testProgramID)
	if err != nil {
		t.Skipf("%s not found, skipping lookup test: %v", testProgramID, err)
	}

	if clsid == (windows.GUID{}) {
		t.Errorf("Expected non-empty CLSID for %s", testProgramID)
	}
}

func TestCreateInstance(t *testing.T) {
	_, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	clsid, err := ClassIdFromString(testProgramID)
	if err != nil {
		t.Skipf("%s not found, skipping instance creation test: %v", testProgramID, err)
	}

	unknown, err := CreateInstance[*IUnknown](clsid, IID_IUnknown)
	if err != nil {
		t.Fatalf("CreateInstance failed: %v", err)
	}
	if unknown == nil {
		t.Fatal("Expected unknown to be non-nil")
	}
	(*unknown).Release()
}

func TestGetActiveObject(t *testing.T) {
	_, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	clsid, err := ClassIdFromString(testProgramID)
	if err != nil {
		t.Skipf("%s not found, skipping active object test: %v", testProgramID, err)
	}

	obj, err := GetActiveObject[*IDispatch](clsid, IID_IDispatch)
	if err != nil {
		t.Logf("GetActiveObject failed (expected if not running): %v", err)
		return
	}
	if obj != nil {
		(*obj).Release()
	}
}
