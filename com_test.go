//go:build windows
// +build windows

package ole

import (
	"errors"
	"testing"

	"golang.org/x/sys/windows"
)

const testCOMProgramID = "Shell.Application"

func TestComSetupAndShutDown(t *testing.T) {
	result, err := InitializeMultithreaded()
	if err != nil {
		t.Fatalf("InitializeMultithreaded failed: %v", err)
	}
	defer Uninitialize()

	if result != SuccessfullyInitialized && result != AlreadyInitialized {
		t.Fatalf("InitializeMultithreaded result = %v, want initialized state", result)
	}
}

func TestComPublicSetupAndShutDown(t *testing.T) {
	result, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	if result != SuccessfullyInitialized && result != AlreadyInitialized {
		t.Fatalf("Initialize result = %v, want initialized state", result)
	}
}

func TestComPublicSetupAndShutDown_WithValue(t *testing.T) {
	result, err := Initialize(ConcurrencyModel(5))
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	if result != SuccessfullyInitialized &&
		result != AlreadyInitialized &&
		result != IncompatibleConcurrencyModelAlreadyInitialized {
		t.Fatalf("Initialize result = %v, want valid state", result)
	}
}

func TestComExSetupAndShutDown(t *testing.T) {
	result, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	if result != SuccessfullyInitialized && result != AlreadyInitialized {
		t.Fatalf("Initialize result = %v, want initialized state", result)
	}
}

func TestComPublicExSetupAndShutDown(t *testing.T) {
	result, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	if result != SuccessfullyInitialized && result != AlreadyInitialized {
		t.Fatalf("Initialize result = %v, want initialized state", result)
	}
}

func TestComPublicExSetupAndShutDown_WithValue(t *testing.T) {
	result, err := Initialize(ConcurrencyModel(5))
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	if result != SuccessfullyInitialized &&
		result != AlreadyInitialized &&
		result != IncompatibleConcurrencyModelAlreadyInitialized {
		t.Fatalf("Initialize result = %v, want valid state", result)
	}
}

func TestClsidFromProgID(t *testing.T) {
	expected, err := ClassIdFromProgramId(testCOMProgramID)
	if err != nil {
		t.Skipf("%s not available: %v", testCOMProgramID, err)
	}

	actual, err := ClassIdFromString(testCOMProgramID)
	if err != nil {
		t.Fatalf("ClassIdFromString failed: %v", err)
	}

	if actual != expected {
		t.Fatalf("ClassIdFromString(%q) = %v, want %v", testCOMProgramID, actual, expected)
	}
}

func TestClsidFromString(t *testing.T) {
	expected, err := ClassIdFromGuidString("{00000000-0000-0000-C000-000000000046}")
	if err != nil {
		t.Fatalf("ClassIdFromGuidString failed: %v", err)
	}

	if expected != IID_IUnknown {
		t.Fatalf("ClassIdFromGuidString() = %v, want %v", expected, IID_IUnknown)
	}
}

func TestCreateInstance_FromProgramID(t *testing.T) {
	_, err := InitializeMultithreaded()
	if err != nil {
		t.Fatalf("InitializeMultithreaded failed: %v", err)
	}
	defer Uninitialize()

	classID, err := ClassIdFromProgramId(testCOMProgramID)
	if err != nil {
		t.Skipf("%s not available: %v", testCOMProgramID, err)
	}

	unknown, err := CreateInstance[IUnknown](classID, IID_IUnknown)
	if err != nil {
		t.Fatalf("CreateInstance failed: %v", err)
	}
	if unknown == nil {
		t.Fatal("CreateInstance returned nil")
	}

	unknown.Release()
}

func TestError(t *testing.T) {
	_, err := ClassIdFromProgramId("INTERFACE-NOT-FOUND")
	if err == nil {
		t.Fatal("ClassIdFromProgramId should fail")
	}

	if !errors.Is(err, InvalidClassIdForProgramId) && !errors.Is(err, UnknownLookupArgument) {
		t.Fatalf("ClassIdFromProgramId error = %v, want lookup error", err)
	}
}

func TestGetUserDefaultLCID(t *testing.T) {
	lcid := GetUserDefaultLCID()
	if lcid == 0 {
		t.Fatal("GetUserDefaultLCID returned 0")
	}
}

func TestClassIdToStringIUnknown(t *testing.T) {
	got, err := ClassIdToString(IID_IUnknown)
	if err != nil {
		t.Fatalf("ClassIdToString failed: %v", err)
	}

	if got != "{00000000-0000-0000-C000-000000000046}" {
		t.Fatalf("ClassIdToString() = %q, want IUnknown GUID", got)
	}
}

func TestInterfaceIdToStringIDispatch(t *testing.T) {
	got, err := InterfaceIdToString(IID_IDispatch)
	if err != nil {
		t.Fatalf("InterfaceIdToString failed: %v", err)
	}

	if got != "{00020400-0000-0000-C000-000000000046}" {
		t.Fatalf("InterfaceIdToString() = %q, want IDispatch GUID", got)
	}
}

func TestClassIdFromStringWithGUID(t *testing.T) {
	got, err := ClassIdFromString("{00000000-0000-0000-C000-000000000046}")
	if err != nil {
		t.Fatalf("ClassIdFromString failed: %v", err)
	}

	if got != windows.GUID(IID_IUnknown) {
		t.Fatalf("ClassIdFromString() = %v, want %v", got, IID_IUnknown)
	}
}
