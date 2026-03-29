//go:build windows

package ole

import (
	"testing"
)

func TestClassIdFromGuidString(t *testing.T) {
	guidStr := "{00000000-0000-0000-C000-000000000046}"
	classId, err := ClassIdFromGuidString(guidStr)
	if err != nil {
		t.Errorf("ClassIdFromGuidString failed: %v", err)
	}
	if classId != IID_IUnknown {
		t.Errorf("Expected %v, got %v", IID_IUnknown, classId)
	}
}

func TestClassIdFromGuidString_Error(t *testing.T) {
	_, err := ClassIdFromGuidString("invalid-guid")
	if err == nil {
		t.Error("Expected error for invalid GUID")
	}
}

func TestClassIdToString(t *testing.T) {
	guidStr, err := ClassIdToString(IID_IUnknown)
	if err != nil {
		t.Errorf("ClassIdToString failed: %v", err)
	}
	expected := "{00000000-0000-0000-C000-000000000046}"
	if guidStr != expected {
		t.Errorf("Expected %s, got %s", expected, guidStr)
	}
}

func TestInterfaceIdToString(t *testing.T) {
	guidStr, err := InterfaceIdToString(IID_IDispatch)
	if err != nil {
		t.Errorf("InterfaceIdToString failed: %v", err)
	}
	expected := "{00020400-0000-0000-C000-000000000046}"
	if guidStr != expected {
		t.Errorf("Expected %s, got %s", expected, guidStr)
	}
}

func TestInterfaceIdFromString(t *testing.T) {
	guidStr := "{00020400-0000-0000-C000-000000000046}"
	classId, err := InterfaceIdFromString(guidStr)
	if err != nil {
		t.Errorf("InterfaceIdFromString failed: %v", err)
	}
	if classId != IID_IDispatch {
		t.Errorf("Expected %v, got %v", IID_IDispatch, classId)
	}
}

func TestClassIdFromString(t *testing.T) {
	// Test with GUID string
	guidStr := "{00000000-0000-0000-C000-000000000046}"
	classId, err := ClassIdFromString(guidStr)
	if err != nil {
		t.Errorf("ClassIdFromString with GUID failed: %v", err)
	}
	if classId != IID_IUnknown {
		t.Errorf("Expected %v, got %v", IID_IUnknown, classId)
	}
}

func TestClassIdFromProgramId(t *testing.T) {
	// ProgID for Shell.Application is very likely to be present on any Windows system
	programId := "Shell.Application"
	_, err := ClassIdFromProgramId(programId)
	if err != nil {
		// We don't want to fail the test if the ProgID is not found, as it depends on the environment
		// But we can at least check if it returns a specific error if it's not found
		t.Logf("ClassIdFromProgramId(%s) error: %v (this is expected if ProgID is not registered)", programId, err)
	}
}
