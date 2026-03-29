//go:build windows

package ole

import (
	"fmt"
	"strings"
	"testing"
)

type automationIntegrationCandidate struct {
	programID    string
	probeName    string
	supportsQuit bool
}

func acquireAutomationDispatch(t *testing.T) (*IUnknown, *IDispatch, automationIntegrationCandidate, func()) {
	t.Helper()

	if _, err := Initialize(Multithreaded); err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}

	candidates := []automationIntegrationCandidate{
		{programID: "InternetExplorer.Application", probeName: "Visible", supportsQuit: true},
		{programID: "Shell.Explorer.2", probeName: "Visible", supportsQuit: true},
		{programID: "Shell.Explorer", probeName: "Visible", supportsQuit: true},
		{programID: "Shell.Application", probeName: "Application", supportsQuit: false},
	}

	var failures []string
	for _, candidate := range candidates {
		classID, err := ClassIdFromString(candidate.programID)
		if err != nil {
			failures = append(failures, fmt.Sprintf("%s: class lookup failed: %v", candidate.programID, err))
			continue
		}

		unknownPtr, err := CreateInstance[*IUnknown](classID, IID_IUnknown)
		if err != nil {
			failures = append(failures, fmt.Sprintf("%s: create failed: %v", candidate.programID, err))
			continue
		}
		unknown := *unknownPtr

		dispatch, err := QueryIDispatchFromIUnknown(unknown)
		if err != nil {
			unknown.Release()
			failures = append(failures, fmt.Sprintf("%s: no IDispatch: %v", candidate.programID, err))
			continue
		}

		return unknown, dispatch, candidate, func() {
			if candidate.supportsQuit {
				_, _ = dispatch.CallMethod("Quit")
			}
			dispatch.Release()
			unknown.Release()
			Uninitialize()
		}
	}

	Uninitialize()
	t.Skipf("no suitable automation object found; tried: %s", strings.Join(failures, "; "))
	return nil, nil, automationIntegrationCandidate{}, func() {}
}

func acquireProvideClassInfo(t *testing.T) (*IUnknown, *IDispatch, *IProvideClassInfo, func()) {
	t.Helper()

	if _, err := Initialize(Multithreaded); err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}

	candidates := []automationIntegrationCandidate{
		{programID: "InternetExplorer.Application", probeName: "Visible", supportsQuit: true},
		{programID: "Shell.Explorer.2", probeName: "Visible", supportsQuit: true},
		{programID: "Shell.Explorer", probeName: "Visible", supportsQuit: true},
		{programID: "Shell.Application", probeName: "Application", supportsQuit: false},
	}

	var failures []string
	for _, candidate := range candidates {
		classID, err := ClassIdFromString(candidate.programID)
		if err != nil {
			failures = append(failures, fmt.Sprintf("%s: class lookup failed: %v", candidate.programID, err))
			continue
		}

		unknownPtr, err := CreateInstance[*IUnknown](classID, IID_IUnknown)
		if err != nil {
			failures = append(failures, fmt.Sprintf("%s: create failed: %v", candidate.programID, err))
			continue
		}
		unknown := *unknownPtr

		dispatch, err := QueryIDispatchFromIUnknown(unknown)
		if err != nil {
			unknown.Release()
			failures = append(failures, fmt.Sprintf("%s: no IDispatch: %v", candidate.programID, err))
			continue
		}

		provideClassInfo, err := QueryInterfaceOnIUnknown[IProvideClassInfo](unknown, IID_IProvideClassInfo)
		if err != nil {
			if candidate.supportsQuit {
				_, _ = dispatch.CallMethod("Quit")
			}
			dispatch.Release()
			unknown.Release()
			failures = append(failures, fmt.Sprintf("%s: no IProvideClassInfo: %v", candidate.programID, err))
			continue
		}

		return unknown, dispatch, provideClassInfo, func() {
			provideClassInfo.Release()
			if candidate.supportsQuit {
				_, _ = dispatch.CallMethod("Quit")
			}
			dispatch.Release()
			unknown.Release()
			Uninitialize()
		}
	}

	Uninitialize()
	t.Skipf("no suitable automation object with IProvideClassInfo found; tried: %s", strings.Join(failures, "; "))
	return nil, nil, nil, func() {}
}

func TestCOMObjectImplementsIDispatch(t *testing.T) {
	_, dispatch, candidate, cleanup := acquireAutomationDispatch(t)
	defer cleanup()

	if dispatch == nil {
		t.Fatal("IDispatch is nil")
	}

	if _, err := dispatch.GetSingleIDOfName(candidate.probeName); err != nil {
		t.Fatalf("GetSingleIDOfName(%q) failed: %v", candidate.probeName, err)
	}

	if !dispatch.HasTypeInfo() {
		t.Fatalf("%s.HasTypeInfo() = false, want true", candidate.programID)
	}

	typeInfo := dispatch.GetTypeInfo()
	if typeInfo == nil {
		t.Fatalf("%s.GetTypeInfo() = nil, want non-nil", candidate.programID)
	}
	typeInfo.Release()
}

func TestCOMObjectImplementsIProvideClassInfo(t *testing.T) {
	_, _, provideClassInfo, cleanup := acquireProvideClassInfo(t)
	defer cleanup()

	typeInfo, err := provideClassInfo.GetClassInfo()
	if err != nil {
		t.Fatalf("GetClassInfo failed: %v", err)
	}
	if typeInfo == nil {
		t.Fatal("GetClassInfo() = nil, want non-nil")
	}
	defer typeInfo.Release()
}

func TestCOMObjectImplementsITypeInfo(t *testing.T) {
	_, dispatch, _, cleanup := acquireAutomationDispatch(t)
	defer cleanup()

	typeInfo := dispatch.GetTypeInfo()
	if typeInfo == nil {
		t.Fatal("GetTypeInfo() = nil, want non-nil")
	}
	defer typeInfo.Release()

	typeAttr, err := typeInfo.GetTypeAttr()
	if err != nil {
		t.Fatalf("GetTypeAttr failed: %v", err)
	}
	if typeAttr == nil {
		t.Fatal("GetTypeAttr() = nil, want non-nil")
	}
	typeInfo.ReleaseTypeAttr(typeAttr)

	typeComp, err := typeInfo.GetTypeComp()
	if err != nil {
		t.Fatalf("GetTypeComp failed: %v", err)
	}
	if typeComp == nil {
		t.Fatal("GetTypeComp() = nil, want non-nil")
	}
	typeComp.Release()

	name, _, _, _, err := typeInfo.GetDocumentation(0)
	if err != nil {
		t.Fatalf("GetDocumentation failed: %v", err)
	}
	if name == "" {
		t.Fatal("GetDocumentation() name = empty, want non-empty")
	}

	typeLib, index, err := typeInfo.GetContainingTypeLib()
	if err != nil {
		t.Fatalf("GetContainingTypeLib failed: %v", err)
	}
	if typeLib == nil {
		t.Fatal("GetContainingTypeLib() typelib = nil, want non-nil")
	}
	defer typeLib.Release()
	if index > 1<<20 {
		t.Fatalf("GetContainingTypeLib() index = %d, looks invalid", index)
	}
}
