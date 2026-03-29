//go:build windows

package ole

import (
	"errors"
	"fmt"
	"testing"

	"golang.org/x/sys/windows"
)

var iidDWebBrowserEvents2Test, _ = windows.GUIDFromString("{34A715A0-6587-11D0-924A-0020AFC7AC4D}")

func acquireWebBrowserConnectionPoint(t *testing.T) (*IUnknown, *IConnectionPointContainer, *IConnectionPoint, func()) {
	t.Helper()

	if _, err := Initialize(Multithreaded); err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}

	cleanup := func() {
		Uninitialize()
	}

	candidates := []string{
		"InternetExplorer.Application",
		"Shell.Explorer.2",
		"Shell.Explorer",
	}

	var failures []string
	for _, programID := range candidates {
		classID, err := ClassIdFromString(programID)
		if err != nil {
			failures = append(failures, fmt.Sprintf("%s: class lookup failed: %v", programID, err))
			continue
		}

		unknownPtr, err := CreateInstance[*IUnknown](classID, IID_IUnknown)
		if err != nil {
			failures = append(failures, fmt.Sprintf("%s: create failed: %v", programID, err))
			continue
		}
		unknown := *unknownPtr

		dispatch, dispatchErr := QueryIDispatchFromIUnknown(unknown)
		if dispatchErr != nil {
			dispatch = nil
		}

		container, err := QueryIConnectionPointContainerFromIUnknown(unknown)
		if err != nil {
			if dispatch != nil {
				_, _ = dispatch.CallMethod("Quit")
				dispatch.Release()
			}
			unknown.Release()
			failures = append(failures, fmt.Sprintf("%s: no connection point container: %v", programID, err))
			continue
		}

		point, err := container.FindConnectionPoint(iidDWebBrowserEvents2Test)
		if err != nil {
			container.Release()
			if dispatch != nil {
				_, _ = dispatch.CallMethod("Quit")
				dispatch.Release()
			}
			unknown.Release()
			failures = append(failures, fmt.Sprintf("%s: DWebBrowserEvents2 point missing: %v", programID, err))
			continue
		}

		return unknown, container, point, func() {
			point.Release()
			container.Release()
			if dispatch != nil {
				_, _ = dispatch.CallMethod("Quit")
				dispatch.Release()
			}
			unknown.Release()
			Uninitialize()
		}
	}

	cleanup()
	t.Skipf("no suitable COM object with IConnectionPoint found; tried: %v", failures)
	return nil, nil, nil, func() {}
}

func TestCOMObjectImplementsIConnectionPointForDWebBrowserEvents2(t *testing.T) {
	_, container, point, cleanup := acquireWebBrowserConnectionPoint(t)
	defer cleanup()

	interfaceID, err := point.GetConnectionInterface()
	if err != nil {
		t.Fatalf("GetConnectionInterface failed: %v", err)
	}
	if interfaceID != iidDWebBrowserEvents2Test {
		t.Fatalf("GetConnectionInterface = %v, want %v", interfaceID, iidDWebBrowserEvents2Test)
	}

	containerFromPoint, err := point.GetConnectionPointContainer()
	if err != nil {
		t.Fatalf("GetConnectionPointContainer failed: %v", err)
	}
	defer containerFromPoint.Release()

	pointFromReturnedContainer, err := containerFromPoint.FindConnectionPoint(iidDWebBrowserEvents2Test)
	if err != nil {
		t.Fatalf("FindConnectionPoint on returned container failed: %v", err)
	}
	defer pointFromReturnedContainer.Release()

	interfaceIDFromReturnedPoint, err := pointFromReturnedContainer.GetConnectionInterface()
	if err != nil {
		t.Fatalf("GetConnectionInterface on returned point failed: %v", err)
	}
	if interfaceIDFromReturnedPoint != iidDWebBrowserEvents2Test {
		t.Fatalf("returned point interface = %v, want %v", interfaceIDFromReturnedPoint, iidDWebBrowserEvents2Test)
	}

	if container == nil {
		t.Fatal("original connection point container is nil")
	}
}

func TestCOMObjectIConnectionPointEnumConnectionsStartsEmpty(t *testing.T) {
	_, _, point, cleanup := acquireWebBrowserConnectionPoint(t)
	defer cleanup()

	enum, err := point.EnumConnections()
	if err != nil {
		t.Fatalf("EnumConnections failed: %v", err)
	}
	defer enum.Release()

	if !enum.Reset() {
		t.Fatal("Reset() = false, want true for a fresh enumerator")
	}

	items := enum.Next(1)
	if len(items) != 0 {
		t.Fatalf("EnumConnections.Next(1) len = %d, want 0 for a fresh object with no advised sinks", len(items))
	}
}

func TestCOMObjectIConnectionPointAdviseRejectsIUnknownSink(t *testing.T) {
	unknown, _, point, cleanup := acquireWebBrowserConnectionPoint(t)
	defer cleanup()

	sink := IsIUnknown(unknown)
	cookie, err := point.Advise(&sink)
	if err == nil {
		if cookie != 0 {
			_ = point.Unadvise(cookie)
		}
		t.Fatal("Advise unexpectedly succeeded with plain IUnknown sink")
	}

	if errors.Is(err, ComInterfaceNotImplementedError) {
		return
	}

	var errno windows.Errno
	if !errors.As(err, &errno) {
		t.Fatalf("Advise error = %v, want COM interface failure", err)
	}
}
