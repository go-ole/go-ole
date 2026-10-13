// +build windows

package ole

import (
	"fmt"
	"runtime"
	"runtime/debug"
	"strings"
	"sync"
	"testing"
	"time"
)

func TestComSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := coInitialize()
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	CoUninitialize()
}

func TestComPublicSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := CoInitialize(0)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	CoUninitialize()
}

func TestComPublicSetupAndShutDown_WithValue(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := CoInitialize(5)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	CoUninitialize()
}

func TestComExSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := coInitializeEx(COINIT_MULTITHREADED)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	CoUninitialize()
}

func TestComPublicExSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := CoInitializeEx(0, COINIT_MULTITHREADED)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	CoUninitialize()
}

func TestComPublicExSetupAndShutDown_WithValue(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := CoInitializeEx(5, COINIT_MULTITHREADED)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	CoUninitialize()
}

func TestClsidFromProgID_WindowsMediaNSSManager(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	expected := &GUID{0x92498132, 0x4D1A, 0x4297, [8]byte{0x9B, 0x78, 0x9E, 0x2E, 0x4B, 0xA9, 0x9C, 0x07}}

	coInitialize()
	defer CoUninitialize()
	actual, err := CLSIDFromProgID("WMPNSSCI.NSSManager")
	if err == nil {
		if !IsEqualGUID(expected, actual) {
			t.Log(err)
			t.Log(fmt.Sprintf("Actual GUID: %+v\n", actual))
			t.Fail()
		}
	}
}

func TestClsidFromString_WindowsMediaNSSManager(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	expected := &GUID{0x92498132, 0x4D1A, 0x4297, [8]byte{0x9B, 0x78, 0x9E, 0x2E, 0x4B, 0xA9, 0x9C, 0x07}}

	coInitialize()
	defer CoUninitialize()
	actual, err := CLSIDFromString("{92498132-4D1A-4297-9B78-9E2E4BA99C07}")

	if !IsEqualGUID(expected, actual) {
		t.Log(err)
		t.Log(fmt.Sprintf("Actual GUID: %+v\n", actual))
		t.Fail()
	}
}

func TestCreateInstance_WindowsMediaNSSManager(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	expected := &GUID{0x92498132, 0x4D1A, 0x4297, [8]byte{0x9B, 0x78, 0x9E, 0x2E, 0x4B, 0xA9, 0x9C, 0x07}}

	coInitialize()
	defer CoUninitialize()
	actual, err := CLSIDFromProgID("WMPNSSCI.NSSManager")

	if err == nil {
		if !IsEqualGUID(expected, actual) {
			t.Log(err)
			t.Log(fmt.Sprintf("Actual GUID: %+v\n", actual))
			t.Fail()
		}

		unknown, err := CreateInstance(actual, IID_IUnknown)
		if err != nil {
			t.Log(err)
			t.Fail()
		}
		unknown.Release()
	}
}

func TestError(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	coInitialize()
	defer CoUninitialize()
	_, err := CLSIDFromProgID("INTERFACE-NOT-FOUND")
	if err == nil {
		t.Fatalf("should be fail: %v", err)
	}

	switch vt := err.(type) {
	case *OleError:
	default:
		t.Fatalf("should be *ole.OleError %t", vt)
	}
}

// Scripting.FileSystemObject ships with every Windows release, unlike WMPNSSCI.NSSManager.
const fileSystemObjectProgID = "Scripting.FileSystemObject"
const fileSystemObjectCLSID = "{0D43FE01-F093-11CF-8940-00A0C9054228}"
const iDispatchIIDString = "{00020400-0000-0000-C000-000000000046}"

func TestClsidFromProgID_FileSystemObject(t *testing.T) {
	coInitialize()
	defer CoUninitialize()

	actual, err := CLSIDFromProgID(fileSystemObjectProgID)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.EqualFold(actual.String(), fileSystemObjectCLSID) {
		t.Fatalf("expected %s, got %s", fileSystemObjectCLSID, actual)
	}
}

func TestIIDFromString_IDispatch(t *testing.T) {
	coInitialize()
	defer CoUninitialize()

	actual, err := IIDFromString(iDispatchIIDString)
	if err != nil {
		t.Fatal(err)
	}
	if !IsEqualGUID(IID_IDispatch, actual) {
		t.Fatalf("expected %s, got %s", IID_IDispatch, actual)
	}
}

func TestIIDFromString_Invalid(t *testing.T) {
	coInitialize()
	defer CoUninitialize()

	if _, err := IIDFromString("not-a-guid"); err == nil {
		t.Fatal("expected an error for a malformed IID string")
	}
}

// If a UTF-16 buffer is collected while ole32 still reads it, this faults inside combase.
func TestStringToGUIDConversionsUnderGCPressure(t *testing.T) {
	if testing.Short() {
		t.Skip("skipping GC pressure test in short mode")
	}

	const workerCount = 8
	const runDuration = 5 * time.Second

	previousGCPercent := debug.SetGCPercent(1)
	defer debug.SetGCPercent(previousGCPercent)

	stop := make(chan struct{})
	failures := make(chan string, workerCount)
	var wg sync.WaitGroup

	for i := 0; i < runtime.NumCPU(); i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for {
				select {
				case <-stop:
					return
				default:
					_ = make([]uint16, 512*1024)
				}
			}
		}()
	}

	for i := 0; i < workerCount; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			runtime.LockOSThread()
			defer runtime.UnlockOSThread()

			// Other tests leave COM initialised on pooled threads: accept S_FALSE and RPC_E_CHANGED_MODE.
			err := CoInitializeEx(0, COINIT_APARTMENTTHREADED)
			if oleErr, ok := err.(*OleError); err == nil || (ok && oleErr.Code() == 0x00000001) {
				defer CoUninitialize()
			} else if !ok || oleErr.Code() != 0x80010106 {
				failures <- fmt.Sprintf("CoInitializeEx: %v", err)
				return
			}

			for {
				select {
				case <-stop:
					return
				default:
				}

				clsid, err := CLSIDFromProgID(fileSystemObjectProgID)
				if err != nil || !strings.EqualFold(clsid.String(), fileSystemObjectCLSID) {
					failures <- fmt.Sprintf("CLSIDFromProgID: %v %v", clsid, err)
					return
				}
				clsid, err = CLSIDFromString(fileSystemObjectCLSID)
				if err != nil || !strings.EqualFold(clsid.String(), fileSystemObjectCLSID) {
					failures <- fmt.Sprintf("CLSIDFromString: %v %v", clsid, err)
					return
				}
				iid, err := IIDFromString(iDispatchIIDString)
				if err != nil || !IsEqualGUID(IID_IDispatch, iid) {
					failures <- fmt.Sprintf("IIDFromString: %v %v", iid, err)
					return
				}
			}
		}()
	}

	time.Sleep(runDuration)
	close(stop)
	wg.Wait()
	close(failures)

	for failure := range failures {
		t.Error(failure)
	}
}
