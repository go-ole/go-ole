package ole

import (
	"testing"
	"fmt"
)

func TestComSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()
	
	coInitialize()
	CoUninitialize()
}

func TestComPublicSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()
	
	CoInitialize(0)
	CoUninitialize()
}

func TestComPublicSetupAndShutDown_WithValue(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()
	
	CoInitialize(5)
	CoUninitialize()
}

func TestComExSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()
	
	coInitializeEx(COINIT_MULTITHREADED)
	CoUninitialize()
}

func TestComPublicExSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()
	
	CoInitializeEx(0, COINIT_MULTITHREADED)
	CoUninitialize()
}

func TestComPublicExSetupAndShutDown_WithValue(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()
	
	CoInitializeEx(5, COINIT_MULTITHREADED)
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
	actual, err := CLSIDFromProgID("WMPNSSCI.NSSManager")
	CoUninitialize()

	if ! IsEqualGUID(expected, actual) {
		t.Log(err)
		t.Log(fmt.Sprintf("Actual GUID: %+v\n", actual))
		t.Fail()
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
	actual, err := CLSIDFromString("{92498132-4D1A-4297-9B78-9E2E4BA99C07}")
	CoUninitialize()

	if ! IsEqualGUID(expected, actual) {
		t.Log(err)
		t.Log(fmt.Sprintf("Actual GUID: %+v\n", actual))
		t.Fail()
	}
}