//go:build windows
// +build windows

package legacy

import (
	"github.com/go-ole/go-ole"
	"testing"
)

func TestIUnknown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Error(r)
		}
	}()

	var err error

	err = CoInitialize(0)
	if err != nil {
		t.Fatal(err)
	}

	defer CoUninitialize()

	var unknown *ole.IUnknown

	// oleutil.CreateObject()
	unknown, err = CreateInstance(CLSID_COMEchoTestObject, IID_IUnknown)
	if err != nil {
		t.Fatal(err)
		return
	}
	unknown.Release()
}
