//go:build windows

package ole

import (
	"testing"
)

func TestIUnknown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Error(r)
		}
	}()

	var err error

	err = Initialize(0)
	if err != nil {
		t.Fatal(err)
	}

	defer Uninitialize()

	var unknown *IUnknown

	unknown, err = CreateInstance(CLSID_COMEchoTestObject, IID_IUnknown)
	defer unknown.Release()
	if err != nil {
		t.Fatal(err)
		return
	}
}
