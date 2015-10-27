// +build windows

package ole

import (
	"testing"

	"github.com/go-ole/go-ole/oleutil"
)

func TestIEnumVariant_wmi(t *testing.T) {
	defer func() {
		r := recover()
		if r != nil {
			t.Error(r)
		}
	}()

	err := CoInitializeEx(0, COINIT_APARTMENTTHREADED)
	if err != nil {
		t.Errorf("Initialize error: %v", err)
	}
	defer CoUninitialize()

	comserver, err := oleutil.CreateObject("WbemScripting.SWbemLocator")
	if err != nil {
		t.Errorf("CreateObject WbemScripting.SWbemLocator returned with %v", err)
	}
	if comserver == nil {
		t.Error("CreateObject WbemScripting.SWbemLocator not an object")
	}
	defer comserver.Release()

	IID_ISWbemLocator := &GUID{0x76a6415b, 0xcb41, 0x11d1, [8]byte{0x8b, 0x02, 0x00, 0x60, 0x08, 0x06, 0xd9, 0xb6}}
	IID_IEnumVariant := &GUID{0x027947E1, 0xD731, 0x11CE, [8]byte{0xA3, 0x57, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01}}

	dispatch, err := comserver.QueryInterface(IID_ISWbemLocator)
	if err != nil {
		t.Errorf("context.iunknown.QueryInterface returned with %v", err)
	}
	defer dispatch.Release()

	wbemServices, err := oleutil.CallMethod(dispatch, "ConnectServer")
	if err != nil {
		t.Errorf("ExecQuery failed with %v", err)
	}
	defer wbemServices.Clear()

	wbemServices_dispatch := wbemServices.ToIDispatch()
	defer wbemServices_dispatch.Release()

	objectset, err := oleutil.CallMethod(wbemServices_dispatch, "ExecQuery", "SELECT * FROM WIN32_Process")
	if err != nil {
		t.Errorf("ExecQuery failed with %v", err)
	}
	defer objectset.Clear()

	objectset_dispatch := objectset.ToIDispatch()
	defer objectset_dispatch.Release()

	variant, err := oleutil.GetProperty(objectset_dispatch, "_NewEnum")
	if err != nil {
		t.Errorf("Get _NewEnum property failed with %v", err)
	}
	defer variant.Clear()

	object2, err := variant.ToIUnknown().IEnumVARIANT(IID_IEnumVariant)
	if err != nil {
		t.Errorf("enum.Next() returned with %v", err)
	}
	defer object2.Release()

	a, l, err := object2.Next(1)
	if err != nil {
		t.Errorf("enum.Next() returned with %v", err)
	}

	t.Log(&a)
	t.Log(&l)
}
