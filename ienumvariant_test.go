// +build windows

package ole

import "testing"

func TestIEnumVariant_wmi(t *testing.T) {
	defer func() {
		r := recover()
		if r != nil {
			t.Error(r)
		}
	}()

	var err error
	var classID *GUID
	var displayID int32

	err = CoInitializeEx(0, COINIT_APARTMENTTHREADED)
	if err != nil {
		t.Errorf("Initialize error: %v", err)
	}
	defer CoUninitialize()

	classID, err = CLSIDFromProgID("WbemScripting.SWbemLocator")
	if err != nil {
		classID, err = CLSIDFromString("WbemScripting.SWbemLocator")
		if err != nil {
			t.Errorf("CreateObject WbemScripting.SWbemLocator returned with %v", err)
		}
	}

	comserver, err := CreateInstance(classID, IID_IUnknown)
	if err != nil {
		t.Errorf("CreateInstance WbemScripting.SWbemLocator returned with %v", err)
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

	displayID, err = GetSingleIDOfName(dispatch, "ConnectServer")
	if err != nil {
		t.Errorf("ConnectServer display id failed with %v", err)
	}

	wbemServices, err := dispatch.Invoke(displayID, DISPATCH_METHOD)
	if err != nil {
		t.Errorf("ExecQuery failed with %v", err)
	}
	defer wbemServices.Clear()

	wbemServices_dispatch := wbemServices.ToIDispatch()
	defer wbemServices_dispatch.Release()

	displayID, err = GetSingleIDOfName(wbemServices_dispatch, "ExecQuery")
	if err != nil {
		t.Errorf("ExecQuery display id failed with %v", err)
	}

	objectset, err := wbemServices_dispatch.Invoke(displayID, DISPATCH_METHOD, "SELECT * FROM WIN32_Process")
	if err != nil {
		t.Errorf("ExecQuery failed with %v", err)
	}
	defer objectset.Clear()

	objectset_dispatch := objectset.ToIDispatch()
	defer objectset_dispatch.Release()

	displayID, err = GetSingleIDOfName(objectset_dispatch, "_NewEnum")
	if err != nil {
		t.Errorf("_NewEnum display id failed with %v", err)
	}

	variant, err := objectset_dispatch.Invoke(displayID, DISPATCH_PROPERTYGET)
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
