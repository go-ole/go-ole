// +build windows

package ole

import "testing"

func TestIEnumVariant_wmi(t *testing.T) {
	var err error
	var classID *GUID
	var displayID int32

	err = CoInitialize(0)
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
	IID_IEnumVariant := &GUID{0x00020404, 0x0000, 0x0000, [8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

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

	displayID, err = GetSingleIDOfName(objectset_dispatch, "_NewEnum")
	if err != nil {
		t.Errorf("_NewEnum display id failed with %v", err)
	}

	variant, err := objectset_dispatch.Invoke(displayID, DISPATCH_PROPERTYGET)
	if err != nil {
		t.Errorf("Get _NewEnum property failed with %v", err)
	}
	defer variant.Clear()

	object2 := variant.ToIUnknown()
	if object2 == nil {
		t.Errorf("Object 2 iunknown is nil returned with %v", err)
	}

	enum, err := object2.IEnumVARIANT(IID_IEnumVariant)
	if err != nil {
		t.Errorf("IEnumVARIANT() returned with %v", err)
	}
	if enum == nil {
		t.Error("Enum is nil")
		t.FailNow()
	}

	var tmp_dispatch *IDispatch

	for tmp, length, err := enum.Next(1); length > 0; tmp, length, err = enum.Next(1) {
		if err != nil {
			t.Errorf("Next() returned with %v", err)
		}
		tmp_dispatch = tmp.ToIDispatch()
		defer tmp_dispatch.Release()

		displayID, err = GetSingleIDOfName(tmp_dispatch, "Properties_")
		if err != nil {
			t.Errorf("Properties_ display id failed with %v", err)
		}

		props, err := tmp_dispatch.Invoke(displayID, DISPATCH_PROPERTYGET)
		if err != nil {
			t.Errorf("Get Properties_ property failed with %v", err)
		}
		defer props.Clear()

		props_dispatch := props.ToIDispatch()

		displayID, err = GetSingleIDOfName(props_dispatch, "_NewEnum")
		if err != nil {
			t.Errorf("_NewEnum display id failed with %v", err)
		}

		props_enum_property, err := props_dispatch.Invoke(displayID, DISPATCH_PROPERTYGET)
		if err != nil {
			t.Errorf("Get _NewEnum property failed with %v", err)
		}
		defer props_enum_property.Clear()

		_, err = props_enum_property.ToIUnknown().IEnumVARIANT(IID_IEnumVariant)
		if err != nil {
			t.Errorf("IEnumVARIANT failed with %v", err)
		}

		displayID, err = GetSingleIDOfName(tmp_dispatch, "Name")
		if err != nil {
			t.Errorf("Name display id failed with %v", err)
		}

		class_variant, err := tmp_dispatch.Invoke(displayID, DISPATCH_PROPERTYGET)
		if err != nil {
			t.Errorf("Get Name property failed with %v", err)
		}
		defer class_variant.Clear()

		class_name := class_variant.ToString()
		t.Logf("Got %v", class_name)
	}
}
