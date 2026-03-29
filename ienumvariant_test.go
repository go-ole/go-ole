//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"testing"
)

func TestIEnumVariant_wmi(t *testing.T) {
	var err error
	var classID windows.GUID

	IID_ISWbemLocator := windows.GUID{Data1: 0x76a6415b, Data2: 0xcb41, Data3: 0x11d1, Data4: [8]byte{0x8b, 0x02, 0x00, 0x60, 0x08, 0x06, 0xd9, 0xb6}}

	_, err = InitializeMultithreaded()
	if err != nil {
		t.Fatalf("Initialize error: %v", err)
	}
	defer Uninitialize()
	RegisterVariantConverters()

	classID, err = ClassIdFromString("WbemScripting.SWbemLocator")
	if err != nil {
		t.Fatalf("ClassIdFromString WbemScripting.SWbemLocator returned with %v", err)
	}

	unknownPtr, err := CreateInstance[*IUnknown](classID, IID_IUnknown)
	if err != nil {
		t.Fatalf("CreateInstance WbemScripting.SWbemLocator returned with %v", err)
	}
	unknown := *unknownPtr
	if unknown == nil {
		t.Fatal("CreateInstance WbemScripting.SWbemLocator returned nil")
	}
	defer unknown.Release()

	dispatch, err := QueryInterfaceOnIUnknown[IDispatch](unknown, IID_ISWbemLocator)
	if err != nil {
		t.Fatalf("QueryInterfaceOnIUnknown returned with %v", err)
	}
	defer dispatch.Release()

	wbemServices, err := dispatch.CallMethod("ConnectServer")
	if err != nil {
		t.Fatalf("ConnectServer failed with %v", err)
	}
	defer wbemServices.Clear()

	wbemServicesDispatch := VariantToComObject[*IDispatch](wbemServices)
	objectset, err := (*wbemServicesDispatch).CallMethod("ExecQuery", StringToBStrVariant("SELECT * FROM WIN32_Process"))
	if err != nil {
		t.Fatalf("ExecQuery failed with %v", err)
	}
	defer objectset.Clear()

	objectsetDispatch := VariantToComObject[*IDispatch](objectset)
	enumProperty, err := (*objectsetDispatch).GetProperty("_NewEnum")
	if err != nil {
		t.Fatalf("Get _NewEnum property failed with %v", err)
	}
	defer enumProperty.Clear()

	enumUnknown := VariantToComObject[*IUnknown](enumProperty)
	enum, err := QueryIEnumVariantFromIUnknown(*enumUnknown)
	if err != nil {
		t.Fatalf("QueryIEnumVariantFromIUnknown returned with %v", err)
	}
	if enum == nil {
		t.Fatal("Enum is nil")
	}
	defer enum.Release()

	for items := enum.Next(1); len(items) > 0; items = enum.Next(1) {
		itemDispatch := VariantToComObject[*IDispatch](items[0])
		defer (*itemDispatch).Release()

		nameVariant, err := (*itemDispatch).GetProperty("Name")
		if err != nil {
			t.Fatalf("Get Name property failed with %v", err)
		}
		defer nameVariant.Clear()

		name := UnwrapVariant[string](nameVariant)
		t.Logf("Got %v", name)
	}
}
