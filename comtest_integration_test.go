//go:build windows

package ole

import (
	"testing"

	"golang.org/x/sys/windows"
)

// createTestDispatch creates an IDispatch from a test COM server class by ProgID.
// Skips the test if the COM server is not registered.
func createTestDispatch(t *testing.T, progID string) (*IUnknown, *IDispatch, func()) {
	t.Helper()

	if _, err := Initialize(Multithreaded); err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}

	classID, err := ClassIdFromProgramId(progID)
	if err != nil {
		Uninitialize()
		t.Skipf("COM server %s not registered: %v", progID, err)
	}

	unknown, err := CreateInstance[IUnknown](classID, IID_IUnknown)
	if err != nil {
		Uninitialize()
		t.Skipf("CreateInstance(%s) failed: %v", progID, err)
	}

	dispatch, err := QueryIDispatchFromIUnknown(unknown)
	if err != nil {
		unknown.Release()
		Uninitialize()
		t.Fatalf("QueryIDispatchFromIUnknown(%s) failed: %v", progID, err)
	}

	return unknown, dispatch, func() {
		dispatch.Release()
		unknown.Release()
		Uninitialize()
	}
}

// createTestUnknown creates an IUnknown from a test COM server class by CLSID.
// Skips the test if the COM server is not registered.
func createTestUnknown(t *testing.T, clsid windows.GUID) (*IUnknown, func()) {
	t.Helper()

	if _, err := Initialize(Multithreaded); err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}

	// Probe whether the test COM server is registered at all
	_, err := ClassIdFromProgramId("TestCOM.String")
	if err != nil {
		Uninitialize()
		t.Skipf("Test COM server not registered: %v", err)
	}

	unknown, err := CreateInstance[IUnknown](clsid, IID_IUnknown)
	if err != nil {
		Uninitialize()
		t.Fatalf("CreateInstance failed: %v", err)
	}

	return unknown, func() {
		unknown.Release()
		Uninitialize()
	}
}

// TestCOMTestServer_CreateInstance_Types verifies that each type class
// in the test COM server can be instantiated via its ProgId.
func TestCOMTestServer_CreateInstance_Types(t *testing.T) {
	typeClasses := []struct {
		name   string
		progID string
	}{
		{"Int8", "TestCOM.Int8"},
		{"Int16", "TestCOM.Int16"},
		{"Int32", "TestCOM.Int32"},
		{"Int64", "TestCOM.Int64"},
		{"Float32", "TestCOM.Float32"},
		{"Float64", "TestCOM.Float64"},
		{"String", "TestCOM.String"},
		{"Boolean", "TestCOM.Boolean"},
		{"Currency", "TestCOM.Currency"},
		{"Date", "TestCOM.Date"},
		{"Decimal", "TestCOM.Decimal"},
		{"Error", "TestCOM.Error"},
		{"Variant", "TestCOM.Variant"},
		{"Unknown", "TestCOM.Unknown"},
		{"Dispatch", "TestCOM.Dispatch"},
		{"Empty", "TestCOM.Empty"},
		{"Clsid", "TestCOM.Clsid"},
		{"HResult", "TestCOM.HResult"},
		{"FileTime", "TestCOM.FileTime"},
	}

	_, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	for _, tc := range typeClasses {
		t.Run(tc.name, func(t *testing.T) {
			classID, err := ClassIdFromProgramId(tc.progID)
			if err != nil {
				t.Skipf("%s not registered: %v", tc.progID, err)
			}

			unknown, err := CreateInstance[IUnknown](classID, IID_IUnknown)
			if err != nil {
				t.Fatalf("CreateInstance failed: %v", err)
			}
			defer unknown.Release()

			dispatch, err := QueryIDispatchFromIUnknown(unknown)
			if err != nil {
				t.Fatalf("QueryIDispatchFromIUnknown failed: %v", err)
			}
			defer dispatch.Release()

			if !dispatch.HasTypeInfo() {
				t.Error("HasTypeInfo() = false, want true")
			}
		})
	}
}

// TestCOMTestServer_StringEcho tests string round-trip through CallMethod and GetProperty/PutProperty.
func TestCOMTestServer_StringEcho(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.String")
	defer cleanup()

	RegisterVariantConverters()

	t.Run("EchoString", func(t *testing.T) {
		param, err := WrapVariant("hello world")
		if err != nil {
			t.Fatalf("WrapVariant failed: %v", err)
		}
		defer param.Clear()

		result, err := dispatch.CallMethod("EchoString", param)
		if err != nil {
			t.Fatalf("CallMethod(EchoString) failed: %v", err)
		}
		defer result.Clear()

		got := UnwrapVariant[string](result)
		if got != "hello world" {
			t.Fatalf("EchoString() = %q, want %q", got, "hello world")
		}
	})

	t.Run("PutProperty_GetProperty", func(t *testing.T) {
		param, err := WrapVariant("test value")
		if err != nil {
			t.Fatalf("WrapVariant failed: %v", err)
		}
		defer param.Clear()

		_, err = dispatch.PutProperty("StringField", param)
		if err != nil {
			t.Fatalf("PutProperty(StringField) failed: %v", err)
		}

		result, err := dispatch.GetProperty("StringField")
		if err != nil {
			t.Fatalf("GetProperty(StringField) failed: %v", err)
		}
		defer result.Clear()

		got := UnwrapVariant[string](result)
		if got != "test value" {
			t.Fatalf("StringField = %q, want %q", got, "test value")
		}
	})

	t.Run("PutString_GetString", func(t *testing.T) {
		param, err := WrapVariant("method value")
		if err != nil {
			t.Fatalf("WrapVariant failed: %v", err)
		}
		defer param.Clear()

		putResult, err := dispatch.CallMethod("PutString", param)
		if err != nil {
			t.Fatalf("CallMethod(PutString) failed: %v", err)
		}
		defer putResult.Clear()

		length := UnwrapVariant[int32](putResult)
		if length != 12 {
			t.Fatalf("PutString() returned length = %d, want 12", length)
		}

		result, err := dispatch.CallMethod("GetString")
		if err != nil {
			t.Fatalf("CallMethod(GetString) failed: %v", err)
		}
		defer result.Clear()

		got := UnwrapVariant[string](result)
		if got != "method value" {
			t.Fatalf("GetString() = %q, want %q", got, "method value")
		}
	})
}

// TestCOMTestServer_Int32Echo tests int32 round-trip.
func TestCOMTestServer_Int32Echo(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Int32")
	defer cleanup()

	RegisterVariantConverters()

	t.Run("EchoInt32", func(t *testing.T) {
		param, err := WrapVariant(int32(42))
		if err != nil {
			t.Fatalf("WrapVariant failed: %v", err)
		}
		defer param.Clear()

		result, err := dispatch.CallMethod("EchoInt32", param)
		if err != nil {
			t.Fatalf("CallMethod(EchoInt32) failed: %v", err)
		}
		defer result.Clear()

		got := UnwrapVariant[int32](result)
		if got != 42 {
			t.Fatalf("EchoInt32(42) = %d, want 42", got)
		}
	})

	t.Run("PutProperty_GetProperty", func(t *testing.T) {
		param, err := WrapVariant(int32(99))
		if err != nil {
			t.Fatalf("WrapVariant failed: %v", err)
		}
		defer param.Clear()

		_, err = dispatch.PutProperty("Int32Field", param)
		if err != nil {
			t.Fatalf("PutProperty(Int32Field) failed: %v", err)
		}

		result, err := dispatch.GetProperty("Int32Field")
		if err != nil {
			t.Fatalf("GetProperty(Int32Field) failed: %v", err)
		}
		defer result.Clear()

		got := UnwrapVariant[int32](result)
		if got != 99 {
			t.Fatalf("Int32Field = %d, want 99", got)
		}
	})

	t.Run("NegativeValue", func(t *testing.T) {
		param, err := WrapVariant(int32(-2147483648))
		if err != nil {
			t.Fatalf("WrapVariant failed: %v", err)
		}
		defer param.Clear()

		result, err := dispatch.CallMethod("EchoInt32", param)
		if err != nil {
			t.Fatalf("CallMethod(EchoInt32) failed: %v", err)
		}
		defer result.Clear()

		got := UnwrapVariant[int32](result)
		if got != -2147483648 {
			t.Fatalf("EchoInt32(-2147483648) = %d, want -2147483648", got)
		}
	})
}

// TestCOMTestServer_Int64Echo tests int64 round-trip.
func TestCOMTestServer_Int64Echo(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Int64")
	defer cleanup()

	RegisterVariantConverters()

	param, err := WrapVariant(int64(9223372036854775807))
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	defer param.Clear()

	result, err := dispatch.CallMethod("EchoInt64", param)
	if err != nil {
		t.Fatalf("CallMethod(EchoInt64) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[int64](result)
	if got != 9223372036854775807 {
		t.Fatalf("EchoInt64(max) = %d, want 9223372036854775807", got)
	}
}

// TestCOMTestServer_Int16Echo tests int16 round-trip.
func TestCOMTestServer_Int16Echo(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Int16")
	defer cleanup()

	RegisterVariantConverters()

	param, err := WrapVariant(int16(32767))
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	defer param.Clear()

	result, err := dispatch.CallMethod("EchoInt16", param)
	if err != nil {
		t.Fatalf("CallMethod(EchoInt16) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[int16](result)
	if got != 32767 {
		t.Fatalf("EchoInt16(32767) = %d, want 32767", got)
	}
}

// TestCOMTestServer_Int8Echo tests int8 round-trip.
func TestCOMTestServer_Int8Echo(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Int8")
	defer cleanup()

	RegisterVariantConverters()

	param, err := WrapVariant(int8(127))
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	defer param.Clear()

	result, err := dispatch.CallMethod("EchoInt8", param)
	if err != nil {
		t.Fatalf("CallMethod(EchoInt8) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[int8](result)
	if got != 127 {
		t.Fatalf("EchoInt8(127) = %d, want 127", got)
	}
}

// TestCOMTestServer_Float32Echo tests float32 round-trip.
func TestCOMTestServer_Float32Echo(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Float32")
	defer cleanup()

	RegisterVariantConverters()

	param, err := WrapVariant(float32(3.14))
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	defer param.Clear()

	result, err := dispatch.CallMethod("EchoFloat32", param)
	if err != nil {
		t.Fatalf("CallMethod(EchoFloat32) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[float32](result)
	if got != float32(3.14) {
		t.Fatalf("EchoFloat32(3.14) = %f, want 3.14", got)
	}
}

// TestCOMTestServer_Float64Echo tests float64 round-trip.
func TestCOMTestServer_Float64Echo(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Float64")
	defer cleanup()

	RegisterVariantConverters()

	param, err := WrapVariant(float64(2.718281828459045))
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	defer param.Clear()

	result, err := dispatch.CallMethod("EchoFloat64", param)
	if err != nil {
		t.Fatalf("CallMethod(EchoFloat64) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[float64](result)
	if got != 2.718281828459045 {
		t.Fatalf("EchoFloat64(e) = %f, want 2.718281828459045", got)
	}
}

// TestCOMTestServer_BooleanEcho tests boolean round-trip.
func TestCOMTestServer_BooleanEcho(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Boolean")
	defer cleanup()

	RegisterVariantConverters()

	for _, val := range []bool{true, false} {
		name := "false"
		if val {
			name = "true"
		}
		t.Run(name, func(t *testing.T) {
			param, err := WrapVariant(val)
			if err != nil {
				t.Fatalf("WrapVariant failed: %v", err)
			}
			defer param.Clear()

			result, err := dispatch.CallMethod("EchoBoolean", param)
			if err != nil {
				t.Fatalf("CallMethod(EchoBoolean) failed: %v", err)
			}
			defer result.Clear()

			got := UnwrapVariant[bool](result)
			if got != val {
				t.Fatalf("EchoBoolean(%v) = %v", val, got)
			}
		})
	}
}

// TestCOMTestServer_Empty tests VT_EMPTY and VT_NULL handling.
func TestCOMTestServer_Empty(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Empty")
	defer cleanup()

	RegisterVariantConverters()

	t.Run("GetEmpty", func(t *testing.T) {
		result, err := dispatch.CallMethod("GetEmpty")
		if err != nil {
			t.Fatalf("CallMethod(GetEmpty) failed: %v", err)
		}
		defer result.Clear()

		if result.VT != VT_EMPTY {
			t.Fatalf("GetEmpty() VT = %d, want VT_EMPTY (%d)", result.VT, VT_EMPTY)
		}
	})

	t.Run("GetNull", func(t *testing.T) {
		result, err := dispatch.CallMethod("GetNull")
		if err != nil {
			t.Fatalf("CallMethod(GetNull) failed: %v", err)
		}
		defer result.Clear()

		if result.VT != VT_NULL {
			t.Fatalf("GetNull() VT = %d, want VT_NULL (%d)", result.VT, VT_NULL)
		}
	})

	t.Run("IsEmpty", func(t *testing.T) {
		param := MakeEmptyVariant()

		result, err := dispatch.CallMethod("IsEmpty", param)
		if err != nil {
			t.Fatalf("CallMethod(IsEmpty) failed: %v", err)
		}
		defer result.Clear()

		got := UnwrapVariant[bool](result)
		if !got {
			t.Fatal("IsEmpty(VT_EMPTY) = false, want true")
		}
	})

	t.Run("IsNull", func(t *testing.T) {
		param := MakeNullVariant()

		result, err := dispatch.CallMethod("IsNull", param)
		if err != nil {
			t.Fatalf("CallMethod(IsNull) failed: %v", err)
		}
		defer result.Clear()

		got := UnwrapVariant[bool](result)
		if !got {
			t.Fatal("IsNull(VT_NULL) = false, want true")
		}
	})
}

// TestCOMTestServer_DualInterface tests that dual interface classes support
// both IUnknown and IDispatch, and methods can be called.
func TestCOMTestServer_DualInterface(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Interface.Dual")
	defer cleanup()

	RegisterVariantConverters()

	result, err := dispatch.CallMethod("GetValue")
	if err != nil {
		t.Fatalf("CallMethod(GetValue) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[int32](result)
	if got != 42 {
		t.Fatalf("GetValue() = %d, want 42", got)
	}
}

// TestCOMTestServer_DispatchOnly tests that dispatch-only interface classes
// support IDispatch and methods can be called.
func TestCOMTestServer_DispatchOnly(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Interface.DispatchOnly")
	defer cleanup()

	RegisterVariantConverters()

	result, err := dispatch.CallMethod("DispatchGetValue")
	if err != nil {
		t.Fatalf("CallMethod(DispatchGetValue) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[int32](result)
	if got != 126 {
		t.Fatalf("DispatchGetValue() = %d, want 126", got)
	}
}

// TestCOMTestServer_CreateInstanceFromCLSID creates an instance using the
// CLSID directly (not ProgID) and verifies the object works.
func TestCOMTestServer_CreateInstanceFromCLSID(t *testing.T) {
	tests := []struct {
		name  string
		clsid windows.GUID
	}{
		{"COMTestString", CLSID_COMTestString},
		{"COMTestInt32", CLSID_COMTestInt32},
		{"COMTestBoolean", CLSID_COMTestBoolean},
		{"COMTestFloat64", CLSID_COMTestFloat64},
		{"COMTestDualInterface", CLSID_COMTestDualInterface},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			unknown, cleanup := createTestUnknown(t, tc.clsid)
			defer cleanup()

			dispatch, err := QueryIDispatchFromIUnknown(unknown)
			if err != nil {
				t.Fatalf("QueryIDispatchFromIUnknown failed: %v", err)
			}
			defer dispatch.Release()

			if !dispatch.HasTypeInfo() {
				t.Error("HasTypeInfo() = false, want true")
			}
		})
	}
}

// TestCOMTestServer_CreateInstanceFromCLSID_WithCall creates an instance
// using a CLSID and makes a method call to verify it works end-to-end.
func TestCOMTestServer_CreateInstanceFromCLSID_WithCall(t *testing.T) {
	unknown, cleanup := createTestUnknown(t, CLSID_COMTestString)
	defer cleanup()

	dispatch, err := QueryIDispatchFromIUnknown(unknown)
	if err != nil {
		t.Fatalf("QueryIDispatchFromIUnknown failed: %v", err)
	}
	defer dispatch.Release()

	RegisterVariantConverters()

	param, err := WrapVariant("clsid test")
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	defer param.Clear()

	result, err := dispatch.CallMethod("EchoString", param)
	if err != nil {
		t.Fatalf("CallMethod(EchoString) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[string](result)
	if got != "clsid test" {
		t.Fatalf("EchoString() = %q, want %q", got, "clsid test")
	}
}

// TestCOMTestServer_HResult tests HResult round-trip via property accessors.
func TestCOMTestServer_HResult(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.HResult")
	defer cleanup()

	RegisterVariantConverters()

	param, err := WrapVariant(int32(0))
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	defer param.Clear()

	_, err = dispatch.PutProperty("HResultField", param)
	if err != nil {
		t.Fatalf("PutProperty(HResultField) failed: %v", err)
	}

	result, err := dispatch.GetProperty("HResultField")
	if err != nil {
		t.Fatalf("GetProperty(HResultField) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[int32](result)
	if got != 0 {
		t.Fatalf("HResultField = %d, want 0 (S_OK)", got)
	}
}

// TestCOMTestServer_MultipleInterfaces tests that the primary dispatch
// interface exposes GetValueA from the first interface.
func TestCOMTestServer_MultipleInterfaces(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Interface.Multiple")
	defer cleanup()

	RegisterVariantConverters()

	result, err := dispatch.CallMethod("GetValueA")
	if err != nil {
		t.Fatalf("CallMethod(GetValueA) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[int32](result)
	if got != 10 {
		t.Fatalf("GetValueA() = %d, want 10", got)
	}
}

// TestCOMTestServer_InheritedInterface tests that both base and derived
// interface methods are accessible through IDispatch.
func TestCOMTestServer_InheritedInterface(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Interface.Inherited")
	defer cleanup()

	RegisterVariantConverters()

	t.Run("BaseMethod", func(t *testing.T) {
		result, err := dispatch.CallMethod("GetBaseValue")
		if err != nil {
			t.Fatalf("CallMethod(GetBaseValue) failed: %v", err)
		}
		defer result.Clear()

		got := UnwrapVariant[int32](result)
		if got != 100 {
			t.Fatalf("GetBaseValue() = %d, want 100", got)
		}
	})

	t.Run("DerivedMethod", func(t *testing.T) {
		result, err := dispatch.CallMethod("GetDerivedValue")
		if err != nil {
			t.Fatalf("CallMethod(GetDerivedValue) failed: %v", err)
		}
		defer result.Clear()

		got := UnwrapVariant[int32](result)
		if got != 200 {
			t.Fatalf("GetDerivedValue() = %d, want 200", got)
		}
	})
}

// TestCOMTestServer_ConnectionPoint tests IConnectionPointContainer and
// IConnectionPoint on the test COM server's connection point class.
func TestCOMTestServer_ConnectionPoint(t *testing.T) {
	_, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	classID, err := ClassIdFromProgramId("TestCOM.Interface.ConnectionPoint")
	if err != nil {
		t.Skipf("TestCOM.Interface.ConnectionPoint not registered: %v", err)
	}

	unknown, err := CreateInstance[IUnknown](classID, IID_IUnknown)
	if err != nil {
		t.Fatalf("CreateInstance failed: %v", err)
	}
	defer unknown.Release()

	container, err := QueryIConnectionPointContainerFromIUnknown(unknown)
	if err != nil {
		t.Fatalf("QueryIConnectionPointContainerFromIUnknown failed: %v", err)
	}
	defer container.Release()

	point, err := container.FindConnectionPoint(IID_ICOMTestConnectionPointEvents)
	if err != nil {
		t.Fatalf("FindConnectionPoint(ICOMTestConnectionPointEvents) failed: %v", err)
	}
	defer point.Release()

	interfaceID, err := point.GetConnectionInterface()
	if err != nil {
		t.Fatalf("GetConnectionInterface failed: %v", err)
	}

	if interfaceID != IID_ICOMTestConnectionPointEvents {
		t.Fatalf("GetConnectionInterface = %v, want %v", interfaceID, IID_ICOMTestConnectionPointEvents)
	}
}

// TestCOMTestServer_GetIDsOfNames tests name resolution through IDispatch.
func TestCOMTestServer_GetIDsOfNames(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.String")
	defer cleanup()

	names := []string{"EchoString", "PutString", "GetString", "StringField"}
	for _, name := range names {
		t.Run(name, func(t *testing.T) {
			id, err := dispatch.GetSingleIDOfName(name)
			if err != nil {
				t.Fatalf("GetSingleIDOfName(%q) failed: %v", name, err)
			}
			if id == DISPID_UNKNOWN {
				t.Fatalf("GetSingleIDOfName(%q) = DISPID_UNKNOWN", name)
			}
		})
	}
}

// TestCOMTestServer_TypeInfo tests ITypeInfo retrieval and validates the
// interface GUID matches ICOMTestString.
func TestCOMTestServer_TypeInfo(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.String")
	defer cleanup()

	if !dispatch.HasTypeInfo() {
		t.Fatal("HasTypeInfo() = false, want true")
	}

	typeInfo := dispatch.GetTypeInfo()
	if typeInfo == nil {
		t.Fatal("GetTypeInfo() = nil")
	}
	defer typeInfo.Release()

	typeAttr, err := typeInfo.GetTypeAttr()
	if err != nil {
		t.Fatalf("GetTypeAttr failed: %v", err)
	}
	if typeAttr == nil {
		t.Fatal("GetTypeAttr() = nil")
	}
	defer typeInfo.ReleaseTypeAttr(typeAttr)

	if typeAttr.CFuncs == 0 {
		t.Fatal("TypeAttr.CFuncs = 0, want > 0")
	}

	if typeAttr.Guid != IID_ICOMTestString {
		t.Fatalf("TypeAttr.Guid = %v, want %v", typeAttr.Guid, IID_ICOMTestString)
	}
}

// TestCOMTestServer_CreateInstanceFromString tests CreateInstanceFromString
// using a ProgID to create and use a COM object.
func TestCOMTestServer_CreateInstanceFromString(t *testing.T) {
	_, err := Initialize(Multithreaded)
	if err != nil {
		t.Fatalf("Initialize failed: %v", err)
	}
	defer Uninitialize()

	unknown, err := CreateInstanceFromString[IUnknown]("TestCOM.String", IID_IUnknown)
	if err != nil {
		t.Skipf("TestCOM.String not registered: %v", err)
	}
	defer unknown.Release()

	dispatch, err := QueryIDispatchFromIUnknown(unknown)
	if err != nil {
		t.Fatalf("QueryIDispatchFromIUnknown failed: %v", err)
	}
	defer dispatch.Release()

	RegisterVariantConverters()

	param, err := WrapVariant("from string")
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	defer param.Clear()

	result, err := dispatch.CallMethod("EchoString", param)
	if err != nil {
		t.Fatalf("CallMethod(EchoString) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[string](result)
	if got != "from string" {
		t.Fatalf("EchoString() = %q, want %q", got, "from string")
	}
}

// TestCOMTestServer_FileTimeEcho tests int64 round-trip via FileTime methods.
func TestCOMTestServer_FileTimeEcho(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.FileTime")
	defer cleanup()

	RegisterVariantConverters()

	// FILETIME is stored as int64 (100-nanosecond intervals since Jan 1, 1601)
	fileTime := int64(132500000000000000) // some arbitrary FILETIME value

	param, err := WrapVariant(fileTime)
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	defer param.Clear()

	result, err := dispatch.CallMethod("EchoFileTime", param)
	if err != nil {
		t.Fatalf("CallMethod(EchoFileTime) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[int64](result)
	if got != fileTime {
		t.Fatalf("EchoFileTime() = %d, want %d", got, fileTime)
	}
}

// TestCOMTestServer_CurrencyEcho tests Currency (VT_CY) round-trip.
func TestCOMTestServer_CurrencyEcho(t *testing.T) {
	_, dispatch, cleanup := createTestDispatch(t, "TestCOM.Currency")
	defer cleanup()

	RegisterVariantConverters()

	// Currency is int64 scaled by 10,000. 12345 = $1.2345
	param, err := WrapVariant(int64(123450000))
	if err != nil {
		t.Fatalf("WrapVariant failed: %v", err)
	}
	defer param.Clear()

	result, err := dispatch.CallMethod("EchoCurrency", param)
	if err != nil {
		t.Fatalf("CallMethod(EchoCurrency) failed: %v", err)
	}
	defer result.Clear()

	got := UnwrapVariant[int64](result)
	if got != 123450000 {
		t.Fatalf("EchoCurrency() = %d, want 123450000", got)
	}
}
