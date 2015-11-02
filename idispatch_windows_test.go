// +build windows

package ole

import "testing"

func wrapCOMExecute(t *testing.T, callback func(*testing.T)) {
	defer func() {
		if r := recover(); r != nil {
			t.Error(r)
		}
	}()

	err := CoInitialize(0)
	if err != nil {
		t.Fatal(err)
	}
	defer CoUninitialize()

	callback(t)
}

func wrapDispatch(t *testing.T, ClassID, UnknownInterfaceID, DispatchInterfaceID *GUID, callback func(*testing.T, *IUnknown, *IDispatch)) {
	var unknown *IUnknown
	var dispatch *IDispatch
	var err error

	unknown, err = CreateInstance(ClassID, UnknownInterfaceID)
	if err != nil {
		t.Error(err)
		return
	}
	defer unknown.Release()

	dispatch, err = unknown.QueryInterface(IID_IDispatch)
	if err != nil {
		t.Error(err)
		return
	}
	defer dispatch.Release()

	callback(t, unknown, dispatch)
}

func wrapGoOLETestCOMServer(t *testing.T, callback func(*testing.T, *IUnknown, *IDispatch)) {
	wrapCOMExecute(t, func(t *testing.T) {
		wrapDispatch(t, CLSID_COMEchoTestObject, IID_IUnknown, IID_ICOMEchoTestObject, callback)
	})
}

func TestIDispatch_goolecomserver_echostring(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoString"
		expected := "Test String"
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(string)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "string", variant.VT, variant.Val)
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echoint8(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoInt8"
		expected := int8(1)
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(int8)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int8", variant.VT, variant.Val)
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echouint8(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoUInt8"
		expected := uint8(1)
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(uint8)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint8", variant.VT, variant.Val)
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echoint16(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoInt16"
		expected := int16(1)
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(int16)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int16", variant.VT, variant.Val)
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echouint16(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoUInt16"
		expected := uint16(1)
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(uint16)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint16", variant.VT, variant.Val)
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echoint32(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoInt32"
		expected := int32(2)
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(int32)
		if passed {
			if actual != expected {
				t.Errorf("%s() expected %v did not match %v", method, expected, actual)
			}
		}

		actualInt, passed := variant.Value().(int)
		if passed {
			if actualInt != int(expected) {
				t.Errorf("%s() expected %v did not match %v", method, expected, actualInt)
			}
		} else {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int32", variant.VT, variant.Val)
		}
	})
}

func TestIDispatch_goolecomserver_echouint32(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoUInt32"
		expected := uint32(4)
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(uint32)
		if passed {
			if actual != expected {
				t.Errorf("%s() expected %v did not match %v", method, expected, actual)
			}
		}

		actualUInt, passed := variant.Value().(uint)
		if passed {
			if actualUInt != uint(expected) {
				t.Errorf("%s() expected %v did not match %v", method, expected, actualUInt)
			}
		} else {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint32", variant.VT, variant.Val)
		}
	})
}

func TestIDispatch_goolecomserver_echoint64(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoInt64"
		expected := int64(1)
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(int64)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int64", variant.VT, variant.Val)
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echouint64(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoUInt64"
		expected := uint64(1)
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(uint64)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint64", variant.VT, variant.Val)
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echofloat32(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoFloat32"
		expected := float32(2.2)
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(float32)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "float32", variant.VT, variant.Val)
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echofloat64(t *testing.T) {
	wrapGoOLETestCOMServer(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoFloat64"
		expected := float64(2.2)
		variant, err := idispatch.CallMethod(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant.Clear()
		actual, passed := variant.Value().(float64)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "float64", variant.VT, variant.Val)
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}
