// +build windows

package ole

import (
	"reflect"
	"testing"
)

func TestIDispatch(t *testing.T) {
	var at string
	defer func() {
		if r := recover(); r != nil {
			t.Errorf("Recovered %v at %s", r, at)
		}
	}()

	var err error

	err = CoInitialize(0)
	if err != nil {
		t.Fatal(err)
	}

	defer CoUninitialize()

	var unknown *IUnknown
	var dispatch *IDispatch

	// oleutil.CreateObject()
	unknown, err = CreateInstance(CLSID_COMEchoTestObject, IID_IUnknown)
	if err != nil {
		t.Error(err)
		return
	}
	defer unknown.Release()

	dispatch, err = unknown.QueryInterface(IID_ICOMEchoTestObject)
	if err != nil {
		t.Fatal(err)
		return
	}
	defer dispatch.Release()

	echoValue := func(method string, value interface{}) (interface{}, bool) {
		var dispid []int32
		var err error

		dispid, err = dispatch.GetIDsOfName([]string{method})
		if err != nil {
			t.Fatal(err)
			return nil, false
		}

		result, err := dispatch.Invoke(dispid[0], DISPATCH_METHOD, value)
		if err != nil {
			t.Fatal(err)
			return nil, false
		}

		return result.Value(), true
	}

	methods := map[string]interface{}{
		"EchoInt8":   int8(1),
		"EchoInt16":  int16(1),
		"EchoInt64":  int64(1),
		"EchoUInt8":  uint8(1),
		"EchoUInt16": uint16(1),
		"EchoUInt64": uint64(1)}

	for method, expected := range methods {
		at = method
		if actual, passed := echoValue(method, expected); passed {
			if !reflect.DeepEqual(expected, actual) {
				t.Errorf("%s() expected %v did not match %v", method, expected, actual)
			}
		}
	}

	at = "EchoInt32"
	valueInt32 := int32(2)
	if actual, passed := echoValue("EchoInt32", valueInt32); passed {
		value := actual.(int32)
		if value != valueInt32 {
			t.Errorf("%s() expected %v did not match %v", "EchoInt32", valueInt32, value)
		}
	}

	at = "EchoUInt32"
	valueUInt32 := uint32(4)
	if actual, passed := echoValue("EchoUInt32", valueUInt32); passed {
		value := actual.(uint32)
		if value != valueUInt32 {
			t.Errorf("%s() expected %v did not match %v", "EchoUInt32", valueUInt32, value)
		}
	}

	at = "EchoFloat32"
	valueFloat32 := float32(2.2)
	if actual, passed := echoValue("EchoFloat32", valueFloat32); passed {
		value := actual.(float32)
		if value != valueFloat32 {
			t.Errorf("%s() expected %v did not match %v", "EchoFloat32", valueFloat32, value)
		}
	}

	at = "EchoFloat64"
	valueFloat64 := float64(2.2)
	if actual, passed := echoValue("EchoFloat64", valueFloat64); passed {
		value := actual.(float64)
		if value != valueFloat64 {
			t.Errorf("%s() expected %v did not match %v", "EchoFloat64", valueFloat64, value)
		}
	}

	at = "EchoString"
	valueString := "Test String"
	if actual, passed := echoValue("EchoString", valueString); passed {
		value := actual.(string)
		if value != valueString {
			t.Errorf("%s() expected %v did not match %v", "EchoString", valueString, value)
		}
	}
}
