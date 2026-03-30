//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
)

var (
	// IID_NULL is null Interface ID, used when no other Interface ID is known.
	IID_NULL, _ = windows.GUIDFromString("{00000000-0000-0000-0000-000000000000}")

	// IID_IUnknown is for IUnknown interfaces.
	IID_IUnknown, _ = windows.GUIDFromString("{00000000-0000-0000-C000-000000000046}")

	// IID_IDispatch is for IDispatch interfaces.
	IID_IDispatch, _ = windows.GUIDFromString("{00020400-0000-0000-C000-000000000046}")

	// IID_IEnumVariant is for IEnumVariant interfaces
	IID_IEnumVariant, _ = windows.GUIDFromString("{00020404-0000-0000-C000-000000000046}")

	// IID_IConnectionPointContainer is for IConnectionPointContainer interfaces.
	IID_IConnectionPointContainer, _ = windows.GUIDFromString("{B196B284-BAB4-101A-B69C-00AA00341D07}")

	// IID_IConnectionPoint is for IConnectionPoint interfaces.
	IID_IConnectionPoint, _ = windows.GUIDFromString("{B196B286-BAB4-101A-B69C-00AA00341D07}")

	// IID_IInspectable is for IInspectable interfaces.
	IID_IInspectable, _ = windows.GUIDFromString("{AF86E2E0-B12D-4C6A-9C5A-D7AA65101E90}")

	// IID_IProvideClassInfo is for IProvideClassInfo interfaces.
	IID_IProvideClassInfo, _ = windows.GUIDFromString("{B196B283-BAB4-101A-B69C-00AA00341D07}")

	// IID_IActivationFactory is for IActivationFactory interfaces.
	IID_IActivationFactory, _ = windows.GUIDFromString("{00000035-0000-0000-C000-000000000046}")
)

// Test COM Server v2 - Types
// These are for testing and not part of any library.
// Source: https://github.com/go-ole/test-com-server v2.0.0
var (
	// IID_ICOMTestInt8 is the interface ID for ICOMTestInt8.
	IID_ICOMTestInt8, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000001}")
	// CLSID_COMTestInt8 is the class ID for COMTestInt8.
	CLSID_COMTestInt8, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000002}")

	// IID_ICOMTestInt16 is the interface ID for ICOMTestInt16.
	IID_ICOMTestInt16, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000003}")
	// CLSID_COMTestInt16 is the class ID for COMTestInt16.
	CLSID_COMTestInt16, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000004}")

	// IID_ICOMTestInt32 is the interface ID for ICOMTestInt32.
	IID_ICOMTestInt32, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000005}")
	// CLSID_COMTestInt32 is the class ID for COMTestInt32.
	CLSID_COMTestInt32, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000006}")

	// IID_ICOMTestInt64 is the interface ID for ICOMTestInt64.
	IID_ICOMTestInt64, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000007}")
	// CLSID_COMTestInt64 is the class ID for COMTestInt64.
	CLSID_COMTestInt64, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000008}")

	// IID_ICOMTestFloat32 is the interface ID for ICOMTestFloat32.
	IID_ICOMTestFloat32, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000009}")
	// CLSID_COMTestFloat32 is the class ID for COMTestFloat32.
	CLSID_COMTestFloat32, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000000A}")

	// IID_ICOMTestFloat64 is the interface ID for ICOMTestFloat64.
	IID_ICOMTestFloat64, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000000B}")
	// CLSID_COMTestFloat64 is the class ID for COMTestFloat64.
	CLSID_COMTestFloat64, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000000C}")

	// IID_ICOMTestString is the interface ID for ICOMTestString.
	IID_ICOMTestString, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000000D}")
	// CLSID_COMTestString is the class ID for COMTestString.
	CLSID_COMTestString, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000000E}")

	// IID_ICOMTestBoolean is the interface ID for ICOMTestBoolean.
	IID_ICOMTestBoolean, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000000F}")
	// CLSID_COMTestBoolean is the class ID for COMTestBoolean.
	CLSID_COMTestBoolean, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000010}")

	// IID_ICOMTestCurrency is the interface ID for ICOMTestCurrency.
	IID_ICOMTestCurrency, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000011}")
	// CLSID_COMTestCurrency is the class ID for COMTestCurrency.
	CLSID_COMTestCurrency, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000012}")

	// IID_ICOMTestDate is the interface ID for ICOMTestDate.
	IID_ICOMTestDate, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000013}")
	// CLSID_COMTestDate is the class ID for COMTestDate.
	CLSID_COMTestDate, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000014}")

	// IID_ICOMTestDecimal is the interface ID for ICOMTestDecimal.
	IID_ICOMTestDecimal, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000015}")
	// CLSID_COMTestDecimal is the class ID for COMTestDecimal.
	CLSID_COMTestDecimal, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000016}")

	// IID_ICOMTestError is the interface ID for ICOMTestError.
	IID_ICOMTestError, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000017}")
	// CLSID_COMTestError is the class ID for COMTestError.
	CLSID_COMTestError, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000018}")

	// IID_ICOMTestVariant is the interface ID for ICOMTestVariant.
	IID_ICOMTestVariant, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000019}")
	// CLSID_COMTestVariant is the class ID for COMTestVariant.
	CLSID_COMTestVariant, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000001A}")

	// IID_ICOMTestUnknown is the interface ID for ICOMTestUnknown.
	IID_ICOMTestUnknown, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000001B}")
	// CLSID_COMTestUnknown is the class ID for COMTestUnknown.
	CLSID_COMTestUnknown, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000001C}")

	// IID_ICOMTestDispatch is the interface ID for ICOMTestDispatch.
	IID_ICOMTestDispatch, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000001D}")
	// CLSID_COMTestDispatch is the class ID for COMTestDispatch.
	CLSID_COMTestDispatch, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000001E}")

	// IID_ICOMTestEmpty is the interface ID for ICOMTestEmpty.
	IID_ICOMTestEmpty, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000001F}")
	// CLSID_COMTestEmpty is the class ID for COMTestEmpty.
	CLSID_COMTestEmpty, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000020}")

	// IID_ICOMTestClsid is the interface ID for ICOMTestClsid.
	IID_ICOMTestClsid, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000021}")
	// CLSID_COMTestClsid is the class ID for COMTestClsid.
	CLSID_COMTestClsid, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000022}")

	// IID_ICOMTestHResult is the interface ID for ICOMTestHResult.
	IID_ICOMTestHResult, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000023}")
	// CLSID_COMTestHResult is the class ID for COMTestHResult.
	CLSID_COMTestHResult, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000024}")

	// IID_ICOMTestFileTime is the interface ID for ICOMTestFileTime.
	IID_ICOMTestFileTime, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000025}")
	// CLSID_COMTestFileTime is the class ID for COMTestFileTime.
	CLSID_COMTestFileTime, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000026}")

	// IID_ICOMTestStream is the interface ID for ICOMTestStream.
	IID_ICOMTestStream, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000027}")
	// CLSID_COMTestStream is the class ID for COMTestStream.
	CLSID_COMTestStream, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-000000000028}")

	// IID_ICOMTestBlob is the interface ID for ICOMTestBlob.
	IID_ICOMTestBlob, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000002A}")
	// CLSID_COMTestBlob is the class ID for COMTestBlob.
	CLSID_COMTestBlob, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000002B}")

	// IID_ICOMTestPtr is the interface ID for ICOMTestPtr.
	IID_ICOMTestPtr, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000002C}")
	// CLSID_COMTestPtr is the class ID for COMTestPtr.
	CLSID_COMTestPtr, _ = windows.GUIDFromString("{A1B2C3D4-1111-1111-1111-00000000002D}")
)

// Test COM Server v2 - SafeArrays
var (
	// IID_ICOMSafeArrayInt8 is the interface ID for ICOMSafeArrayInt8.
	IID_ICOMSafeArrayInt8, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000001}")
	// CLSID_COMSafeArrayInt8 is the class ID for COMSafeArrayInt8.
	CLSID_COMSafeArrayInt8, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000002}")

	// IID_ICOMSafeArrayInt16 is the interface ID for ICOMSafeArrayInt16.
	IID_ICOMSafeArrayInt16, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000003}")
	// CLSID_COMSafeArrayInt16 is the class ID for COMSafeArrayInt16.
	CLSID_COMSafeArrayInt16, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000004}")

	// IID_ICOMSafeArrayInt32 is the interface ID for ICOMSafeArrayInt32.
	IID_ICOMSafeArrayInt32, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000005}")
	// CLSID_COMSafeArrayInt32 is the class ID for COMSafeArrayInt32.
	CLSID_COMSafeArrayInt32, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000006}")

	// IID_ICOMSafeArrayInt64 is the interface ID for ICOMSafeArrayInt64.
	IID_ICOMSafeArrayInt64, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000007}")
	// CLSID_COMSafeArrayInt64 is the class ID for COMSafeArrayInt64.
	CLSID_COMSafeArrayInt64, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000008}")

	// IID_ICOMSafeArrayFloat32 is the interface ID for ICOMSafeArrayFloat32.
	IID_ICOMSafeArrayFloat32, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000009}")
	// CLSID_COMSafeArrayFloat32 is the class ID for COMSafeArrayFloat32.
	CLSID_COMSafeArrayFloat32, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-00000000000A}")

	// IID_ICOMSafeArrayFloat64 is the interface ID for ICOMSafeArrayFloat64.
	IID_ICOMSafeArrayFloat64, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-00000000000B}")
	// CLSID_COMSafeArrayFloat64 is the class ID for COMSafeArrayFloat64.
	CLSID_COMSafeArrayFloat64, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-00000000000C}")

	// IID_ICOMSafeArrayString is the interface ID for ICOMSafeArrayString.
	IID_ICOMSafeArrayString, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-00000000000D}")
	// CLSID_COMSafeArrayString is the class ID for COMSafeArrayString.
	CLSID_COMSafeArrayString, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-00000000000E}")

	// IID_ICOMSafeArrayBoolean is the interface ID for ICOMSafeArrayBoolean.
	IID_ICOMSafeArrayBoolean, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-00000000000F}")
	// CLSID_COMSafeArrayBoolean is the class ID for COMSafeArrayBoolean.
	CLSID_COMSafeArrayBoolean, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000010}")

	// IID_ICOMSafeArrayCurrency is the interface ID for ICOMSafeArrayCurrency.
	IID_ICOMSafeArrayCurrency, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000011}")
	// CLSID_COMSafeArrayCurrency is the class ID for COMSafeArrayCurrency.
	CLSID_COMSafeArrayCurrency, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000012}")

	// IID_ICOMSafeArrayDate is the interface ID for ICOMSafeArrayDate.
	IID_ICOMSafeArrayDate, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000013}")
	// CLSID_COMSafeArrayDate is the class ID for COMSafeArrayDate.
	CLSID_COMSafeArrayDate, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000014}")

	// IID_ICOMSafeArrayDecimal is the interface ID for ICOMSafeArrayDecimal.
	IID_ICOMSafeArrayDecimal, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000015}")
	// CLSID_COMSafeArrayDecimal is the class ID for COMSafeArrayDecimal.
	CLSID_COMSafeArrayDecimal, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000016}")

	// IID_ICOMSafeArrayVariant is the interface ID for ICOMSafeArrayVariant.
	IID_ICOMSafeArrayVariant, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000017}")
	// CLSID_COMSafeArrayVariant is the class ID for COMSafeArrayVariant.
	CLSID_COMSafeArrayVariant, _ = windows.GUIDFromString("{A1B2C3D4-2222-2222-2222-000000000018}")
)

// Test COM Server v2 - Interfaces
var (
	// IID_ICOMTestDualInterface is the interface ID for ICOMTestDualInterface.
	IID_ICOMTestDualInterface, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-000000000001}")
	// CLSID_COMTestDualInterface is the class ID for COMTestDualInterface.
	CLSID_COMTestDualInterface, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-000000000002}")

	// IID_ICOMTestUnknownOnly is the interface ID for ICOMTestUnknownOnly.
	IID_ICOMTestUnknownOnly, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-000000000003}")
	// CLSID_COMTestUnknownOnly is the class ID for COMTestUnknownOnly.
	CLSID_COMTestUnknownOnly, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-000000000004}")

	// IID_ICOMTestDispatchOnly is the interface ID for ICOMTestDispatchOnly.
	IID_ICOMTestDispatchOnly, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-000000000005}")
	// CLSID_COMTestDispatchOnly is the class ID for COMTestDispatchOnly.
	CLSID_COMTestDispatchOnly, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-000000000006}")

	// IID_ICOMTestMultiA is the interface ID for ICOMTestMultiA.
	IID_ICOMTestMultiA, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-000000000007}")
	// IID_ICOMTestMultiB is the interface ID for ICOMTestMultiB.
	IID_ICOMTestMultiB, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-000000000008}")
	// IID_ICOMTestMultiC is the interface ID for ICOMTestMultiC.
	IID_ICOMTestMultiC, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-000000000009}")
	// CLSID_COMTestMultipleInterfaces is the class ID for COMTestMultipleInterfaces.
	CLSID_COMTestMultipleInterfaces, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-00000000000A}")

	// IID_ICOMTestBaseInterface is the interface ID for ICOMTestBaseInterface.
	IID_ICOMTestBaseInterface, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-00000000000B}")
	// IID_ICOMTestDerivedInterface is the interface ID for ICOMTestDerivedInterface.
	IID_ICOMTestDerivedInterface, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-00000000000C}")
	// CLSID_COMTestInheritedInterface is the class ID for COMTestInheritedInterface.
	CLSID_COMTestInheritedInterface, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-00000000000D}")

	// IID_ICOMTestConnectionPoint is the interface ID for ICOMTestConnectionPoint.
	IID_ICOMTestConnectionPoint, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-00000000000E}")
	// IID_ICOMTestConnectionPointEvents is the interface ID for ICOMTestConnectionPointEvents.
	IID_ICOMTestConnectionPointEvents, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-00000000000F}")
	// CLSID_COMTestConnectionPoint is the class ID for COMTestConnectionPoint.
	CLSID_COMTestConnectionPoint, _ = windows.GUIDFromString("{A1B2C3D4-3333-3333-3333-000000000010}")
)
