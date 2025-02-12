package ole

import (
	"golang.org/x/sys/windows"
)

var (
	// IID_NULL is null Interface ID, used when no other Interface ID is known.
	IID_NULL = windows.GUIDFromString("{00000000-0000-0000-0000-000000000000}")

	// IID_IUnknown is for IUnknown interfaces.
	IID_IUnknown = windows.GUIDFromString("{00000000-0000-0000-C000-000000000046}")

	// IID_IDispatch is for IDispatch interfaces.
	IID_IDispatch = windows.GUIDFromString("{00020400-0000-0000-C000-000000000046}")

	// IID_IEnumVariant is for IEnumVariant interfaces
	IID_IEnumVariant = windows.GUIDFromString("{00020404-0000-0000-C000-000000000046}")

	// IID_IConnectionPointContainer is for IConnectionPointContainer interfaces.
	IID_IConnectionPointContainer = windows.GUIDFromString("{B196B284-BAB4-101A-B69C-00AA00341D07}")

	// IID_IConnectionPoint is for IConnectionPoint interfaces.
	IID_IConnectionPoint = windows.GUIDFromString("{B196B286-BAB4-101A-B69C-00AA00341D07}")

	// IID_IInspectable is for IInspectable interfaces.
	IID_IInspectable = windows.GUIDFromString("{AF86E2E0-B12D-4C6A-9C5A-D7AA65101E90}")

	// IID_IProvideClassInfo is for IProvideClassInfo interfaces.
	IID_IProvideClassInfo = windows.GUIDFromString("{B196B283-BAB4-101A-B69C-00AA00341D07}")
)

// These are for testing and not part of any library.
var (
	// IID_ICOMTestString is for ICOMTestString interfaces.
	//
	// {E0133EB4-C36F-469A-9D3D-C66B84BE19ED}
	IID_ICOMTestString = windows.GUIDFromString("{E0133EB4-C36F-469A-9D3D-C66B84BE19ED}")

	// IID_ICOMTestInt8 is for ICOMTestInt8 interfaces.
	//
	// {BEB06610-EB84-4155-AF58-E2BFF53680B4}
	IID_ICOMTestInt8 = windows.GUIDFromString("{BEB06610-EB84-4155-AF58-E2BFF53680B4}")

	// IID_ICOMTestInt16 is for ICOMTestInt16 interfaces.
	//
	// {DAA3F9FA-761E-4976-A860-8364CE55F6FC}
	IID_ICOMTestInt16 = windows.GUIDFromString("{DAA3F9FA-761E-4976-A860-8364CE55F6FC}")

	// IID_ICOMTestInt32 is for ICOMTestInt32 interfaces.
	//
	// {E3DEDEE7-38A2-4540-91D1-2EEF1D8891B0}
	IID_ICOMTestInt32 = windows.GUIDFromString("{E3DEDEE7-38A2-4540-91D1-2EEF1D8891B0}")

	// IID_ICOMTestInt64 is for ICOMTestInt64 interfaces.
	//
	// {8D437CBC-B3ED-485C-BC32-C336432A1623}
	IID_ICOMTestInt64 = windows.GUIDFromString("{8D437CBC-B3ED-485C-BC32-C336432A1623}")

	// IID_ICOMTestFloat is for ICOMTestFloat interfaces.
	//
	// {BF1ED004-EA02-456A-AA55-2AC8AC6B054C}
	IID_ICOMTestFloat = windows.GUIDFromString("{BF1ED004-EA02-456A-AA55-2AC8AC6B054C}")

	// IID_ICOMTestDouble is for ICOMTestDouble interfaces.
	//
	// {BF908A81-8687-4E93-999F-D86FAB284BA0}
	IID_ICOMTestDouble = windows.GUIDFromString("{BF908A81-8687-4E93-999F-D86FAB284BA0}")

	// IID_ICOMTestBoolean is for ICOMTestBoolean interfaces.
	//
	// {D530E7A6-4EE8-40D1-8931-3D63B8605010}
	IID_ICOMTestBoolean = windows.GUIDFromString("{D530E7A6-4EE8-40D1-8931-3D63B8605010}")

	// IID_ICOMEchoTestObject is for ICOMEchoTestObject interfaces.
	//
	// {6485B1EF-D780-4834-A4FE-1EBB51746CA3}
	IID_ICOMEchoTestObject = windows.GUIDFromString("{6485B1EF-D780-4834-A4FE-1EBB51746CA3}")

	// IID_ICOMTestTypes is for ICOMTestTypes interfaces.
	//
	// {CCA8D7AE-91C0-4277-A8B3-FF4EDF28D3C0}
	IID_ICOMTestTypes = windows.GUIDFromString("{CCA8D7AE-91C0-4277-A8B3-FF4EDF28D3C0}")

	// CLSID_COMEchoTestObject is for COMEchoTestObject class.
	//
	// {3C24506A-AE9E-4D50-9157-EF317281F1B0}
	CLSID_COMEchoTestObject = windows.GUIDFromString("{3C24506A-AE9E-4D50-9157-EF317281F1B0}")

	// CLSID_COMTestScalarClass is for COMTestScalarClass class.
	//
	// {865B85C5-0334-4AC6-9EF6-AACEC8FC5E86}
	CLSID_COMTestScalarClass = windows.GUIDFromString("{865B85C5-0334-4AC6-9EF6-AACEC8FC5E86}")
)
