//go:build windows

package ole

// These examples demonstrate COM object creation and interaction patterns.
// They require COM servers registered on the system and cannot produce
// deterministic output, so they serve as documentation only.

// CreateInstance creates a COM object given a class ID and interface ID.
// The result is a pointer to the interface pointer and must be dereferenced.
func ExampleCreateInstance() {
	InitializeMultithreaded()
	defer Uninitialize()

	clsid, err := ClassIdFromString("Excel.Application")
	if err != nil {
		return
	}
	unknownPtr, err := CreateInstance[*IUnknown](clsid, IID_IUnknown)
	if err != nil {
		return
	}
	unknown := *unknownPtr
	defer unknown.Release()
}

// CreateInstance can target IDispatch directly when the COM class supports it.
func ExampleCreateInstance_dispatch() {
	InitializeMultithreaded()
	defer Uninitialize()

	clsid, err := ClassIdFromString("Scripting.Dictionary")
	if err != nil {
		return
	}
	dispatchPtr, err := CreateInstance[*IDispatch](clsid, IID_IDispatch)
	if err != nil {
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()
}

// CreateInstance is typically followed by QueryInterfaceOnIUnknown to
// obtain the IDispatch interface for method calls and property access.
func ExampleCreateInstance_withQueryInterface() {
	InitializeMultithreaded()
	defer Uninitialize()
	RegisterVariantConverters()

	clsid, err := ClassIdFromString("Excel.Application")
	if err != nil {
		return
	}
	unknownPtr, err := CreateInstance[*IUnknown](clsid, IID_IUnknown)
	if err != nil {
		return
	}
	unknown := *unknownPtr
	defer unknown.Release()

	dispatch, err := QueryInterfaceOnIUnknown[IDispatch](unknown, IID_IDispatch)
	if err != nil {
		return
	}
	defer dispatch.Release()

	dispatch.PutProperty("Visible", BoolToVariant(true))
	dispatch.CallMethod("Quit")
}

// CreateInstanceFromString combines ClassIdFromString and CreateInstance
// into a single call using the program ID or GUID string.
func ExampleCreateInstanceFromString() {
	InitializeMultithreaded()
	defer Uninitialize()

	unknownPtr, err := CreateInstanceFromString[*IUnknown]("Excel.Application", IID_IUnknown)
	if err != nil {
		return
	}
	unknown := *unknownPtr
	defer unknown.Release()
}

// CreateInstanceFromString works with any registered COM program ID.
func ExampleCreateInstanceFromString_dictionary() {
	InitializeMultithreaded()
	defer Uninitialize()

	dispatchPtr, err := CreateInstanceFromString[*IDispatch]("Scripting.Dictionary", IID_IDispatch)
	if err != nil {
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()
}

// CreateInstanceFromString can also accept a GUID string.
func ExampleCreateInstanceFromString_guid() {
	InitializeMultithreaded()
	defer Uninitialize()

	unknownPtr, err := CreateInstanceFromString[*IUnknown]("{00024500-0000-0000-C000-000000000046}", IID_IUnknown)
	if err != nil {
		return
	}
	unknown := *unknownPtr
	defer unknown.Release()
}

// GetActiveObject connects to an already-running COM object.
func ExampleGetActiveObject() {
	InitializeMultithreaded()
	defer Uninitialize()

	clsid, err := ClassIdFromString("Excel.Application")
	if err != nil {
		return
	}
	dispatchPtr, err := GetActiveObject[*IDispatch](clsid, IID_IDispatch)
	if err != nil {
		// No running Excel instance found.
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()
}

// GetActiveObjectFromString connects to a running COM object by program ID.
func ExampleGetActiveObjectFromString() {
	InitializeMultithreaded()
	defer Uninitialize()

	dispatchPtr, err := GetActiveObjectFromString[*IDispatch]("Excel.Application", IID_IDispatch)
	if err != nil {
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()
}

// QueryInterfaceOnIUnknown casts an IUnknown to a specific COM interface.
func ExampleQueryInterfaceOnIUnknown() {
	InitializeMultithreaded()
	defer Uninitialize()

	unknownPtr, err := CreateInstanceFromString[*IUnknown]("Excel.Application", IID_IUnknown)
	if err != nil {
		return
	}
	unknown := *unknownPtr
	defer unknown.Release()

	dispatch, err := QueryInterfaceOnIUnknown[IDispatch](unknown, IID_IDispatch)
	if err != nil {
		return
	}
	defer dispatch.Release()
}

// MustQueryInterfaceOnIUnknown is the panicking variant of QueryInterfaceOnIUnknown.
func ExampleMustQueryInterfaceOnIUnknown() {
	InitializeMultithreaded()
	defer Uninitialize()

	unknownPtr, err := CreateInstanceFromString[*IUnknown]("Scripting.Dictionary", IID_IUnknown)
	if err != nil {
		return
	}
	unknown := *unknownPtr
	defer unknown.Release()

	dispatch := MustQueryInterfaceOnIUnknown[IDispatch](unknown, IID_IDispatch)
	defer dispatch.Release()
}

// QueryIDispatchFromIUnknown is a convenience for obtaining IDispatch from IUnknown.
func ExampleQueryIDispatchFromIUnknown() {
	InitializeMultithreaded()
	defer Uninitialize()

	unknownPtr, err := CreateInstanceFromString[*IUnknown]("Scripting.Dictionary", IID_IUnknown)
	if err != nil {
		return
	}
	unknown := *unknownPtr
	defer unknown.Release()

	dispatch, err := QueryIDispatchFromIUnknown(unknown)
	if err != nil {
		return
	}
	defer dispatch.Release()
}

// CallMethod invokes a method on an IDispatch object.
// Parameters must be *VARIANT values.
func ExampleIDispatch_CallMethod() {
	InitializeMultithreaded()
	defer Uninitialize()
	RegisterVariantConverters()

	dispatchPtr, err := CreateInstanceFromString[*IDispatch]("Scripting.Dictionary", IID_IDispatch)
	if err != nil {
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()

	dispatch.CallMethod("Add", StringToBStrVariant("key"), StringToBStrVariant("value"))
}

// MustCallMethod is the panicking variant that simplifies chaining.
func ExampleIDispatch_MustCallMethod() {
	InitializeMultithreaded()
	defer Uninitialize()
	RegisterVariantConverters()

	dispatchPtr, err := CreateInstanceFromString[*IDispatch]("Scripting.Dictionary", IID_IDispatch)
	if err != nil {
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()

	dispatch.MustCallMethod("Add", StringToBStrVariant("key"), Int32ToVariant(int32(42)))
}

// GetProperty reads a property value from a COM object.
func ExampleIDispatch_GetProperty() {
	InitializeMultithreaded()
	defer Uninitialize()
	RegisterVariantConverters()

	dispatchPtr, err := CreateInstanceFromString[*IDispatch]("Scripting.Dictionary", IID_IDispatch)
	if err != nil {
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()

	count, err := dispatch.GetProperty("Count")
	if err != nil {
		return
	}
	_ = UnwrapVariant[int32](count)
}

// PutProperty sets a property value on a COM object.
func ExampleIDispatch_PutProperty() {
	InitializeMultithreaded()
	defer Uninitialize()
	RegisterVariantConverters()

	unknownPtr, err := CreateInstanceFromString[*IUnknown]("Excel.Application", IID_IUnknown)
	if err != nil {
		return
	}
	unknown := *unknownPtr
	defer unknown.Release()

	dispatch, err := QueryInterfaceOnIUnknown[IDispatch](unknown, IID_IDispatch)
	if err != nil {
		return
	}
	defer dispatch.Release()

	dispatch.PutProperty("Visible", BoolToVariant(true))
}

// Invoke allows specifying the dispatch type (method, property get/put).
func ExampleIDispatch_Invoke() {
	InitializeMultithreaded()
	defer Uninitialize()
	RegisterVariantConverters()

	dispatchPtr, err := CreateInstanceFromString[*IDispatch]("Scripting.Dictionary", IID_IDispatch)
	if err != nil {
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()

	result, err := dispatch.Invoke("Count", DISPATCH_PROPERTYGET)
	if err != nil {
		return
	}
	_ = result
}

// HasTypeInfo checks whether the COM object exposes type information.
func ExampleIDispatch_HasTypeInfo() {
	InitializeMultithreaded()
	defer Uninitialize()

	dispatchPtr, err := CreateInstanceFromString[*IDispatch]("Scripting.Dictionary", IID_IDispatch)
	if err != nil {
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()

	if dispatch.HasTypeInfo() {
		typeInfo := dispatch.GetTypeInfo()
		_ = typeInfo
	}
}

// GetIDsOfNames resolves method/property names to dispatch IDs.
func ExampleIDispatch_GetIDsOfNames() {
	InitializeMultithreaded()
	defer Uninitialize()

	dispatchPtr, err := CreateInstanceFromString[*IDispatch]("Scripting.Dictionary", IID_IDispatch)
	if err != nil {
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()

	ids, err := dispatch.GetIDsOfNames([]string{"Count", "Add"})
	if err != nil {
		return
	}
	_ = ids
}

// GetSingleIDOfName resolves a single name to its dispatch ID.
func ExampleIDispatch_GetSingleIDOfName() {
	InitializeMultithreaded()
	defer Uninitialize()

	dispatchPtr, err := CreateInstanceFromString[*IDispatch]("Scripting.Dictionary", IID_IDispatch)
	if err != nil {
		return
	}
	dispatch := *dispatchPtr
	defer dispatch.Release()

	id, err := dispatch.GetSingleIDOfName("Count")
	if err != nil {
		return
	}
	_ = id
}

// QueryIConnectionPointContainerFromIUnknown casts to the connection point
// container interface for event subscription.
func ExampleQueryIConnectionPointContainerFromIUnknown() {
	InitializeMultithreaded()
	defer Uninitialize()

	unknownPtr, err := CreateInstanceFromString[*IUnknown]("InternetExplorer.Application", IID_IUnknown)
	if err != nil {
		return
	}
	unknown := *unknownPtr
	defer unknown.Release()

	container, err := QueryIConnectionPointContainerFromIUnknown(unknown)
	if err != nil {
		return
	}
	defer container.Release()
}

// FindConnectionPoint locates a specific event interface on a connection
// point container.
func ExampleIConnectionPointContainer_FindConnectionPoint() {
	InitializeMultithreaded()
	defer Uninitialize()

	unknownPtr, err := CreateInstanceFromString[*IUnknown]("InternetExplorer.Application", IID_IUnknown)
	if err != nil {
		return
	}
	unknown := *unknownPtr
	defer unknown.Release()

	container, err := QueryIConnectionPointContainerFromIUnknown(unknown)
	if err != nil {
		return
	}
	defer container.Release()

	eventIID, _ := ClassIdFromGuidString("{34A715A0-6587-11D0-924A-0020AFC7AC4D}")
	point, err := container.FindConnectionPoint(eventIID)
	if err != nil {
		return
	}
	defer point.Release()
}

// ClassIdFromProgramId looks up a CLSID from the Windows registry using a
// registered Program ID like "Excel.Application".
func ExampleClassIdFromProgramId() {
	_, err := ClassIdFromProgramId("Scripting.Dictionary")
	if err != nil {
		return
	}
}

// CoInitializeSecurity sets process-wide COM security. Must be called
// before any COM objects are created.
func ExampleCoInitializeSecurity() {
	InitializeMultithreaded()
	defer Uninitialize()

	err := CoInitializeSecurity(-1, 0, 3, 0)
	if err != nil {
		return
	}
}

// GetUserDefaultLCID returns the current user's default locale identifier.
func ExampleGetUserDefaultLCID() {
	lcid := GetUserDefaultLCID()
	_ = lcid
}
