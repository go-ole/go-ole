//go:build windows

package ole

// CreateObject creates object from programID based on interface type.
//
// Only supports IUnknown.
//
// Program ID can be either program ID or application string.
func CreateObject(programID string) (unknown *IUnknown, err error) {
	classID, err := LookupClassId(programID)
	if err != nil {
		return
	}

	unknown, err = CreateInstance(classID, IID_IUnknown)
	if err != nil {
		return
	}

	return
}

// GetObject retrieves active object for program ID and interface ID based on interface type.
//
// Only supports IUnknown.
//
// Program ID can be either program ID or application string.
func GetObject(programID string) (unknown *IUnknown, err error) {
	classID, err := LookupClassId(programID)
	if err != nil {
		return
	}

	unknown, err = GetActiveObject(classID, IID_IUnknown)
	if err != nil {
		return
	}

	return
}

// CallMethod calls method on IDispatch with parameters.
func CallMethod(disp *IDispatch, name string, params ...interface{}) (result *VARIANT, err error) {
	return disp.InvokeWithOptionalArgs(name, DISPATCH_METHOD, params)
}

// MustCallMethod calls method on IDispatch with parameters or panics.
func MustCallMethod(disp *IDispatch, name string, params ...interface{}) (result *VARIANT) {
	r, err := CallMethod(disp, name, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// GetProperty retrieves property from IDispatch.
func GetProperty(disp *IDispatch, name string, params ...interface{}) (result *VARIANT, err error) {
	return disp.InvokeWithOptionalArgs(name, DISPATCH_PROPERTYGET, params)
}

// MustGetProperty retrieves property from IDispatch or panics.
func MustGetProperty(disp *IDispatch, name string, params ...interface{}) (result *VARIANT) {
	r, err := GetProperty(disp, name, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// PutProperty mutates property.
func PutProperty(disp *IDispatch, name string, params ...interface{}) (result *VARIANT, err error) {
	return disp.InvokeWithOptionalArgs(name, DISPATCH_PROPERTYPUT, params)
}

// MustPutProperty mutates property or panics.
func MustPutProperty(disp *IDispatch, name string, params ...interface{}) (result *VARIANT) {
	r, err := PutProperty(disp, name, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// PutPropertyRef mutates property reference.
func PutPropertyRef(disp *IDispatch, name string, params ...interface{}) (result *VARIANT, err error) {
	return disp.InvokeWithOptionalArgs(name, DISPATCH_PROPERTYPUTREF, params)
}

// MustPutPropertyRef mutates property reference or panics.
func MustPutPropertyRef(disp *IDispatch, name string, params ...interface{}) (result *VARIANT) {
	r, err := PutPropertyRef(disp, name, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}
