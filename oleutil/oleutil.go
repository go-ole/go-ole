package oleutil

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/legacy"
)

// ClassIDFrom retrieves class ID whether given is program ID or application string.
func ClassIDFrom(programID string) (classID *legacy.GUID, err error) {
	return legacy.ClassIDFrom(programID)
}

// CreateObject creates object from programID based on interface type.
//
// Only supports IUnknown.
//
// Program ID can be either program ID or application string.
func CreateObject(programID string) (unknown *ole.IUnknown, err error) {
	classID, err := legacy.ClassIDFrom(programID)
	if err != nil {
		return
	}

	unknown, err = legacy.CreateInstance(classID, legacy.IID_IUnknown)
	if err != nil {
		return
	}

	return
}

// GetActiveObject retrieves active object for program ID and interface ID based
// on interface type.
//
// Only supports IUnknown.
//
// Program ID can be either program ID or application string.
func GetActiveObject(programID string) (unknown *ole.IUnknown, err error) {
	classID, err := legacy.ClassIDFrom(programID)
	if err != nil {
		return
	}

	unknown, err = legacy.GetActiveObject(classID, legacy.IID_IUnknown)
	if err != nil {
		return
	}

	return
}

// CallMethod calls method on IDispatch with parameters.
func CallMethod(disp *ole.IDispatch, name string, params ...interface{}) (result *legacy.VARIANT, err error) {
	return disp.InvokeWithOptionalArgs(name, legacy.DISPATCH_METHOD, params)
}

// MustCallMethod calls method on IDispatch with parameters or panics.
func MustCallMethod(disp *ole.IDispatch, name string, params ...interface{}) (result *legacy.VARIANT) {
	r, err := CallMethod(disp, name, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// GetProperty retrieves property from IDispatch.
func GetProperty(disp *ole.IDispatch, name string, params ...interface{}) (result *legacy.VARIANT, err error) {
	return disp.InvokeWithOptionalArgs(name, legacy.DISPATCH_PROPERTYGET, params)
}

// MustGetProperty retrieves property from IDispatch or panics.
func MustGetProperty(disp *ole.IDispatch, name string, params ...interface{}) (result *legacy.VARIANT) {
	r, err := GetProperty(disp, name, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// PutProperty mutates property.
func PutProperty(disp *ole.IDispatch, name string, params ...interface{}) (result *legacy.VARIANT, err error) {
	return disp.InvokeWithOptionalArgs(name, legacy.DISPATCH_PROPERTYPUT, params)
}

// MustPutProperty mutates property or panics.
func MustPutProperty(disp *ole.IDispatch, name string, params ...interface{}) (result *legacy.VARIANT) {
	r, err := PutProperty(disp, name, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// PutPropertyRef mutates property reference.
func PutPropertyRef(disp *ole.IDispatch, name string, params ...interface{}) (result *legacy.VARIANT, err error) {
	return disp.InvokeWithOptionalArgs(name, legacy.DISPATCH_PROPERTYPUTREF, params)
}

// MustPutPropertyRef mutates property reference or panics.
func MustPutPropertyRef(disp *ole.IDispatch, name string, params ...interface{}) (result *legacy.VARIANT) {
	r, err := PutPropertyRef(disp, name, params...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

func ForEach(disp *ole.IDispatch, f func(v *legacy.VARIANT) error) error {
	newEnum, err := disp.GetProperty("_NewEnum")
	if err != nil {
		return err
	}
	defer newEnum.Clear()

	enum, err := newEnum.ToIUnknown().IEnumVARIANT(legacy.IID_IEnumVariant)
	if err != nil {
		return err
	}
	defer enum.Release()

	for item, length, err := enum.Next(1); length > 0; item, length, err = enum.Next(1) {
		if err != nil {
			return err
		}
		if ferr := f(&item); ferr != nil {
			return ferr
		}
	}
	return nil
}
