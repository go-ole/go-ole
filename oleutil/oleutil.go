package oleutil

import "ole"
import "os"

func CreateDispatch(progId string) (dispatch *ole.IDispatch, err os.Error) {
	var clsid *ole.GUID
	clsid, err = ole.CLSIDFromProgID(progId)
	if err != nil {
		return
	}

	var unknown *ole.IUnknown
	unknown, err = ole.CreateInstance(clsid)
	if err != nil {
		return
	}

	dispatch, err = unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return
	}
	return
}

func CallMethod(disp *ole.IDispatch, name string, params ...interface{}) (result *ole.VARIANT, err os.Error) {
	var dispid []int32
	dispid, err = disp.GetIDsOfName([]string{name})
	result, err = disp.Invoke(dispid[0], ole.DISPATCH_METHOD, params...)
	return
}

func GetProperty(disp *ole.IDispatch, name string, params ...interface{}) (result *ole.VARIANT, err os.Error) {
	var dispid []int32
	dispid, err = disp.GetIDsOfName([]string{name})
	result, err = disp.Invoke(dispid[0], ole.DISPATCH_PROPERTYGET, params...)
	return
}

func PutProperty(disp *ole.IDispatch, name string, params ...interface{}) (result *ole.VARIANT, err os.Error) {
	var dispid []int32
	dispid, err = disp.GetIDsOfName([]string{name})
	result, err = disp.Invoke(dispid[0], ole.DISPATCH_PROPERTYPUT, params...)
	return
}
