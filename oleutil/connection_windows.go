//go:build windows
// +build windows

package oleutil

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/legacy"
	"reflect"
	"syscall"
	"unsafe"
)

// ConnectObject creates a connection point between two services for communication.
func ConnectObject(disp *ole.IDispatch, iid *legacy.GUID, idisp interface{}) (cookie uint32, err error) {
	unknown, err := disp.QueryInterface(legacy.IID_IConnectionPointContainer)
	if err != nil {
		return
	}

	container := (*legacy.IConnectionPointContainer)(unsafe.Pointer(unknown))
	var point *legacy.IConnectionPoint
	err = container.FindConnectionPoint(iid, &point)
	if err != nil {
		return
	}
	if edisp, ok := idisp.(*ole.IUnknown); ok {
		cookie, err = point.Advise(edisp)
		container.Release()
		if err != nil {
			return
		}
	}
	rv := reflect.ValueOf(disp).Elem()
	if rv.Type().Kind() == reflect.Struct {
		dest := &stdDispatch{}
		dest.lpVtbl = &stdDispatchVtbl{}
		dest.lpVtbl.pQueryInterface = syscall.NewCallback(dispQueryInterface)
		dest.lpVtbl.pAddRef = syscall.NewCallback(dispAddRef)
		dest.lpVtbl.pRelease = syscall.NewCallback(dispRelease)
		dest.lpVtbl.pGetTypeInfoCount = syscall.NewCallback(dispGetTypeInfoCount)
		dest.lpVtbl.pGetTypeInfo = syscall.NewCallback(dispGetTypeInfo)
		dest.lpVtbl.pGetIDsOfNames = syscall.NewCallback(dispGetIDsOfNames)
		dest.lpVtbl.pInvoke = syscall.NewCallback(dispInvoke)
		dest.iface = disp
		dest.iid = iid
		cookie, err = point.Advise((*ole.IUnknown)(unsafe.Pointer(dest)))
		container.Release()
		if err != nil {
			point.Release()
			return
		}
		return
	}

	container.Release()

	return 0, legacy.NewError(legacy.E_INVALIDARG)
}
