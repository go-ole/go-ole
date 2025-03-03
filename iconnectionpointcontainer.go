//go:build windows

package ole

import (
	"golang.org/x/sys/windows"
	"unsafe"
)

type IConnectionPointContainer struct {
	// IUnknown
	QueryInterface uintptr
	addRef         uintptr
	release        uintptr
	// IConnectionPointContainer
	enumConnectionPoints uintptr
	findConnectionPoint  uintptr
}

func (obj *IConnectionPointContainer) EnumConnectionPoints(points interface{}) error {
	return NewError(E_NOTIMPL)
}

func (obj *IConnectionPointContainer) FindConnectionPoint(iid windows.GUID) (point *IConnectionPoint, err error) {
	hr, _, _ := windows.Syscall(
		obj.findConnectionPoint,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&point)))
	if hr != windows.S_OK {
		err = hr
	}
	return
}

func QueryIConnectionPointContainerFromIUnknown(unknown *IsIUnknown) (obj *IConnectionPointContainer, err error) {
	if unknown == nil {
		return nil, ComInterfaceIsNilPointer
	}

	enum, err = QueryInterfaceOnIUnknown[IConnectionPointContainer](unknown, IID_IConnectionPointContainer)
	if err != nil {
		return nil, err
	}
	return
}

func (obj *IDispatch) ConnectObject(interfaceId windows.GUID, unknown IsIUnknown) (cookie uint32, err error) {
	container, err := QueryIConnectionPointContainerFromIUnknown(obj)
	if err != nil {
		return
	}

	defer container.Release()

	point, err := container.FindConnectionPoint(interfaceId)
	if err != nil {
		return
	}

	defer point.Release()

	cookie, err = point.Advise(unknown)
	rv := reflect.ValueOf(obj).Elem()
	if rv.Type().Kind() == reflect.Struct {
		dest := &stdDispatch{}
		dest.lpVtbl = &stdDispatchVtbl{}
		dest.lpVtbl.pQueryInterface = windows.NewCallback(dispQueryInterface)
		dest.lpVtbl.pAddRef = windows.NewCallback(dispAddRef)
		dest.lpVtbl.pRelease = windows.NewCallback(dispRelease)
		dest.lpVtbl.pGetTypeInfoCount = windows.NewCallback(dispGetTypeInfoCount)
		dest.lpVtbl.pGetTypeInfo = windows.NewCallback(dispGetTypeInfo)
		dest.lpVtbl.pGetIDsOfNames = windows.NewCallback(dispGetIDsOfNames)
		dest.lpVtbl.pInvoke = windows.NewCallback(dispInvoke)
		dest.iface = obj
		dest.iid = interfaceId
		cookie, err = point.Advise((*IUnknown)(unsafe.Pointer(dest)))
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

// ConnectObject creates a connection point between two services for communication.
func ConnectObject(dispatch *IDispatch, iid windows.GUID, idisp interface{}) (cookie uint32, err error) {
	unknown, err := dispatch.QueryInterface(IID_IConnectionPointContainer)
	if err != nil {
		return
	}

	container := (*IConnectionPointContainer)(unsafe.Pointer(unknown))
	var point *IConnectionPoint
	err = container.FindConnectionPoint(iid, &point)
	if err != nil {
		return
	}
	if edisp, ok := idisp.(*IUnknown); ok {
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
		cookie, err = point.Advise((*IUnknown)(unsafe.Pointer(dest)))
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
