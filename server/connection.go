//go:build windows

package server

import (
	"github.com/go-ole/go-ole"
	"reflect"
	"unsafe"

	"golang.org/x/sys/windows"
)

type stdDispatch[T any] struct {
	lpVtbl  *stdDispatchVtbl
	ref     int32
	iid     windows.GUID
	iface   T
	funcMap map[string]int32
}

type stdDispatchVtbl struct {
	pQueryInterface   uintptr
	pAddRef           uintptr
	pRelease          uintptr
	pGetTypeInfoCount uintptr
	pGetTypeInfo      uintptr
	pGetIDsOfNames    uintptr
	pInvoke           uintptr
}

// ConnectObject creates a connection point between two services for communication.
func MakeStdDispatch(disp *ole.IDispatch, iid windows.GUID, idisp interface{}) (cookie uint32, err error) {
	unknown, err := disp.QueryInterface(legacy.IID_IConnectionPointContainer)
	if err != nil {
		return
	}

	container := (*ole.IConnectionPointContainer)(unsafe.Pointer(unknown))
	var point *ole.IConnectionPoint
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

	return 0, windows.E_INVALIDARG
}

func dispQueryInterface(this *interface{}, iid windows.GUID, punk **IUnknown) uint32 {
	pthis := (*stdDispatch)(unsafe.Pointer(this))
	*punk = nil
	if cmp.Equal(iid, IID_IUnknown) || cmp.Equal(iid, IID_IDispatch) {
		dispAddRef(this)
		*punk = this
		return windows.S_OK
	}
	if cmp.Equal(iid, pthis.iid) {
		dispAddRef(this)
		*punk = this
		return windows.S_OK
	}
	return windows.E_NOINTERFACE
}

func dispAddRef(this *interface{}) int32 {
	pthis := (*stdDispatch)(unsafe.Pointer(this))
	pthis.ref++
	return pthis.ref
}

func dispRelease(this *interface{}) int32 {
	pthis := (*stdDispatch)(unsafe.Pointer(this))
	pthis.ref--
	return pthis.ref
}

func dispGetIDsOfNames(this *interface{}, iid windows.GUID, wnames []*uint16, namelen int, lcid int, pdisp []int32) uintptr {
	pthis := (*stdDispatch)(unsafe.Pointer(this))
	names := make([]string, len(wnames))
	for i := 0; i < len(names); i++ {
		names[i] = LpOleStrToString(wnames[i])
	}
	for n := 0; n < namelen; n++ {
		if id, ok := pthis.funcMap[names[n]]; ok {
			pdisp[n] = id
		}
	}
	return windows.S_OK
}

func dispGetTypeInfoCount(pcount *int) uintptr {
	if pcount != nil {
		*pcount = 0
	}
	return windows.S_OK
}

func dispGetTypeInfo(ptypeif *uintptr) uintptr {
	return windows.E_NOTIMPL
}

func dispInvoke(this *interface{}, dispid int32, riid windows.GUID, lcid int, flags int16, dispparams *DISPPARAMS, result *VARIANT, pexcepinfo *EXCEPINFO, nerr *uint) uintptr {
	pthis := (*stdDispatch)(unsafe.Pointer(this))
	found := ""
	for name, id := range pthis.funcMap {
		if id == dispid {
			found = name
		}
	}
	if found != "" {
		rv := reflect.ValueOf(pthis.iface).Elem()
		rm := rv.MethodByName(found)
		rr := rm.Call([]reflect.Value{})
		println(len(rr))
		return windows.S_OK
	}
	return windows.E_NOTIMPL
}
