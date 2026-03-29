//go:build windows
// +build windows

package main

import (
	"github.com/go-ole/go-ole"
	"golang.org/x/sys/windows"
	"log"
	"syscall"
	"unsafe"
)

type EventReceiver struct {
	lpVtbl *EventReceiverVtbl
	ref    int32
	host   *ole.IDispatch
}

type EventReceiverVtbl struct {
	pQueryInterface   uintptr
	pAddRef           uintptr
	pRelease          uintptr
	pGetTypeInfoCount uintptr
	pGetTypeInfo      uintptr
	pGetIDsOfNames    uintptr
	pInvoke           uintptr
}

func QueryInterface(this *ole.IUnknown, iid *windows.GUID, punk **ole.IUnknown) uint32 {
	*punk = nil
	if *iid == ole.IID_IUnknown || *iid == ole.IID_IDispatch {
		AddRef(this)
		*punk = this
		return uint32(windows.S_OK)
	}
	if iid.String() == "{248DD893-BB45-11CF-9ABC-0080C7E7B78D}" {
		AddRef(this)
		*punk = this
		return uint32(windows.S_OK)
	}
	return uint32(windows.E_NOINTERFACE)
}

func AddRef(this *ole.IUnknown) int32 {
	pthis := (*EventReceiver)(unsafe.Pointer(this))
	pthis.ref++
	return pthis.ref
}

func Release(this *ole.IUnknown) int32 {
	pthis := (*EventReceiver)(unsafe.Pointer(this))
	pthis.ref--
	return pthis.ref
}

func GetIDsOfNames(this *ole.IUnknown, iid *windows.GUID, wnames []*uint16, namelen int, lcid int, pdisp []int32) uintptr {
	for n := 0; n < namelen; n++ {
		pdisp[n] = int32(n)
	}
	return uintptr(windows.S_OK)
}

func GetTypeInfoCount(pcount *int) uintptr {
	if pcount != nil {
		*pcount = 0
	}
	return uintptr(windows.S_OK)
}

func GetTypeInfo(ptypeif *uintptr) uintptr {
	return uintptr(windows.E_NOTIMPL)
}

func Invoke(this *ole.IDispatch, dispid int, riid *windows.GUID, lcid int, flags int16, dispparams *ole.DISPPARAMS, result *ole.VARIANT, pexcepinfo *ole.EXCEPINFO, nerr *uint) uintptr {
	switch dispid {
	case 0:
		log.Println("DataArrival")
		winsock := (*EventReceiver)(unsafe.Pointer(this)).host
		var data ole.VARIANT
		ole.VariantInit(&data)
		winsock.CallMethod("GetData", &data)
		bytes, _ := ole.ToSlice[[]byte]((*ole.SafeArray)(unsafe.Pointer(uintptr(data.Val))))
		s := string(bytes)
		println()
		println(s)
		println()
	case 1:
		log.Println("Connected")
		winsock := (*EventReceiver)(unsafe.Pointer(this)).host
		winsock.CallMethod("SendData", ole.StringToBStrVariant("GET / HTTP/1.0\r\n\r\n"))
	case 3:
		log.Println("SendProgress")
	case 4:
		log.Println("SendComplete")
	case 5:
		log.Println("Close")
		this.Release()
	case 6:
		log.Fatal("Error")
	default:
		log.Println(dispid)
	}
	return uintptr(windows.E_NOTIMPL)
}

func main() {
	ole.InitializeMultithreaded()
	defer ole.Uninitialize()
	ole.RegisterVariantConverters()

	clsid, _ := windows.GUIDFromString("{248DD896-BB45-11CF-9ABC-0080C7E7B78D}")
	unknown, err := ole.CreateInstance[ole.IUnknown](clsid, ole.IID_IUnknown)
	if err != nil {
		panic(err.Error())
	}
	defer unknown.Release()

	winsock, _ := ole.QueryInterfaceOnIUnknown[ole.IDispatch](unknown, ole.IID_IDispatch)
	defer winsock.Release()

	iid, _ := windows.GUIDFromString("{248DD893-BB45-11CF-9ABC-0080C7E7B78D}")

	container, err := ole.QueryIConnectionPointContainerFromIUnknown(unknown)
	if err != nil {
		panic(err.Error())
	}
	defer container.Release()

	point, err := container.FindConnectionPoint(iid)
	if err != nil {
		panic(err.Error())
	}
	defer point.Release()

	dest := &EventReceiver{}
	dest.lpVtbl = &EventReceiverVtbl{}
	dest.lpVtbl.pQueryInterface = syscall.NewCallback(QueryInterface)
	dest.lpVtbl.pAddRef = syscall.NewCallback(AddRef)
	dest.lpVtbl.pRelease = syscall.NewCallback(Release)
	dest.lpVtbl.pGetTypeInfoCount = syscall.NewCallback(GetTypeInfoCount)
	dest.lpVtbl.pGetTypeInfo = syscall.NewCallback(GetTypeInfo)
	dest.lpVtbl.pGetIDsOfNames = syscall.NewCallback(GetIDsOfNames)
	dest.lpVtbl.pInvoke = syscall.NewCallback(Invoke)
	dest.host = winsock

	sinkUnknown := (*ole.IUnknown)(unsafe.Pointer(dest))
	var sink ole.IsIUnknown = sinkUnknown
	_, err = point.Advise(&sink)
	if err != nil {
		panic(err.Error())
	}

	_, err = winsock.CallMethod("Connect", ole.StringToBStrVariant("127.0.0.1"), ole.Int32ToVariant(80))
	if err != nil {
		log.Fatal(err)
	}

	var m windows.Msg
	for dest.ref != 0 {
		windows.GetMessage(&m, 0, 0, 0)
		windows.DispatchMessage(&m)
	}
}
