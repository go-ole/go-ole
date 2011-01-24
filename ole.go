package ole

import "syscall"
import "unsafe"
import "os"

var (
	modkernel32, _ = syscall.LoadLibrary("kernel32.dll")
	modole32, _    = syscall.LoadLibrary("ole32.dll")
	modoleaut32, _ = syscall.LoadLibrary("oleaut32.dll")

	procCoInitialize, _       = syscall.GetProcAddress(modole32, "CoInitialize")
	procCoInitializeEx, _     = syscall.GetProcAddress(modole32, "CoInitializeEx")
	procCoCreateInstance, _   = syscall.GetProcAddress(modole32, "CoCreateInstance")
	procCLSIDFromProgID, _    = syscall.GetProcAddress(modole32, "CLSIDFromProgID")
	procCLSIDFromString, _    = syscall.GetProcAddress(modole32, "CLSIDFromString")
	procGetUserDefaultLCID, _ = syscall.GetProcAddress(modkernel32, "GetUserDefaultLCID")
	procCopyMemory, _         = syscall.GetProcAddress(modkernel32, "RtlMoveMemory")
	procVariantInit, _        = syscall.GetProcAddress(modoleaut32, "VariantInit")
	procSysAllocString, _     = syscall.GetProcAddress(modoleaut32, "SysAllocString")
	procSysFreeString, _      = syscall.GetProcAddress(modoleaut32, "SysFreeString")

	IID_NULL      = &GUID{0x00000000, 0x0000, 0x0000, [8]byte{0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}}
	IID_IUnknown  = &GUID{0x00000000, 0x0000, 0x0000, [8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}
	IID_IDispatch = &GUID{0x00020400, 0x0000, 0x0000, [8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}
)

const (
	CLSCTX_INPROC_SERVER   = 1
	CLSCTX_INPROC_HANDLER  = 2
	CLSCTX_LOCAL_SERVER    = 4
	CLSCTX_INPROC_SERVER16 = 8
	CLSCTX_REMOTE_SERVER   = 16
	CLSCTX_ALL             = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER
	CLSCTX_INPROC          = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER
	CLSCTX_SERVER          = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER
)

const (
	COINIT_APARTMENTTHREADED = 0x2
	COINIT_MULTITHREADED     = 0x0
	COINIT_DISABLE_OLE1DDE   = 0x4
	COINIT_SPEED_OVER_MEMORY = 0x8
)

const (
	DISPATCH_METHOD         = 1
	DISPATCH_PROPERTYGET    = 2
	DISPATCH_PROPERTYPUT    = 4
	DISPATCH_PROPERTYPUTREF = 8
)

type DISPPARAMS struct {
	rgvarg            uintptr
	rgdispidNamedArgs uintptr
	cArgs             uint32
	cNamedArgs        uint32
}

type GUID struct {
	Data1 uint32
	Data2 uint16
	Data3 uint16
	Data4 [8]byte
}

type IUnknown struct {
	lpVtbl   *pIUnknownVtbl
	pAddRef  uintptr
	pRelease uintptr
}

type pIUnknownVtbl struct {
	pQueryInterface uintptr
	pAddRef         uintptr
	pRelease        uintptr
}

type UnknownLike interface {
	QueryInterface(iid *GUID) (disp *IDispatch, err os.Error)
	AddRef() int32
	Release() int32
}

func (v *IUnknown) QueryInterface(iid *GUID) (disp *IDispatch, err os.Error) {
	disp, err = queryInterface(v, iid)
	return
}

func (v *IUnknown) AddRef() int32 {
	return addRef(v)
}

func (v *IUnknown) Release() int32 {
	return release(v)
}

type IDispatch struct {
	lpVtbl *pIDispatchVtbl
}

type pIDispatchVtbl struct {
	pQueryInterface   uintptr
	pAddRef           uintptr
	pRelease          uintptr
	pGetTypeInfoCount uintptr
	pGetTypeInfo      uintptr
	pGetIDsOfNames    uintptr
	pInvoke           uintptr
}

func (v *IDispatch) QueryInterface(iid *GUID) (disp *IDispatch, err os.Error) {
	disp, err = queryInterface((*IUnknown)(unsafe.Pointer(v)), iid)
	return
}

func (v *IDispatch) AddRef() int32 {
	return addRef((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IDispatch) Release() int32 {
	return release((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IDispatch) GetIDsOfName(names []string) (dispid []int32, err os.Error) {
	dispid, err = getIDsOfName(v, names)
	return
}

func (v *IDispatch) Invoke(dispid int32, dispatch int16, params ...interface{}) (result *VARIANT, err os.Error) {
	result, err = invoke(v, dispid, dispatch, params...)
	return
}

type VARIANT struct {
	VT         uint16 //  2
	wReserved1 uint16 //  4
	wReserved2 uint16 //  6
	wReserved3 uint16 //  8
	Val        int64  // 16
}

func (v *VARIANT) ToIUnknown() *IUnknown {
	return (*IUnknown)(unsafe.Pointer(uintptr(v.Val)))
}

func (v *VARIANT) ToIDispatch() *IDispatch {
	return (*IDispatch)(unsafe.Pointer(uintptr(v.Val)))
}

func CoInitialize(p uintptr) (err os.Error) {
	hr, _, _ := syscall.Syscall(uintptr(procCoInitialize), p, 0, 0)
	if hr != 0 {
		err = os.NewError(syscall.Errstr(int(hr)))
	}
	return
}

func CoInitializeEx(p uintptr, coinit uint32) (err os.Error) {
	hr, _, _ := syscall.Syscall(uintptr(procCoInitializeEx), p, uintptr(coinit), 0)
	if hr != 0 {
		err = os.NewError(syscall.Errstr(int(hr)))
	}
	return
}

func CLSIDFromProgID(progId string) (clsid *GUID, err os.Error) {
	var guid GUID
	hr, _, _ := syscall.Syscall(
		uintptr(procCLSIDFromProgID),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(progId))),
		uintptr(unsafe.Pointer(&guid)),
		0)
	if hr != 0 {
		err = os.NewError(syscall.Errstr(int(hr)))
	}
	clsid = &guid
	return
}

func CLSIDFromString(str string) (clsid *GUID, err os.Error) {
	var guid GUID
	hr, _, _ := syscall.Syscall(
		uintptr(procCLSIDFromString),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(str))),
		uintptr(unsafe.Pointer(&guid)),
		0)
	if hr != 0 {
		err = os.NewError(syscall.Errstr(int(hr)))
	}
	clsid = &guid
	return
}

func CreateInstance(clsid *GUID, iid *GUID) (unk *IUnknown, err os.Error) {
	if iid == nil {
		iid = IID_IUnknown
	}
	hr, _, _ := syscall.Syscall6(
		uintptr(procCoCreateInstance),
		uintptr(unsafe.Pointer(clsid)),
		0,
		CLSCTX_SERVER,
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&unk)),
		0)
	if hr != 0 {
		err = os.NewError(syscall.Errstr(int(hr)))
	}
	return
}

func queryInterface(unk *IUnknown, iid *GUID) (disp *IDispatch, err os.Error) {
	hr, _, _ := syscall.Syscall(
		unk.lpVtbl.pQueryInterface,
		uintptr(unsafe.Pointer(unk)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&disp)))
	if hr != 0 {
		err = os.NewError(syscall.Errstr(int(hr)))
	}
	return
}

func addRef(unk *IUnknown) int32 {
	ret, _, _ := syscall.Syscall(
		unk.lpVtbl.pAddRef,
		uintptr(unsafe.Pointer(unk)),
		0,
		0)
	return int32(ret)
}

func release(unk *IUnknown) int32 {
	ret, _, _ := syscall.Syscall(
		unk.lpVtbl.pRelease,
		uintptr(unsafe.Pointer(unk)),
		0,
		0)
	return int32(ret)
}

func getIDsOfName(disp *IDispatch, names []string) (dispid []int32, err os.Error) {
	wnames := make([]*uint16, len(names))
	dispid = make([]int32, len(names))
	for i := 0; i < len(names); i++ {
		wnames[i] = syscall.StringToUTF16Ptr(names[i])
	}
	hr, _, _ := syscall.Syscall6(
		disp.lpVtbl.pGetIDsOfNames,
		uintptr(unsafe.Pointer(disp)),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(unsafe.Pointer(&wnames[0])),
		uintptr(len(names)),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&dispid[0])))
	if hr != 0 {
		err = os.NewError(syscall.Errstr(int(hr)))
	}
	return
}

type EXCEPINFO struct {
	wCode             uint16
	wReserved         uint16
	bstrSource        *uint16
	bstrDescription   *uint16
	bstrHelpFile      *uint16
	dwHelpContext     uint32
	pvReserved        uintptr
	pfnDeferredFillIn uintptr
	scode             int32
}

func VariantInit(v *VARIANT) (err os.Error) {
	hr, _, _ := syscall.Syscall(
		uintptr(procVariantInit),
		uintptr(unsafe.Pointer(v)),
		0,
		0)
	if hr != 0 {
		err = os.NewError(syscall.Errstr(int(hr)))
	}
	return
}

func SysAllocString(v string) (ss *int16) {
	pss, _, _ := syscall.Syscall(
		uintptr(procSysAllocString),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(v))),
		0,
		0)
	ss = (*int16)(unsafe.Pointer(pss))
	return
}

func SysFreeString(v *int16) (err os.Error) {
	hr, _, _ := syscall.Syscall(
		uintptr(procSysAllocString),
		uintptr(unsafe.Pointer(v)),
		0,
		0)
	if hr != 0 {
		err = os.NewError(syscall.Errstr(int(hr)))
	}
	return
}

func copyMemory(dest unsafe.Pointer, src unsafe.Pointer, length uint32) {
	syscall.Syscall(
		uintptr(procCopyMemory),
		uintptr(dest),
		uintptr(src),
		uintptr(length))
}

const (
	VT_EMPTY           = 0x0
	VT_NULL            = 0x1
	VT_I2              = 0x2
	VT_I4              = 0x3
	VT_R4              = 0x4
	VT_R8              = 0x5
	VT_CY              = 0x6
	VT_DATE            = 0x7
	VT_BSTR            = 0x8
	VT_DISPATCH        = 0x9
	VT_ERROR           = 0xa
	VT_BOOL            = 0xb
	VT_VARIANT         = 0xc
	VT_UNKNOWN         = 0xd
	VT_DECIMAL         = 0xe
	VT_I1              = 0x10
	VT_UI1             = 0x11
	VT_UI2             = 0x12
	VT_UI4             = 0x13
	VT_I8              = 0x14
	VT_UI8             = 0x15
	VT_INT             = 0x16
	VT_UINT            = 0x17
	VT_VOID            = 0x18
	VT_HRESULT         = 0x19
	VT_PTR             = 0x1a
	VT_SAFEARRAY       = 0x1b
	VT_CARRAY          = 0x1c
	VT_USERDEFINED     = 0x1d
	VT_LPSTR           = 0x1e
	VT_LPWSTR          = 0x1f
	VT_RECORD          = 0x24
	VT_INT_PTR         = 0x25
	VT_UINT_PTR        = 0x26
	VT_FILETIME        = 0x40
	VT_BLOB            = 0x41
	VT_STREAM          = 0x42
	VT_STORAGE         = 0x43
	VT_STREAMED_OBJECT = 0x44
	VT_STORED_OBJECT   = 0x45
	VT_BLOB_OBJECT     = 0x46
	VT_CF              = 0x47
	VT_CLSID           = 0x48
	VT_BSTR_BLOB       = 0xfff
	VT_VECTOR          = 0x1000
	VT_ARRAY           = 0x2000
	VT_BYREF           = 0x4000
	VT_RESERVED        = 0x8000
	VT_ILLEGAL         = 0xffff
	VT_ILLEGALMASKED   = 0xfff
	VT_TYPEMASK        = 0xfff
)

const (
	DISPID_UNKNOWN     = -1
	DISPID_VALUE       = 0
	DISPID_PROPERTYPUT = -3
	DISPID_NEWENUM     = -4
	DISPID_EVALUATE    = -5
	DISPID_CONSTRUCTOR = -6
	DISPID_DESTRUCTOR  = -7
	DISPID_COLLECT     = -8
)

func invoke(disp *IDispatch, dispid int32, dispatch int16, params ...interface{}) (result *VARIANT, err os.Error) {
	var dispparams DISPPARAMS

	if dispatch&DISPATCH_PROPERTYPUT != 0 {
		dispnames := [1]int32{DISPID_PROPERTYPUT}
		dispparams.rgdispidNamedArgs = uintptr(unsafe.Pointer(&dispnames[0]))
		dispparams.cNamedArgs = 1
	}
	var vargs []VARIANT
	if len(params) > 0 {
		vargs = make([]VARIANT, len(params))
		for i, v := range params {
			//n := len(params)-i-1
			n := len(params) - i - 1
			VariantInit(&vargs[n])
			switch v.(type) {
			case bool:
				if v.(bool) {
					vargs[n] = VARIANT{VT_BOOL, 0, 0, 0, 0xffff}
				} else {
					vargs[n] = VARIANT{VT_BOOL, 0, 0, 0, 0}
				}
			case *bool:
				vargs[n] = VARIANT{VT_BOOL|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*bool))))}
			case byte:
				vargs[n] = VARIANT{VT_I1, 0, 0, 0, int64(v.(byte))}
			case *byte:
				vargs[n] = VARIANT{VT_I1|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*byte))))}
			case int16:
				vargs[n] = VARIANT{VT_I2, 0, 0, 0, int64(v.(int16))}
			case *int16:
				vargs[n] = VARIANT{VT_I2|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*int16))))}
			case uint16:
				vargs[n] = VARIANT{VT_UI2, 0, 0, 0, int64(v.(int16))}
			case *uint16:
				vargs[n] = VARIANT{VT_UI2|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*uint16))))}
			case int, int32:
				vargs[n] = VARIANT{VT_UI4, 0, 0, 0, int64(v.(uint))}
			case *int, *int32:
				vargs[n] = VARIANT{VT_I4|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*int))))}
			case uint, uint32:
				vargs[n] = VARIANT{VT_UI4, 0, 0, 0, int64(v.(uint))}
			case *uint, *uint32:
				vargs[n] = VARIANT{VT_UI4|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*uint))))}
			case int64:
				vargs[n] = VARIANT{VT_I8, 0, 0, 0, v.(int64)}
			case *int64:
				vargs[n] = VARIANT{VT_I8|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*int64))))}
			case uint64:
				vargs[n] = VARIANT{VT_UI8, 0, 0, 0, int64(v.(uint64))}
			case *uint64:
				vargs[n] = VARIANT{VT_UI8|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*uint64))))}
			case float32:
				vargs[n] = VARIANT{VT_R4, 0, 0, 0, int64(v.(float32))}
			case *float32:
				vargs[n] = VARIANT{VT_R4|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*float32))))}
			case float64:
				vargs[n] = VARIANT{VT_R8, 0, 0, 0, int64(v.(float64))}
			case *float64:
				vargs[n] = VARIANT{VT_R8|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*float64))))}
			case string:
				vargs[n] = VARIANT{VT_BSTR, 0, 0, 0, int64(uintptr(unsafe.Pointer(SysAllocString(v.(string)))))}
			case *string:
				vargs[n] = VARIANT{VT_BSTR|VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*string))))}
			case *IDispatch:
				vargs[n] = VARIANT{VT_DISPATCH, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*IDispatch))))}
			case **IDispatch:
				vargs[n] = VARIANT{VT_DISPATCH | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(**IDispatch))))}
			case nil:
				vargs[n] = VARIANT{VT_NULL, 0, 0, 0, 0}
			default:
				panic("unknown type")
			}
		}
		dispparams.rgvarg = uintptr(unsafe.Pointer(&vargs[0]))
		dispparams.cArgs = uint32(len(params))
	}

	var ret VARIANT
	var excepInfo EXCEPINFO
	VariantInit(&ret)
	hr, _, _ := syscall.Syscall9(
		disp.lpVtbl.pInvoke,
		uintptr(unsafe.Pointer(disp)),
		uintptr(dispid),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(GetUserDefaultLCID()),
		uintptr(dispatch),
		uintptr(unsafe.Pointer(&dispparams)),
		uintptr(unsafe.Pointer(&ret)),
		uintptr(unsafe.Pointer(&excepInfo)),
		0)
	if hr != 0 {
		err = os.NewError(syscall.Errstr(int(hr)))
		if excepInfo.bstrDescription != nil {
			bs := syscall.UTF16ToString((*[256]uint16)(unsafe.Pointer(excepInfo.bstrDescription))[:])
			println(bs)
		}
	}
	for _, varg := range vargs {
		if varg.VT == VT_BSTR {
			SysFreeString(((*int16)(unsafe.Pointer(uintptr(varg.Val)))))
		}
	}
	result = &ret
	return
}

func GetUserDefaultLCID() (lcid uint32) {
	ret, _, _ := syscall.Syscall(
		uintptr(procGetUserDefaultLCID),
		0,
		0,
		0)
	lcid = uint32(ret)
	return
}
