package ole

import (
	"fmt"
	"syscall"
	"unicode/utf16"
	"unsafe"
)

var (
	modkernel32, _ = syscall.LoadLibrary("kernel32.dll")
	modole32, _    = syscall.LoadLibrary("ole32.dll")
	modoleaut32, _ = syscall.LoadLibrary("oleaut32.dll")
	modmsvcrt, _   = syscall.LoadLibrary("msvcrt.dll")
	moduser32, _   = syscall.LoadLibrary("user32.dll")

	procCoInitialize, _       = syscall.GetProcAddress(modole32, "CoInitialize")
	procCoInitializeEx, _     = syscall.GetProcAddress(modole32, "CoInitializeEx")
	procCoUninitialize, _     = syscall.GetProcAddress(modole32, "CoUninitialize")
	procCoCreateInstance, _   = syscall.GetProcAddress(modole32, "CoCreateInstance")
	procCLSIDFromProgID, _    = syscall.GetProcAddress(modole32, "CLSIDFromProgID")
	procCLSIDFromString, _    = syscall.GetProcAddress(modole32, "CLSIDFromString")
	procStringFromCLSID, _    = syscall.GetProcAddress(modole32, "StringFromCLSID")
	procStringFromIID, _      = syscall.GetProcAddress(modole32, "StringFromIID")
	procIIDFromString, _      = syscall.GetProcAddress(modole32, "IIDFromString")
	procGetUserDefaultLCID, _ = syscall.GetProcAddress(modkernel32, "GetUserDefaultLCID")
	procCopyMemory, _         = syscall.GetProcAddress(modkernel32, "RtlMoveMemory")
	procVariantInit, _        = syscall.GetProcAddress(modoleaut32, "VariantInit")
	procSysAllocString, _     = syscall.GetProcAddress(modoleaut32, "SysAllocString")
	procSysFreeString, _      = syscall.GetProcAddress(modoleaut32, "SysFreeString")
	procSysStringLen, _       = syscall.GetProcAddress(modoleaut32, "SysStringLen")
	procCreateDispTypeInfo, _ = syscall.GetProcAddress(modoleaut32, "CreateDispTypeInfo")
	procCreateStdDispatch, _  = syscall.GetProcAddress(modoleaut32, "CreateStdDispatch")
	procGetActiveObject, _    = syscall.GetProcAddress(modoleaut32, "GetActiveObject")

	procGetMessageW, _      = syscall.GetProcAddress(moduser32, "GetMessageW")
	procDispatchMessageW, _ = syscall.GetProcAddress(moduser32, "DispatchMessageW")

	IID_NULL                      = &GUID{0x00000000, 0x0000, 0x0000, [8]byte{0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}}
	IID_IUnknown                  = &GUID{0x00000000, 0x0000, 0x0000, [8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}
	IID_IDispatch                 = &GUID{0x00020400, 0x0000, 0x0000, [8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}
	IID_IConnectionPointContainer = &GUID{0xB196B284, 0xBAB4, 0x101A, [8]byte{0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07}}
	IID_IConnectionPoint          = &GUID{0xB196B286, 0xBAB4, 0x101A, [8]byte{0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07}}
	IID_IProvideClassInfo         = &GUID{0xb196b283, 0xbab4, 0x101a, [8]byte{0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07}}
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

const (
	S_OK           = 0x00000000
	E_UNEXPECTED   = 0x8000FFFF
	E_NOTIMPL      = 0x80004001
	E_OUTOFMEMORY  = 0x8007000E
	E_INVALIDARG   = 0x80070057
	E_NOINTERFACE  = 0x80004002
	E_POINTER      = 0x80004003
	E_HANDLE       = 0x80070006
	E_ABORT        = 0x80004004
	E_FAIL         = 0x80004005
	E_ACCESSDENIED = 0x80070005
	E_PENDING      = 0x8000000A
)

type OleError uintptr

func errstr(errno int) string {
	// ask windows for the remaining errors
	var flags uint32 = syscall.FORMAT_MESSAGE_FROM_SYSTEM | syscall.FORMAT_MESSAGE_ARGUMENT_ARRAY | syscall.FORMAT_MESSAGE_IGNORE_INSERTS
	b := make([]uint16, 300)
	n, err := syscall.FormatMessage(flags, 0, uint32(errno), 0, b, nil)
	if err != nil {
		return fmt.Sprintf("error %d (FormatMessage failed with: %v)", errno, err)
	}
	// trim terminating \r and \n
	for ; n > 0 && (b[n-1] == '\n' || b[n-1] == '\r'); n-- {
	}
	return string(utf16.Decode(b[:n]))
}

func NewError(hr uintptr) OleError {
	return OleError(hr)
}

func (v OleError) Code() uintptr {
	return uintptr(v)
}

func (v OleError) Error() string {
	return errstr(int(v))
}

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
	lpVtbl *pIUnknownVtbl
}

type pIUnknownVtbl struct {
	pQueryInterface uintptr
	pAddRef         uintptr
	pRelease        uintptr
}

type UnknownLike interface {
	QueryInterface(iid *GUID) (disp *IDispatch, err error)
	AddRef() int32
	Release() int32
}

func (v *IUnknown) QueryInterface(iid *GUID) (disp *IDispatch, err error) {
	disp, err = queryInterface(v, iid)
	return
}

func (v *IUnknown) MustQueryInterface(iid *GUID) (disp *IDispatch) {
	disp, _ = queryInterface(v, iid)
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

func (v *IDispatch) QueryInterface(iid *GUID) (disp *IDispatch, err error) {
	disp, err = queryInterface((*IUnknown)(unsafe.Pointer(v)), iid)
	return
}

func (v *IDispatch) MustQueryInterface(iid *GUID) (disp *IDispatch) {
	disp, _ = queryInterface((*IUnknown)(unsafe.Pointer(v)), iid)
	return
}

func (v *IDispatch) AddRef() int32 {
	return addRef((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IDispatch) Release() int32 {
	return release((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IDispatch) GetIDsOfName(names []string) (dispid []int32, err error) {
	dispid, err = getIDsOfName(v, names)
	return
}

func (v *IDispatch) Invoke(dispid int32, dispatch int16, params ...interface{}) (result *VARIANT, err error) {
	result, err = invoke(v, dispid, dispatch, params...)
	return
}

func (v *IDispatch) GetTypeInfoCount() (c uint32, err error) {
	c, err = getTypeInfoCount(v)
	return
}

func (v *IDispatch) GetTypeInfo() (tinfo *ITypeInfo, err error) {
	tinfo, err = getTypeInfo(v)
	return
}

type ITypeInfo struct {
	lpVtbl *pITypeInfoVtbl
}

type pITypeInfoVtbl struct {
	pQueryInterface       uintptr
	pAddRef               uintptr
	pRelease              uintptr
	pGetTypeAttr          uintptr
	pGetTypeComp          uintptr
	pGetFuncDesc          uintptr
	pGetVarDesc           uintptr
	pGetNames             uintptr
	pGetRefTypeOfImplType uintptr
	pGetImplTypeFlags     uintptr
	pGetIDsOfNames        uintptr
	pInvoke               uintptr
	pGetDocumentation     uintptr
	pGetDllEntry          uintptr
	pGetRefTypeInfo       uintptr
	pAddressOfMember      uintptr
	pCreateInstance       uintptr
	pGetMops              uintptr
	pGetContainingTypeLib uintptr
	pReleaseTypeAttr      uintptr
	pReleaseFuncDesc      uintptr
	pReleaseVarDesc       uintptr
}

func (v *ITypeInfo) QueryInterface(iid *GUID) (disp *IDispatch, err error) {
	disp, err = queryInterface((*IUnknown)(unsafe.Pointer(v)), iid)
	return
}

func (v *ITypeInfo) AddRef() int32 {
	return addRef((*IUnknown)(unsafe.Pointer(v)))
}

func (v *ITypeInfo) Release() int32 {
	return release((*IUnknown)(unsafe.Pointer(v)))
}

func (v *ITypeInfo) GetTypeAttr() (tattr *TYPEATTR, err error) {
	hr, _, _ := syscall.Syscall(
		uintptr(v.lpVtbl.pGetTypeAttr),
		2,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(&tattr)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

type IConnectionPointContainer struct {
	lpVtbl *pIConnectionPointContainerVtbl
}

type pIConnectionPointContainerVtbl struct {
	pQueryInterface       uintptr
	pAddRef               uintptr
	pRelease              uintptr
	pEnumConnectionPoints uintptr
	pFindConnectionPoint  uintptr
}

func (v *IConnectionPointContainer) QueryInterface(iid *GUID) (disp *IDispatch, err error) {
	disp, err = queryInterface((*IUnknown)(unsafe.Pointer(v)), iid)
	return
}

func (v *IConnectionPointContainer) AddRef() int32 {
	return addRef((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IConnectionPointContainer) Release() int32 {
	return release((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IConnectionPointContainer) EnumConnectionPoints(points interface{}) (err error) {
	err = NewError(E_NOTIMPL)
	return
}

func (v *IConnectionPointContainer) FindConnectionPoint(iid *GUID, point **IConnectionPoint) (err error) {
	hr, _, _ := syscall.Syscall(
		uintptr(v.lpVtbl.pFindConnectionPoint),
		3,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(point)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

type IProvideClassInfo struct {
	lpVtbl *pIProvideClassInfoVtbl
}

type pIProvideClassInfoVtbl struct {
	pQueryInterface uintptr
	pAddRef         uintptr
	pRelease        uintptr
	pGetClassInfo   uintptr
}

func (v *IProvideClassInfo) QueryInterface(iid *GUID) (disp *IDispatch, err error) {
	disp, err = queryInterface((*IUnknown)(unsafe.Pointer(v)), iid)
	return
}

func (v *IProvideClassInfo) AddRef() int32 {
	return addRef((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IProvideClassInfo) Release() int32 {
	return release((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IProvideClassInfo) GetClassInfo() (cinfo *ITypeInfo, err error) {
	cinfo, err = getClassInfo(v)
	return
}

type IConnectionPoint struct {
	lpVtbl *pIConnectionPointVtbl
}

type pIConnectionPointVtbl struct {
	pQueryInterface              uintptr
	pAddRef                      uintptr
	pRelease                     uintptr
	pGetConnectionInterface      uintptr
	pGetConnectionPointContainer uintptr
	pAdvise                      uintptr
	pUnadvise                    uintptr
	pEnumConnections             uintptr
}

func (v *IConnectionPoint) QueryInterface(iid *GUID) (disp *IDispatch, err error) {
	disp, err = queryInterface((*IUnknown)(unsafe.Pointer(v)), iid)
	return
}

func (v *IConnectionPoint) AddRef() int32 {
	return addRef((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IConnectionPoint) Release() int32 {
	return release((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IConnectionPoint) GetConnectionInterface(piid **GUID) int32 {
	return release((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IConnectionPoint) Advise(unknown *IUnknown) (cookie uint32, err error) {
	hr, _, _ := syscall.Syscall(
		uintptr(v.lpVtbl.pAdvise),
		3,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(unknown)),
		uintptr(unsafe.Pointer(&cookie)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func (v *IConnectionPoint) Unadvise(cookie uint32) (err error) {
	hr, _, _ := syscall.Syscall(
		uintptr(v.lpVtbl.pAdvise),
		2,
		uintptr(unsafe.Pointer(v)),
		uintptr(cookie),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func (v *IConnectionPoint) EnumConnections(p *unsafe.Pointer) (err error) {
	return NewError(E_NOTIMPL)
}

type VARIANT struct {
	VT         uint16 //  2
	wReserved1 uint16 //  4
	wReserved2 uint16 //  6
	wReserved3 uint16 //  8
	Val        int64  // 16
}

type SAFEARRAYBOUND struct {
	CElements uint32
	LLbound   int32
}

type SAFEARRAY struct {
	CDims      uint16
	FFeatures  uint16
	CbElements uint32
	CLocks     uint32
	PvData     uint32
	RgsaBound  SAFEARRAYBOUND
}

func BytePtrToString(p *byte) string {
	a := (*[10000]uint8)(unsafe.Pointer(p))
	i := 0
	for a[i] != 0 {
		i++
	}
	return string(a[:i])
}

func UTF16PtrToString(p *uint16) string {
	a := (*[10000]uint16)(unsafe.Pointer(p))
	i := 0
	for a[i] != 0 {
		i++
	}
	return string(utf16.Decode(a[:i]))
}

func (v *VARIANT) ToIUnknown() *IUnknown {
	return (*IUnknown)(unsafe.Pointer(uintptr(v.Val)))
}

func (v *VARIANT) ToIDispatch() *IDispatch {
	return (*IDispatch)(unsafe.Pointer(uintptr(v.Val)))
}

func (v *VARIANT) ToString() string {
	return UTF16PtrToString(*(**uint16)(unsafe.Pointer(&v.Val)))
}

func CoInitialize(p uintptr) (err error) {
	hr, _, _ := syscall.Syscall(uintptr(procCoInitialize), 1, p, 0, 0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func CoInitializeEx(p uintptr, coinit uint32) (err error) {
	hr, _, _ := syscall.Syscall(uintptr(procCoInitializeEx), 2, p, uintptr(coinit), 0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func CoUninitialize() {
	syscall.Syscall(uintptr(procCoUninitialize), 0, 0, 0, 0)
}

func CLSIDFromProgID(progId string) (clsid *GUID, err error) {
	var guid GUID
	hr, _, _ := syscall.Syscall(
		uintptr(procCLSIDFromProgID),
		2,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(progId))),
		uintptr(unsafe.Pointer(&guid)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	clsid = &guid
	return
}

func CLSIDFromString(str string) (clsid *GUID, err error) {
	var guid GUID
	hr, _, _ := syscall.Syscall(
		uintptr(procCLSIDFromString),
		2,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(str))),
		uintptr(unsafe.Pointer(&guid)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	clsid = &guid
	return
}

func StringFromCLSID(clsid *GUID) (str string, err error) {
	var p *uint16
	hr, _, _ := syscall.Syscall(
		uintptr(procStringFromCLSID),
		2,
		uintptr(unsafe.Pointer(clsid)),
		uintptr(unsafe.Pointer(&p)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	str = UTF16PtrToString(p)
	return
}

func IIDFromString(progId string) (clsid *GUID, err error) {
	var guid GUID
	hr, _, _ := syscall.Syscall(
		uintptr(procIIDFromString),
		2,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(progId))),
		uintptr(unsafe.Pointer(&guid)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	clsid = &guid
	return
}

func StringFromIID(iid *GUID) (str string, err error) {
	var p *uint16
	hr, _, _ := syscall.Syscall(
		uintptr(procStringFromIID),
		2,
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&p)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	str = UTF16PtrToString(p)
	return
}

func CreateInstance(clsid *GUID, iid *GUID) (unk *IUnknown, err error) {
	if iid == nil {
		iid = IID_IUnknown
	}
	hr, _, _ := syscall.Syscall6(
		uintptr(procCoCreateInstance),
		5,
		uintptr(unsafe.Pointer(clsid)),
		0,
		CLSCTX_SERVER,
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&unk)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func GetActiveObject(clsid *GUID, iid *GUID) (unk *IUnknown, err error) {
	if iid == nil {
		iid = IID_IUnknown
	}
	hr, _, _ := syscall.Syscall(
		uintptr(procGetActiveObject),
		3,
		uintptr(unsafe.Pointer(clsid)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&unk)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func queryInterface(unk *IUnknown, iid *GUID) (disp *IDispatch, err error) {
	hr, _, _ := syscall.Syscall(
		unk.lpVtbl.pQueryInterface,
		3,
		uintptr(unsafe.Pointer(unk)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&disp)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func addRef(unk *IUnknown) int32 {
	ret, _, _ := syscall.Syscall(
		unk.lpVtbl.pAddRef,
		1,
		uintptr(unsafe.Pointer(unk)),
		0,
		0)
	return int32(ret)
}

func release(unk *IUnknown) int32 {
	ret, _, _ := syscall.Syscall(
		unk.lpVtbl.pRelease,
		1,
		uintptr(unsafe.Pointer(unk)),
		0,
		0)
	return int32(ret)
}

func getIDsOfName(disp *IDispatch, names []string) (dispid []int32, err error) {
	wnames := make([]*uint16, len(names))
	for i := 0; i < len(names); i++ {
		wnames[i] = syscall.StringToUTF16Ptr(names[i])
	}
	dispid = make([]int32, len(names))
	hr, _, _ := syscall.Syscall6(
		disp.lpVtbl.pGetIDsOfNames,
		6,
		uintptr(unsafe.Pointer(disp)),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(unsafe.Pointer(&wnames[0])),
		uintptr(len(names)),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&dispid[0])))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func getTypeInfoCount(disp *IDispatch) (c uint32, err error) {
	hr, _, _ := syscall.Syscall(
		disp.lpVtbl.pGetTypeInfoCount,
		2,
		uintptr(unsafe.Pointer(disp)),
		uintptr(unsafe.Pointer(&c)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func getTypeInfo(disp *IDispatch) (tinfo *ITypeInfo, err error) {
	hr, _, _ := syscall.Syscall(
		disp.lpVtbl.pGetTypeInfo,
		3,
		uintptr(unsafe.Pointer(disp)),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&tinfo)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func getClassInfo(disp *IProvideClassInfo) (tinfo *ITypeInfo, err error) {
	hr, _, _ := syscall.Syscall(
		disp.lpVtbl.pGetClassInfo,
		2,
		uintptr(unsafe.Pointer(disp)),
		uintptr(unsafe.Pointer(&tinfo)),
		0)
	if hr != 0 {
		err = NewError(hr)
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

func VariantInit(v *VARIANT) (err error) {
	hr, _, _ := syscall.Syscall(
		uintptr(procVariantInit),
		1,
		uintptr(unsafe.Pointer(v)),
		0,
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func SysAllocString(v string) (ss *int16) {
	pss, _, _ := syscall.Syscall(
		uintptr(procSysAllocString),
		1,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(v))),
		0,
		0)
	ss = (*int16)(unsafe.Pointer(pss))
	return
}

func SysFreeString(v *int16) (err error) {
	hr, _, _ := syscall.Syscall(
		uintptr(procSysFreeString),
		1,
		uintptr(unsafe.Pointer(v)),
		0,
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func SysStringLen(v *int16) uint32 {
	l, _, _ := syscall.Syscall(
		uintptr(procSysStringLen),
		1,
		uintptr(unsafe.Pointer(v)),
		0,
		0)
	return uint32(l)
}

type PARAMDATA struct {
	Name *int16
	Vt   uint16
}

type METHODDATA struct {
	Name     *uint16
	Data     *PARAMDATA
	Dispid   int32
	Meth     uint32
	CC       int32
	CArgs    uint32
	Flags    uint16
	VtReturn uint32
}

type INTERFACEDATA struct {
	MethodData *METHODDATA
	CMembers   uint32
}

const (
	CC_FASTCALL = iota
	CC_CDECL
	CC_MSCPASCAL
	CC_PASCAL = CC_MSCPASCAL
	CC_MACPASCAL
	CC_STDCALL
	CC_FPFASTCALL
	CC_SYSCALL
	CC_MPWCDECL
	CC_MPWPASCAL
	CC_MAX = CC_MPWPASCAL
)

func CreateStdDispatch(unk *IUnknown, v uintptr, ptinfo *IUnknown) (disp *IDispatch, err error) {
	hr, _, _ := syscall.Syscall6(
		uintptr(procCreateStdDispatch),
		4,
		uintptr(unsafe.Pointer(unk)),
		v,
		uintptr(unsafe.Pointer(ptinfo)),
		uintptr(unsafe.Pointer(&disp)),
		0,
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func CreateDispTypeInfo(idata *INTERFACEDATA) (pptinfo *IUnknown, err error) {
	hr, _, _ := syscall.Syscall(
		uintptr(procCreateDispTypeInfo),
		3,
		uintptr(unsafe.Pointer(idata)),
		uintptr(GetUserDefaultLCID()),
		uintptr(unsafe.Pointer(&pptinfo)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func copyMemory(dest unsafe.Pointer, src unsafe.Pointer, length uint32) {
	syscall.Syscall(
		uintptr(procCopyMemory),
		3,
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

func invoke(disp *IDispatch, dispid int32, dispatch int16, params ...interface{}) (result *VARIANT, err error) {
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
				vargs[n] = VARIANT{VT_BOOL | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*bool))))}
			case byte:
				vargs[n] = VARIANT{VT_I1, 0, 0, 0, int64(v.(byte))}
			case *byte:
				vargs[n] = VARIANT{VT_I1 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*byte))))}
			case int16:
				vargs[n] = VARIANT{VT_I2, 0, 0, 0, int64(v.(int16))}
			case *int16:
				vargs[n] = VARIANT{VT_I2 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*int16))))}
			case uint16:
				vargs[n] = VARIANT{VT_UI2, 0, 0, 0, int64(v.(int16))}
			case *uint16:
				vargs[n] = VARIANT{VT_UI2 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*uint16))))}
			case int, int32:
				vargs[n] = VARIANT{VT_UI4, 0, 0, 0, int64(v.(int))}
			case *int, *int32:
				vargs[n] = VARIANT{VT_I4 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*int))))}
			case uint, uint32:
				vargs[n] = VARIANT{VT_UI4, 0, 0, 0, int64(v.(uint))}
			case *uint, *uint32:
				vargs[n] = VARIANT{VT_UI4 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*uint))))}
			case int64:
				vargs[n] = VARIANT{VT_I8, 0, 0, 0, int64(v.(int64))}
			case *int64:
				vargs[n] = VARIANT{VT_I8 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*int64))))}
			case uint64:
				vargs[n] = VARIANT{VT_UI8, 0, 0, 0, v.(int64)}
			case *uint64:
				vargs[n] = VARIANT{VT_UI8 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*uint64))))}
			case float32:
				vargs[n] = VARIANT{VT_R4, 0, 0, 0, int64(v.(float32))}
			case *float32:
				vargs[n] = VARIANT{VT_R4 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*float32))))}
			case float64:
				vargs[n] = VARIANT{VT_R8, 0, 0, 0, int64(v.(float64))}
			case *float64:
				vargs[n] = VARIANT{VT_R8 | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*float64))))}
			case string:
				vargs[n] = VARIANT{VT_BSTR, 0, 0, 0, int64(uintptr(unsafe.Pointer(SysAllocString(v.(string)))))}
			case *string:
				vargs[n] = VARIANT{VT_BSTR | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*string))))}
			case *IDispatch:
				vargs[n] = VARIANT{VT_DISPATCH, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*IDispatch))))}
			case **IDispatch:
				vargs[n] = VARIANT{VT_DISPATCH | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(**IDispatch))))}
			case nil:
				vargs[n] = VARIANT{VT_NULL, 0, 0, 0, 0}
			case *VARIANT:
				vargs[n] = VARIANT{VT_VARIANT | VT_BYREF, 0, 0, 0, int64(uintptr(unsafe.Pointer(v.(*VARIANT))))}
			default:
				panic("unknown type")
			}
		}
		dispparams.rgvarg = uintptr(unsafe.Pointer(&vargs[0]))
		dispparams.cArgs = uint32(len(params))
	}

	result = new(VARIANT)
	var excepInfo EXCEPINFO
	VariantInit(result)
	hr, _, _ := syscall.Syscall9(
		disp.lpVtbl.pInvoke,
		9,
		uintptr(unsafe.Pointer(disp)),
		uintptr(dispid),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(GetUserDefaultLCID()),
		uintptr(dispatch),
		uintptr(unsafe.Pointer(&dispparams)),
		uintptr(unsafe.Pointer(result)),
		uintptr(unsafe.Pointer(&excepInfo)),
		0)
	if hr != 0 {
		err = NewError(hr)
		if excepInfo.bstrDescription != nil {
			bs := UTF16PtrToString(excepInfo.bstrDescription)
			println(bs)
		}
	}
	for _, varg := range vargs {
		if varg.VT == VT_BSTR && varg.Val != 0 {
			SysFreeString(((*int16)(unsafe.Pointer(uintptr(varg.Val)))))
		}
		/*
			if varg.VT == (VT_BSTR|VT_BYREF) && varg.Val != 0 {
				*(params[n].(*string)) = UTF16PtrToString((*uint16)(unsafe.Pointer(uintptr(varg.Val))))
				println(*(params[n].(*string)))
			}
		*/
	}
	return
}

func GetUserDefaultLCID() (lcid uint32) {
	ret, _, _ := syscall.Syscall(
		uintptr(procGetUserDefaultLCID),
		0,
		0,
		0,
		0)
	lcid = uint32(ret)
	return
}

func IsEqualGUID(guid1 *GUID, guid2 *GUID) bool {
	return guid1.Data1 == guid2.Data1 &&
		guid1.Data2 == guid2.Data2 &&
		guid1.Data3 == guid2.Data3 &&
		guid1.Data4[0] == guid2.Data4[0] &&
		guid1.Data4[1] == guid2.Data4[1] &&
		guid1.Data4[2] == guid2.Data4[2] &&
		guid1.Data4[3] == guid2.Data4[3] &&
		guid1.Data4[4] == guid2.Data4[4] &&
		guid1.Data4[5] == guid2.Data4[5] &&
		guid1.Data4[6] == guid2.Data4[6] &&
		guid1.Data4[7] == guid2.Data4[7]
}

type Point struct {
	X int32
	Y int32
}

type Msg struct {
	Hwnd    uint32
	Message uint32
	Wparam  int32
	Lparam  int32
	Time    uint32
	Pt      Point
}

func GetMessage(msg *Msg, hwnd uint32, MsgFilterMin uint32, MsgFilterMax uint32) (ret int32, errno int) {
	r0, _, e1 := syscall.Syscall6(uintptr(procGetMessageW), 4, uintptr(unsafe.Pointer(msg)), uintptr(hwnd), uintptr(MsgFilterMin), uintptr(MsgFilterMax), 0, 0)
	ret = int32(r0)
	if ret == -1 {
		if e1 != 0 {
			errno = int(e1)
		} else {
			errno = int(syscall.EINVAL)
		}
	} else {
		errno = 0
	}
	return
}

func DispatchMessage(msg *Msg) (ret int32) {
	r0, _, _ := syscall.Syscall(uintptr(procDispatchMessageW), 1, uintptr(unsafe.Pointer(msg)), 0, 0)
	ret = int32(r0)
	return
}

const (
	TKIND_ENUM      = 1
	TKIND_RECORD    = 2
	TKIND_MODULE    = 3
	TKIND_INTERFACE = 4
	TKIND_DISPATCH  = 5
	TKIND_COCLASS   = 6
	TKIND_ALIAS     = 7
	TKIND_UNION     = 8
	TKIND_MAX       = 9
)

type TYPEDESC struct {
	Hreftype uint32
	VT       uint16
}

type IDLDESC struct {
	DwReserved uint32
	WIDLFlags  uint16
}

type TYPEATTR struct {
	Guid             GUID
	Lcid             uint32
	dwReserved       uint32
	MemidConstructor int32
	MemidDestructor  int32
	LpstrSchema      *uint16
	CbSizeInstance   uint32
	Typekind         int32
	CFuncs           uint16
	CVars            uint16
	CImplTypes       uint16
	CbSizeVft        uint16
	CbAlignment      uint16
	WTypeFlags       uint16
	WMajorVerNum     uint16
	WMinorVerNum     uint16
	TdescAlias       TYPEDESC
	IdldescType      IDLDESC
}
