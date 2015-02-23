package ole

import (
	"fmt"
	"syscall"
	"unicode/utf16"
	"unsafe"
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

func (v *VARIANT) ToArray() *SafeArrayConversion {
	var safeArray *SafeArray = (*SafeArray)(unsafe.Pointer(uintptr(v.Val)))
	return &SafeArrayConversion{safeArray}
}

func (v *VARIANT) ToString() string {
	return UTF16PtrToString(*(**uint16)(unsafe.Pointer(&v.Val)))
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
