package ole

import (
	"reflect"
	"syscall"
	"unicode/utf8"
	"unsafe"
)

var (
	procRoInitialize, _              = modcombase.FindProc("RoInitialize")
	procRoActivateInstance, _        = modcombase.FindProc("RoActivateInstance")
	procRoGetActivationFactory, _    = modcombase.FindProc("RoGetActivationFactory")
	procWindowsCreateString, _       = modcombase.FindProc("WindowsCreateString")
	procWindowsDeleteString, _       = modcombase.FindProc("WindowsDeleteString")
	procWindowsGetStringRawBuffer, _ = modcombase.FindProc("WindowsGetStringRawBuffer")
)

type HString uintptr

func RoInitialize(thread_type uint32) (err error) {
	hr, _, _ := procRoInitialize.Call(uintptr(thread_type))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func RoActivateInstance(clsid string) (ins *IInspectable, err error) {
	hClsid, err := NewHString(clsid)
	if err != nil {
		return nil, err
	}
	defer DeleteHString(hClsid)

	hr, _, _ := procRoActivateInstance.Call(
		uintptr(unsafe.Pointer(hClsid)),
		uintptr(unsafe.Pointer(&ins)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func RoGetActivationFactory(clsid string, iid *GUID) (ins *IInspectable, err error) {
	hClsid, err := NewHString(clsid)
	if err != nil {
		return nil, err
	}
	defer DeleteHString(hClsid)

	hr, _, _ := procRoGetActivationFactory.Call(
		uintptr(unsafe.Pointer(hClsid)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&ins)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func NewHString(s string) (hstring HString, err error) {
	u16 := syscall.StringToUTF16Ptr(s)
	len := uint32(utf8.RuneCountInString(s))
	hr, _, _ := procWindowsCreateString.Call(
		uintptr(unsafe.Pointer(u16)),
		uintptr(len),
		uintptr(unsafe.Pointer(&hstring)))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func DeleteHString(hstring HString) (err error) {
	hr, _, _ := procWindowsDeleteString.Call(uintptr(hstring))
	if hr != 0 {
		err = NewError(hr)
	}
	return
}

func HStringToString(hString HString) (s string) {
	var u16buf uintptr
	var u16len uint32
	u16buf, _, _ = procWindowsGetStringRawBuffer.Call(
		uintptr(unsafe.Pointer(hString)),
		uintptr(unsafe.Pointer(&u16len)))

	u16hdr := *(*reflect.SliceHeader)(unsafe.Pointer(&u16buf))
	u16hdr.Len = int(u16len)
	u16hdr.Cap = int(u16len)
	u16 := *(*[]uint16)(unsafe.Pointer(&u16hdr))
	s = syscall.UTF16ToString(u16)
	return
}
