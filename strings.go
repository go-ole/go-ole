//go:build windows

package ole

import (
	"fmt"
	"unicode/utf16"
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	procSysStringLen      = modoleaut32.NewProc("SysStringLen")
	procSysAllocString    = modoleaut32.NewProc("SysAllocString")
	procSysAllocStringLen = modoleaut32.NewProc("SysAllocStringLen")
	procSysFreeString     = modoleaut32.NewProc("SysFreeString")
)

// SysAllocString allocates memory for string and copies string into memory.
func SysAllocString(v string) (ss *uint16) {
	pss, _, _ := procSysAllocString.Call(uintptr(unsafe.Pointer(windows.StringToUTF16Ptr(v))))
	ss = (*uint16)(unsafe.Pointer(pss))
	return
}

// SysAllocStringLen copies up to length of given string returning pointer.
func SysAllocStringLen(v string) (ss *uint16) {
	utf16 := utf16.Encode([]rune(v + "\x00"))
	ptr := &utf16[0]

	pss, _, _ := procSysAllocStringLen.Call(uintptr(unsafe.Pointer(ptr)), uintptr(len(utf16)-1))
	ss = (*uint16)(unsafe.Pointer(pss))
	return
}

// SysFreeString frees string system memory. This must be called with SysAllocString.
func SysFreeString(v *uint16) (err error) {
	hr, _, _ := procSysFreeString.Call(uintptr(unsafe.Pointer(v)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// SysStringLen is the length of the system allocated string.
func SysStringLen(v *int16) uint32 {
	l, _, _ := procSysStringLen.Call(uintptr(unsafe.Pointer(v)))
	return uint32(l)
}

func GetErrorDescription(errno int) string {
	var flags uint32 = windows.FORMAT_MESSAGE_FROM_SYSTEM | windows.FORMAT_MESSAGE_ARGUMENT_ARRAY | windows.FORMAT_MESSAGE_IGNORE_INSERTS
	b := make([]uint16, 300)
	n, err := windows.FormatMessage(flags, 0, uint32(errno), 0, b, nil)
	if err != nil {
		return fmt.Sprintf("error %d (FormatMessage failed with: %v)", errno, err)
	}
	// trim terminating \r and \n
	for ; n > 0 && (b[n-1] == '\n' || b[n-1] == '\r'); n-- {
	}
	return string(utf16.Decode(b[:n]))
}
