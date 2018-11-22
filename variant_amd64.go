// +build windows,amd64

package ole

import (
	"errors"
	"syscall"
	"time"
	"unsafe"
)

type VARIANT struct {
	VT         VT      //  2
	wReserved1 uint16  //  4
	wReserved2 uint16  //  6
	wReserved3 uint16  //  8
	Val        int64   // 16
	_          [8]byte // 24
}

// GetVariantDate converts COM Variant Time value to Go time.Time.
func GetVariantDate(value uint64) (time.Time, error) {
	var st syscall.Systemtime
	r, _, _ := procVariantTimeToSystemTime.Call(uintptr(value), uintptr(unsafe.Pointer(&st)))
	if r != 0 {
		return time.Date(int(st.Year), time.Month(st.Month), int(st.Day), int(st.Hour), int(st.Minute), int(st.Second), int(st.Milliseconds/1000), time.UTC), nil
	}
	return time.Now(), errors.New("Could not convert to time, passing current time.")
}
