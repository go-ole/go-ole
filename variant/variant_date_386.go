//go:build windows && 386
// +build windows,386

package variant

import (
	"errors"
	"math"
	"syscall"
	"time"
	"unsafe"
)

const ONETHOUSANDMILLISECONDS = 0.0000115740740740

// GetVariantDate converts COM Variant Time value to Go time.Time.
func GetVariantDate(value uint64) (time.Time, error) {
	halfSecond := ONETHOUSANDMILLISECONDS / 2.0
	dVariantTime := math.Float64frombits(value)
	var st syscall.Systemtime
	adjustedVariantTime := dVariantTime - halfSecond
	uAdjustedVariantTime := math.Float64bits(adjustedVariantTime)
	v1 := uint32(uAdjustedVariantTime)
	v2 := uint32(uAdjustedVariantTime >> 32)
	r, _, _ := procVariantTimeToSystemTime.Call(uintptr(v1), uintptr(v2), uintptr(unsafe.Pointer(&st)))
	if r != 0 {
		fraction := dVariantTime - float64(int(dVariantTime))
		hours := (fraction - float64(int(fraction))) * 24
		minutes := (hours - float64(int(hours))) * 60
		seconds := (minutes - float64(int(minutes))) * 60
		milliseconds := (seconds - float64(int(seconds))) * 1000
		milliseconds = milliseconds + 0.5
		if milliseconds < 1.0 || milliseconds > 999.0 {
			var st2 syscall.Systemtime
			v1 = uint32(value)
			v2 = uint32(value >> 32)
			r2, _, _ := procVariantTimeToSystemTime.Call(uintptr(v1), uintptr(v2), uintptr(unsafe.Pointer(&st2)))
			if r2 != 0 {
				return time.Date(int(st2.Year), time.Month(st2.Month), int(st2.Day), int(st2.Hour), int(st2.Minute), int(st2.Second), 0, time.UTC), nil
			} else {
				return time.Now(), errors.New("Could not convert to time, passing current time.")
			}
		}
		return time.Date(int(st.Year), time.Month(st.Month), int(st.Day), int(st.Hour), int(st.Minute), int(st.Second), int(int16(milliseconds))*1e6, time.UTC), nil
	}
	return time.Now(), errors.New("Could not convert to time, passing current time.")
}
