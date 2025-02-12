//go:build 386 || arm || mips || mipsle
// +build 386 arm mips mipsle

package ole

type VARIANT struct {
	VT         VT     //  2
	wReserved1 uint16 //  4
	wReserved2 uint16 //  6
	wReserved3 uint16 //  8
	Val        int64  // 16
}
