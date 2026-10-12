//go:build windows
// +build windows

package ole

import (
	"fmt"
	"syscall"
	"testing"
	"unsafe"
)

// fakeInvoke is an IDispatch::Invoke behaving like a spec-abiding Automation
// server. Through every BYREF argument it writes a value that encodes the
// argument's rgvarg position, so a mix-up between rgvarg order (reversed) and
// parameter order shows up as a wrong value rather than a plausible one:
//
//	VT_BOOL|VT_BYREF  <- 0x0001 (what e.g. zkemkeeper writes; not VARIANT_TRUE)
//	VT_I4|VT_BYREF    <- 100 + rgvarg index
//	VT_BSTR|VT_BYREF  <- fresh BSTR "arg<rgvarg index>"
//
// It lets the out-parameter marshaling in invoke be tested without a
// registered COM server.
func fakeInvoke(this, dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr uintptr) uintptr {
	dp := (*DISPPARAMS)(unsafe.Pointer(pDispParams))
	args := (*[1 << 10]VARIANT)(unsafe.Pointer(dp.rgvarg))[:dp.cArgs:dp.cArgs]
	for i := range args {
		p := unsafe.Pointer(uintptr(args[i].Val))
		switch args[i].VT {
		case VT_BOOL | VT_BYREF:
			*(*int16)(p) = 1
		case VT_I4 | VT_BYREF:
			*(*int32)(p) = int32(100 + i)
		case VT_BSTR | VT_BYREF:
			*(**uint16)(p) = (*uint16)(unsafe.Pointer(SysAllocString(fmt.Sprintf("arg%d", i))))
		}
	}
	return S_OK
}

func newFakeDispatch() *IDispatch {
	vt := &IDispatchVtbl{Invoke: syscall.NewCallback(fakeInvoke)}
	obj := &struct{ vtbl *IDispatchVtbl }{vt}
	return (*IDispatch)(unsafe.Pointer(obj))
}

// The shape of zkemkeeper's SSR_GetUserInfo(machine, id, &name, &password,
// &privilege, &enabled): six parameters, out-parameters at positions 2..5.
// rgvarg index = 5 - parameter index.
func TestInvokeByRefOutParamsKeepTheirPosition(t *testing.T) {
	disp := newFakeDispatch()
	var name, password string
	var privilege int32
	if _, err := invoke(disp, 1, DISPATCH_METHOD, int32(1), int32(4711), &name, &password, &privilege, int32(0)); err != nil {
		t.Fatal(err)
	}
	if name != "arg3" {
		t.Errorf("name (params[2]) = %q, want %q", name, "arg3")
	}
	if password != "arg2" {
		t.Errorf("password (params[3]) = %q, want %q", password, "arg2")
	}
	if privilege != 101 {
		t.Errorf("privilege (params[4]) = %d, want 101", privilege)
	}
}
