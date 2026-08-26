//go:build windows
// +build windows

package ole

import (
	"syscall"
	"testing"
	"unsafe"
)

// fakeInvoke is an IDispatch::Invoke behaving like a spec-abiding Automation
// server: it stores VARIANT_TRUE through every VT_BOOL|VT_BYREF argument and a
// freshly allocated BSTR through every VT_BSTR|VT_BYREF one. It lets the
// out-parameter marshaling in invoke be tested without a registered COM server.
func fakeInvoke(this, dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr uintptr) uintptr {
	dp := (*DISPPARAMS)(unsafe.Pointer(pDispParams))
	args := (*[1 << 10]VARIANT)(unsafe.Pointer(dp.rgvarg))[:dp.cArgs:dp.cArgs]
	for i := range args {
		switch args[i].VT {
		case VT_BOOL | VT_BYREF:
			*(*int16)(unsafe.Pointer(uintptr(args[i].Val))) = -1
		case VT_BSTR | VT_BYREF:
			*(**uint16)(unsafe.Pointer(uintptr(args[i].Val))) = (*uint16)(unsafe.Pointer(SysAllocString("out")))
		}
	}
	return S_OK
}

func newFakeDispatch() *IDispatch {
	vt := &IDispatchVtbl{Invoke: syscall.NewCallback(fakeInvoke)}
	obj := &struct{ vtbl *IDispatchVtbl }{vt}
	return (*IDispatch)(unsafe.Pointer(obj))
}

// rgvarg holds the parameters in reverse order. Out-parameters are placed so
// that their params index and rgvarg index differ and are not mirror images of
// another out-parameter of the same kind: an index mix-up between the two
// orders then reads an untouched cell instead of another parameter's.
func TestInvokeByRefOutParamsKeepTheirPosition(t *testing.T) {
	disp := newFakeDispatch()
	var enabled bool
	var name string
	if _, err := invoke(disp, 1, DISPATCH_METHOD, &enabled, int32(7), &name, "in"); err != nil {
		t.Fatal(err)
	}
	if !enabled {
		t.Errorf("*bool out-parameter at params[0] read back false, want true")
	}
	if name != "out" {
		t.Errorf("*string out-parameter at params[2] read back %q, want %q", name, "out")
	}
}

func TestInvokeByRefBoolInOut(t *testing.T) {
	disp := newFakeDispatch()
	in := true
	if _, err := invoke(disp, 1, DISPATCH_METHOD, int32(1), &in); err != nil {
		t.Fatal(err)
	}
	if !in {
		t.Errorf("*bool in/out parameter lost its value")
	}
}
