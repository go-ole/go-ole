//go:build windows

package ole

import (
	"errors"
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestIDispatchHasTypeInfo(t *testing.T) {
	t.Run("true", func(t *testing.T) {
		virtualTable := &IDispatchVirtualTable{
			GetTypeInfoCount: syscall.NewCallback(func(this uintptr, count uintptr) uintptr {
				*(*uint)(unsafe.Pointer(count)) = 1
				return uintptr(windows.S_OK)
			}),
		}
		dispatch := &IDispatch{VirtualTable: virtualTable}

		if !dispatch.HasTypeInfo() {
			t.Fatal("HasTypeInfo() = false, want true")
		}
	})

	t.Run("false on not impl", func(t *testing.T) {
		virtualTable := &IDispatchVirtualTable{
			GetTypeInfoCount: syscall.NewCallback(func(this uintptr, count uintptr) uintptr {
				return uintptr(windows.E_NOTIMPL)
			}),
		}
		dispatch := &IDispatch{VirtualTable: virtualTable}

		if dispatch.HasTypeInfo() {
			t.Fatal("HasTypeInfo() = true, want false")
		}
	})

	t.Run("false on zero count", func(t *testing.T) {
		virtualTable := &IDispatchVirtualTable{
			GetTypeInfoCount: syscall.NewCallback(func(this uintptr, count uintptr) uintptr {
				*(*uint)(unsafe.Pointer(count)) = 0
				return uintptr(windows.S_OK)
			}),
		}
		dispatch := &IDispatch{VirtualTable: virtualTable}

		if dispatch.HasTypeInfo() {
			t.Fatal("HasTypeInfo() = true, want false")
		}
	})
}

func TestIDispatchGetTypeInfo(t *testing.T) {
	t.Run("success", func(t *testing.T) {
		want := &ITypeInfo{}
		virtualTable := &IDispatchVirtualTable{
			GetTypeInfo: syscall.NewCallback(func(this uintptr, index uintptr, lcid uintptr, info uintptr) uintptr {
				*(**ITypeInfo)(unsafe.Pointer(info)) = want
				return uintptr(windows.S_OK)
			}),
		}
		dispatch := &IDispatch{VirtualTable: virtualTable}

		got := dispatch.GetTypeInfo()
		if got != want {
			t.Fatalf("GetTypeInfo() = %p, want %p", got, want)
		}
	})

	t.Run("bad index returns nil", func(t *testing.T) {
		virtualTable := &IDispatchVirtualTable{
			GetTypeInfo: syscall.NewCallback(func(this uintptr, index uintptr, lcid uintptr, info uintptr) uintptr {
				return uintptr(windows.DISP_E_BADINDEX)
			}),
		}
		dispatch := &IDispatch{VirtualTable: virtualTable}

		if got := dispatch.GetTypeInfo(); got != nil {
			t.Fatalf("GetTypeInfo() = %p, want nil", got)
		}
	})
}

func TestIDispatchGetIDsOfNames(t *testing.T) {
	virtualTable := &IDispatchVirtualTable{
		GetIDsOfNames: syscall.NewCallback(func(this uintptr, iid uintptr, names uintptr, count uintptr, lcid uintptr, ids uintptr) uintptr {
			namePointers := unsafe.Slice((**uint16)(unsafe.Pointer(names)), int(count))
			displayIDs := unsafe.Slice((*int32)(unsafe.Pointer(ids)), int(count))

			for index, namePointer := range namePointers {
				switch windows.UTF16PtrToString(namePointer) {
				case "Alpha":
					displayIDs[index] = 10
				case "Beta":
					displayIDs[index] = 20
				default:
					displayIDs[index] = DISPID_UNKNOWN
				}
			}

			return uintptr(windows.S_OK)
		}),
	}
	dispatch := &IDispatch{VirtualTable: virtualTable}

	got, err := dispatch.GetIDsOfNames([]string{"Alpha", "Beta"})
	if err != nil {
		t.Fatalf("GetIDsOfNames failed: %v", err)
	}
	if got["Alpha"] != 10 || got["Beta"] != 20 {
		t.Fatalf("GetIDsOfNames() = %#v, want Alpha=10 Beta=20", got)
	}
}

func TestIDispatchGetSingleIDOfName(t *testing.T) {
	virtualTable := &IDispatchVirtualTable{
		GetIDsOfNames: syscall.NewCallback(func(this uintptr, iid uintptr, names uintptr, count uintptr, lcid uintptr, ids uintptr) uintptr {
			*(*int32)(unsafe.Pointer(ids)) = 42
			return uintptr(windows.S_OK)
		}),
	}
	dispatch := &IDispatch{VirtualTable: virtualTable}

	got, err := dispatch.GetSingleIDOfName("Answer")
	if err != nil {
		t.Fatalf("GetSingleIDOfName failed: %v", err)
	}
	if got != 42 {
		t.Fatalf("GetSingleIDOfName() = %d, want 42", got)
	}
}

func TestIDispatchInvokeHelpers(t *testing.T) {
	var gotDispatch int16
	var gotDisplayID int32
	var gotArgCount uint32
	var gotNamedArgCount uint32
	var gotNamedArg int32

	virtualTable := &IDispatchVirtualTable{
		GetIDsOfNames: syscall.NewCallback(func(this uintptr, iid uintptr, names uintptr, count uintptr, lcid uintptr, ids uintptr) uintptr {
			*(*int32)(unsafe.Pointer(ids)) = 99
			return uintptr(windows.S_OK)
		}),
		Invoke: syscall.NewCallback(func(this uintptr, dispid uintptr, iid uintptr, lcid uintptr, dispatch uintptr, params uintptr, result uintptr, excepInfo uintptr, argErr uintptr) uintptr {
			gotDisplayID = int32(dispid)
			gotDispatch = int16(dispatch)

			dispParams := (*DISPPARAMS)(unsafe.Pointer(params))
			gotArgCount = dispParams.cArgs
			gotNamedArgCount = dispParams.cNamedArgs
			if dispParams.rgdispidNamedArgs != 0 {
				gotNamedArg = *(*int32)(unsafe.Pointer(dispParams.rgdispidNamedArgs))
			}

			return uintptr(windows.S_OK)
		}),
	}
	dispatch := &IDispatch{VirtualTable: virtualTable}
	param := Int32ToVariant(int32(7))
	defer param.Clear()

	t.Run("CallMethod", func(t *testing.T) {
		_, err := dispatch.CallMethod("Method", param)
		if err != nil {
			t.Fatalf("CallMethod failed: %v", err)
		}
		if gotDisplayID != 99 || gotDispatch != DISPATCH_METHOD || gotArgCount != 1 || gotNamedArgCount != 0 {
			t.Fatalf("CallMethod captured dispid=%d dispatch=%d cArgs=%d cNamedArgs=%d", gotDisplayID, gotDispatch, gotArgCount, gotNamedArgCount)
		}
	})

	t.Run("GetProperty", func(t *testing.T) {
		_, err := dispatch.GetProperty("Property", param)
		if err != nil {
			t.Fatalf("GetProperty failed: %v", err)
		}
		if gotDispatch != DISPATCH_PROPERTYGET || gotArgCount != 1 || gotNamedArgCount != 0 {
			t.Fatalf("GetProperty captured dispatch=%d cArgs=%d cNamedArgs=%d", gotDispatch, gotArgCount, gotNamedArgCount)
		}
	})

	t.Run("PutProperty", func(t *testing.T) {
		_, err := dispatch.PutProperty("Property", param)
		if err != nil {
			t.Fatalf("PutProperty failed: %v", err)
		}
		if gotDispatch != DISPATCH_PROPERTYPUT || gotArgCount != 1 || gotNamedArgCount != 1 || gotNamedArg != DISPID_PROPERTYPUT {
			t.Fatalf("PutProperty captured dispatch=%d cArgs=%d cNamedArgs=%d namedArg=%d", gotDispatch, gotArgCount, gotNamedArgCount, gotNamedArg)
		}
	})

	t.Run("PutPropertyRef", func(t *testing.T) {
		_, err := dispatch.PutPropertyRef("Property", param)
		if err != nil {
			t.Fatalf("PutPropertyRef failed: %v", err)
		}
		if gotDispatch != DISPATCH_PROPERTYPUTREF || gotArgCount != 1 || gotNamedArgCount != 1 || gotNamedArg != DISPID_PROPERTYPUT {
			t.Fatalf("PutPropertyRef captured dispatch=%d cArgs=%d cNamedArgs=%d namedArg=%d", gotDispatch, gotArgCount, gotNamedArgCount, gotNamedArg)
		}
	})
}

func TestInvokeOnIDispatchReturnsJoinedError(t *testing.T) {
	testErr := uintptr(windows.E_POINTER)
	virtualTable := &IDispatchVirtualTable{
		Invoke: syscall.NewCallback(func(this uintptr, dispid uintptr, iid uintptr, lcid uintptr, dispatch uintptr, params uintptr, result uintptr, excepInfo uintptr, argErr uintptr) uintptr {
			info := (*EXCEPINFO)(unsafe.Pointer(excepInfo))
			info.bstrDescription = SysAllocString("dispatch failed")
			return testErr
		}),
	}
	dispatch := &IDispatch{VirtualTable: virtualTable}

	_, err := InvokeOnIDispatch(dispatch, 55, DISPATCH_METHOD)
	if !errors.Is(err, windows.Errno(testErr)) {
		t.Fatalf("InvokeOnIDispatch error = %v, want wrapped %v", err, windows.Errno(testErr))
	}
	if err == nil || err.Error() != "The pointer is invalid.\ndispatch failed" {
		t.Fatalf("InvokeOnIDispatch error text = %q", err)
	}
}

func TestMakeDisplayParams(t *testing.T) {
	param := Int32ToVariant(int32(1))
	defer param.Clear()

	methodParams := MakeDisplayParams(DISPATCH_METHOD, param)
	if methodParams.cArgs != 1 || methodParams.cNamedArgs != 0 {
		t.Fatalf("MakeDisplayParams(method) = %#v", methodParams)
	}

	putParams := MakeDisplayParams(DISPATCH_PROPERTYPUT, param)
	if putParams.cArgs != 1 || putParams.cNamedArgs != 1 {
		t.Fatalf("MakeDisplayParams(property put) = %#v", putParams)
	}
	if got := *(*int32)(unsafe.Pointer(putParams.rgdispidNamedArgs)); got != DISPID_PROPERTYPUT {
		t.Fatalf("MakeDisplayParams(property put) named arg = %d, want %d", got, DISPID_PROPERTYPUT)
	}

	putRefParams := MakeDisplayParams(DISPATCH_PROPERTYPUTREF, param)
	if putRefParams.cArgs != 1 || putRefParams.cNamedArgs != 1 {
		t.Fatalf("MakeDisplayParams(property putref) = %#v", putRefParams)
	}
	if got := *(*int32)(unsafe.Pointer(putRefParams.rgdispidNamedArgs)); got != DISPID_PROPERTYPUT {
		t.Fatalf("MakeDisplayParams(property putref) named arg = %d, want %d", got, DISPID_PROPERTYPUT)
	}
}

func TestQueryIDispatchFromIUnknownNil(t *testing.T) {
	got, err := QueryIDispatchFromIUnknown(nil)
	if got != nil {
		t.Fatalf("QueryIDispatchFromIUnknown(nil) = %p, want nil", got)
	}
	if !errors.Is(err, ComInterfaceIsNilPointer) {
		t.Fatalf("QueryIDispatchFromIUnknown(nil) error = %v, want %v", err, ComInterfaceIsNilPointer)
	}
}
