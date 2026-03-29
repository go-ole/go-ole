//go:build windows

package ole

import (
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestITypeInfo(t *testing.T) {
	t.Run("GetTypeAttr", func(t *testing.T) {
		want := &TYPEATTR{CbSizeInstance: 42}
		virtualTable := &ITypeInfoVirtualTable{
			GetTypeAttr: syscall.NewCallback(func(this uintptr, tattr uintptr) uintptr {
				*(**TYPEATTR)(unsafe.Pointer(tattr)) = want
				return uintptr(windows.S_OK)
			}),
			ReleaseTypeAttr: syscall.NewCallback(func(this uintptr, tattr uintptr) uintptr {
				return 0
			}),
		}
		info := &ITypeInfo{VirtualTable: virtualTable}

		got, err := info.GetTypeAttr()
		if err != nil {
			t.Fatalf("GetTypeAttr failed: %v", err)
		}
		if got.CbSizeInstance != want.CbSizeInstance {
			t.Fatalf("GetTypeAttr() = %v, want %v", got, want)
		}
	})

	t.Run("GetDocumentation", func(t *testing.T) {
		wantName := "TestName"
		wantDoc := "TestDoc"
		virtualTable := &ITypeInfoVirtualTable{
			GetDocumentation: syscall.NewCallback(func(this uintptr, memid uintptr, name uintptr, doc uintptr, context uintptr, help uintptr) uintptr {
				if name != 0 {
					*(*uintptr)(unsafe.Pointer(name)) = uintptr(unsafe.Pointer(SysAllocString(wantName)))
				}
				if doc != 0 {
					*(*uintptr)(unsafe.Pointer(doc)) = uintptr(unsafe.Pointer(SysAllocString(wantDoc)))
				}
				return uintptr(windows.S_OK)
			}),
		}
		info := &ITypeInfo{VirtualTable: virtualTable}

		name, doc, _, _, err := info.GetDocumentation(0)
		if err != nil {
			t.Fatalf("GetDocumentation failed: %v", err)
		}
		if name != wantName {
			t.Fatalf("GetDocumentation() name = %q, want %q", name, wantName)
		}
		if doc != wantDoc {
			t.Fatalf("GetDocumentation() doc = %q, want %q", doc, wantDoc)
		}
	})

	t.Run("GetFuncDesc", func(t *testing.T) {
		want := &FUNCDESC{Memid: 123}
		virtualTable := &ITypeInfoVirtualTable{
			GetFuncDesc: syscall.NewCallback(func(this uintptr, index uintptr, fdesc uintptr) uintptr {
				*(**FUNCDESC)(unsafe.Pointer(fdesc)) = want
				return uintptr(windows.S_OK)
			}),
			ReleaseFuncDesc: syscall.NewCallback(func(this uintptr, fdesc uintptr) uintptr {
				return 0
			}),
		}
		info := &ITypeInfo{VirtualTable: virtualTable}

		got, err := info.GetFuncDesc(0)
		if err != nil {
			t.Fatalf("GetFuncDesc failed: %v", err)
		}
		if got.Memid != want.Memid {
			t.Fatalf("GetFuncDesc() = %v, want %v", got, want)
		}
	})

	t.Run("GetVarDesc", func(t *testing.T) {
		want := &VARDESC{Memid: 456}
		virtualTable := &ITypeInfoVirtualTable{
			GetVarDesc: syscall.NewCallback(func(this uintptr, index uintptr, vdesc uintptr) uintptr {
				*(**VARDESC)(unsafe.Pointer(vdesc)) = want
				return uintptr(windows.S_OK)
			}),
			ReleaseVarDesc: syscall.NewCallback(func(this uintptr, vdesc uintptr) uintptr {
				return 0
			}),
		}
		info := &ITypeInfo{VirtualTable: virtualTable}

		got, err := info.GetVarDesc(0)
		if err != nil {
			t.Fatalf("GetVarDesc failed: %v", err)
		}
		if got.Memid != want.Memid {
			t.Fatalf("GetVarDesc() = %v, want %v", got, want)
		}
	})

	t.Run("GetIDsOfNames", func(t *testing.T) {
		virtualTable := &ITypeInfoVirtualTable{
			GetIDsOfNames: syscall.NewCallback(func(this uintptr, names uintptr, count uintptr, memids uintptr) uintptr {
				ids := unsafe.Slice((*MEMBERID)(unsafe.Pointer(memids)), int(count))
				ids[0] = 10
				ids[1] = 20
				return uintptr(windows.S_OK)
			}),
		}
		info := &ITypeInfo{VirtualTable: virtualTable}

		names := []*uint16{SysAllocString("Alpha"), SysAllocString("Beta")}
		got, err := info.GetIDsOfNames(names, 2)
		if err != nil {
			t.Fatalf("GetIDsOfNames failed: %v", err)
		}
		if len(got) != 2 || got[0] != 10 || got[1] != 20 {
			t.Fatalf("GetIDsOfNames() = %v, want [10 20]", got)
		}
	})
}
