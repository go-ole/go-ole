//go:build windows

package ole

import (
	"errors"
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestIRecordInfoAddressMethods(t *testing.T) {
	virtualTable := &IRecordInfoVirtualTable{
		QueryInterface:   1,
		AddRef:           2,
		Release:          3,
		RecordInit:       4,
		RecordClear:      5,
		RecordCopy:       6,
		GetGuid:          7,
		GetName:          8,
		GetSize:          9,
		GetTypeInfo:      10,
		GetField:         11,
		GetFieldNoCopy:   12,
		PutField:         13,
		PutFieldNoCopy:   14,
		GetFieldNames:    15,
		IsMatchingType:   16,
		RecordCreate:     17,
		RecordCreateCopy: 18,
		RecordDestroy:    19,
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}

	if recordInfo.QueryInterfaceAddress() != 1 || recordInfo.AddRefAddress() != 2 || recordInfo.ReleaseAddress() != 3 {
		t.Fatal("IUnknown address accessors did not mirror the vtable")
	}
	if recordInfo.RecordInitAddress() != 4 || recordInfo.RecordClearAddress() != 5 || recordInfo.RecordCopyAddress() != 6 {
		t.Fatal("record lifecycle address accessors did not mirror the vtable")
	}
	if recordInfo.GetGuidAddress() != 7 || recordInfo.GetNameAddress() != 8 || recordInfo.GetSizeAddress() != 9 || recordInfo.GetTypeInfoAddress() != 10 {
		t.Fatal("metadata address accessors did not mirror the vtable")
	}
	if recordInfo.GetFieldAddress() != 11 || recordInfo.GetFieldNoCopyAddress() != 12 || recordInfo.PutFieldAddress() != 13 || recordInfo.PutFieldNoCopyAddress() != 14 {
		t.Fatal("field access address accessors did not mirror the vtable")
	}
	if recordInfo.GetFieldNamesAddress() != 15 || recordInfo.IsMatchingTypeAddress() != 16 || recordInfo.RecordCreateAddress() != 17 || recordInfo.RecordCreateCopyAddress() != 18 || recordInfo.RecordDestroyAddress() != 19 {
		t.Fatal("remaining address accessors did not mirror the vtable")
	}
}

func TestIRecordInfoRecordInitUsesCallerAllocatedBuffer(t *testing.T) {
	var gotRecord uintptr
	virtualTable := &IRecordInfoVirtualTable{
		RecordInit: syscall.NewCallback(func(this uintptr, newRecord uintptr) uintptr {
			gotRecord = newRecord
			*(*byte)(unsafe.Pointer(newRecord)) = 0x7f
			return uintptr(windows.S_OK)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}
	buffer := [8]byte{}

	if err := recordInfo.RecordInit(uintptr(unsafe.Pointer(&buffer[0]))); err != nil {
		t.Fatalf("RecordInit failed: %v", err)
	}
	if gotRecord != uintptr(unsafe.Pointer(&buffer[0])) {
		t.Fatalf("RecordInit received %#x, want %#x", gotRecord, uintptr(unsafe.Pointer(&buffer[0])))
	}
	if buffer[0] != 0x7f {
		t.Fatalf("RecordInit did not operate on caller buffer")
	}
}

func TestIRecordInfoGetGuid(t *testing.T) {
	want := IID_IProvideClassInfo
	virtualTable := &IRecordInfoVirtualTable{
		GetGuid: syscall.NewCallback(func(this uintptr, guid uintptr) uintptr {
			*(*windows.GUID)(unsafe.Pointer(guid)) = want
			return uintptr(windows.S_OK)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}

	got, err := recordInfo.GetGuid()
	if err != nil {
		t.Fatalf("GetGuid failed: %v", err)
	}
	if got != want {
		t.Fatalf("GetGuid() = %v, want %v", got, want)
	}
}

func TestIRecordInfoGetGuidErrors(t *testing.T) {
	t.Run("invalid state", func(t *testing.T) {
		virtualTable := &IRecordInfoVirtualTable{
			GetGuid: syscall.NewCallback(func(this uintptr, guid uintptr) uintptr {
				return uintptr(windows.TYPE_E_INVALIDSTATE)
			}),
		}
		recordInfo := &IRecordInfo{VirtualTable: virtualTable}

		if _, err := recordInfo.GetGuid(); !errors.Is(err, InvalidArgumentComError) {
			t.Fatalf("GetGuid error = %v, want %v", err, InvalidArgumentComError)
		}
	})

	t.Run("invalid arg", func(t *testing.T) {
		virtualTable := &IRecordInfoVirtualTable{
			GetGuid: syscall.NewCallback(func(this uintptr, guid uintptr) uintptr {
				return uintptr(windows.E_INVALIDARG)
			}),
		}
		recordInfo := &IRecordInfo{VirtualTable: virtualTable}

		if _, err := recordInfo.GetGuid(); !errors.Is(err, InvalidArgumentComError) {
			t.Fatalf("GetGuid error = %v, want %v", err, InvalidArgumentComError)
		}
	})
}

func TestIRecordInfoGetSize(t *testing.T) {
	virtualTable := &IRecordInfoVirtualTable{
		GetSize: syscall.NewCallback(func(this uintptr, size uintptr) uintptr {
			*(*uint32)(unsafe.Pointer(size)) = 64
			return uintptr(windows.S_OK)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}

	got, err := recordInfo.GetSize()
	if err != nil {
		t.Fatalf("GetSize failed: %v", err)
	}
	if got != 64 {
		t.Fatalf("GetSize() = %d, want 64", got)
	}
}

func TestIRecordInfoGetSizeInvalidArg(t *testing.T) {
	virtualTable := &IRecordInfoVirtualTable{
		GetSize: syscall.NewCallback(func(this uintptr, size uintptr) uintptr {
			return uintptr(windows.E_INVALIDARG)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}

	if _, err := recordInfo.GetSize(); !errors.Is(err, InvalidArgumentComError) {
		t.Fatalf("GetSize error = %v, want %v", err, InvalidArgumentComError)
	}
}

func TestIRecordInfoRecordInitInvalidArg(t *testing.T) {
	virtualTable := &IRecordInfoVirtualTable{
		RecordInit: syscall.NewCallback(func(this uintptr, newRecord uintptr) uintptr {
			return uintptr(windows.E_INVALIDARG)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}

	if err := recordInfo.RecordInit(0); !errors.Is(err, InvalidArgumentComError) {
		t.Fatalf("RecordInit error = %v, want %v", err, InvalidArgumentComError)
	}
}

func TestIRecordInfoRecordClear(t *testing.T) {
	var gotExisting uintptr
	virtualTable := &IRecordInfoVirtualTable{
		RecordClear: syscall.NewCallback(func(this uintptr, existing uintptr) uintptr {
			gotExisting = existing
			return uintptr(windows.S_OK)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}
	value := [4]byte{1, 2, 3, 4}

	if err := recordInfo.RecordClear(uintptr(unsafe.Pointer(&value[0]))); err != nil {
		t.Fatalf("RecordClear failed: %v", err)
	}
	if gotExisting != uintptr(unsafe.Pointer(&value[0])) {
		t.Fatalf("RecordClear received %#x, want %#x", gotExisting, uintptr(unsafe.Pointer(&value[0])))
	}
}

func TestIRecordInfoRecordClearInvalidArg(t *testing.T) {
	virtualTable := &IRecordInfoVirtualTable{
		RecordClear: syscall.NewCallback(func(this uintptr, existing uintptr) uintptr {
			return uintptr(windows.E_INVALIDARG)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}

	if err := recordInfo.RecordClear(0); !errors.Is(err, InvalidArgumentComError) {
		t.Fatalf("RecordClear error = %v, want %v", err, InvalidArgumentComError)
	}
}

func TestIRecordInfoRecordCopyUsesDestinationBuffer(t *testing.T) {
	var gotExisting uintptr
	var gotNew uintptr
	virtualTable := &IRecordInfoVirtualTable{
		RecordCopy: syscall.NewCallback(func(this uintptr, existing uintptr, newRecord uintptr) uintptr {
			gotExisting = existing
			gotNew = newRecord
			*(*uint32)(unsafe.Pointer(newRecord)) = *(*uint32)(unsafe.Pointer(existing))
			return uintptr(windows.S_OK)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}
	existing := uint32(1234)
	var copied uint32

	if err := recordInfo.RecordCopy(uintptr(unsafe.Pointer(&existing)), uintptr(unsafe.Pointer(&copied))); err != nil {
		t.Fatalf("RecordCopy failed: %v", err)
	}
	if gotExisting != uintptr(unsafe.Pointer(&existing)) || gotNew != uintptr(unsafe.Pointer(&copied)) {
		t.Fatalf("RecordCopy received existing=%#x new=%#x", gotExisting, gotNew)
	}
	if copied != existing {
		t.Fatalf("RecordCopy destination = %d, want %d", copied, existing)
	}
}

func TestIRecordInfoRecordCopyInvalidArg(t *testing.T) {
	virtualTable := &IRecordInfoVirtualTable{
		RecordCopy: syscall.NewCallback(func(this uintptr, existing uintptr, newRecord uintptr) uintptr {
			return uintptr(windows.E_INVALIDARG)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}

	if err := recordInfo.RecordCopy(0, 0); !errors.Is(err, InvalidArgumentComError) {
		t.Fatalf("RecordCopy error = %v, want %v", err, InvalidArgumentComError)
	}
}

func TestIRecordInfoRecordCreate(t *testing.T) {
	want := uintptr(0x12345678)
	virtualTable := &IRecordInfoVirtualTable{
		RecordCreate: syscall.NewCallback(func(this uintptr) uintptr {
			return want
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}

	got, err := recordInfo.RecordCreate()
	if err != nil {
		t.Fatalf("RecordCreate failed: %v", err)
	}
	if got != want {
		t.Fatalf("RecordCreate() = %#x, want %#x", got, want)
	}
}

func TestIRecordInfoRecordCreateCopy(t *testing.T) {
	t.Run("success", func(t *testing.T) {
		existing := uint32(99)
		want := uintptr(0x87654321)
		var gotExisting uintptr

		virtualTable := &IRecordInfoVirtualTable{
			RecordCreateCopy: syscall.NewCallback(func(this uintptr, existingRecord uintptr, out uintptr) uintptr {
				gotExisting = existingRecord
				*(*uintptr)(unsafe.Pointer(out)) = want
				return uintptr(windows.S_OK)
			}),
		}
		recordInfo := &IRecordInfo{VirtualTable: virtualTable}

		got, err := recordInfo.RecordCreateCopy(uintptr(unsafe.Pointer(&existing)))
		if err != nil {
			t.Fatalf("RecordCreateCopy failed: %v", err)
		}
		if gotExisting != uintptr(unsafe.Pointer(&existing)) {
			t.Fatalf("RecordCreateCopy existing = %#x, want %#x", gotExisting, uintptr(unsafe.Pointer(&existing)))
		}
		if got != want {
			t.Fatalf("RecordCreateCopy() = %#x, want %#x", got, want)
		}
	})

	t.Run("out of memory", func(t *testing.T) {
		virtualTable := &IRecordInfoVirtualTable{
			RecordCreateCopy: syscall.NewCallback(func(this uintptr, existingRecord uintptr, out uintptr) uintptr {
				return uintptr(windows.E_OUTOFMEMORY)
			}),
		}
		recordInfo := &IRecordInfo{VirtualTable: virtualTable}

		if _, err := recordInfo.RecordCreateCopy(0); !errors.Is(err, OutOfMemoryComError) {
			t.Fatalf("RecordCreateCopy error = %v, want %v", err, OutOfMemoryComError)
		}
	})

	t.Run("invalid arg", func(t *testing.T) {
		virtualTable := &IRecordInfoVirtualTable{
			RecordCreateCopy: syscall.NewCallback(func(this uintptr, existingRecord uintptr, out uintptr) uintptr {
				return uintptr(windows.E_INVALIDARG)
			}),
		}
		recordInfo := &IRecordInfo{VirtualTable: virtualTable}

		if _, err := recordInfo.RecordCreateCopy(0); !errors.Is(err, InvalidArgumentComError) {
			t.Fatalf("RecordCreateCopy error = %v, want %v", err, InvalidArgumentComError)
		}
	})
}

func TestIRecordInfoRecordDestroy(t *testing.T) {
	var gotExisting uintptr
	virtualTable := &IRecordInfoVirtualTable{
		RecordDestroy: syscall.NewCallback(func(this uintptr, existing uintptr) uintptr {
			gotExisting = existing
			return uintptr(windows.S_OK)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}
	value := [4]byte{5, 6, 7, 8}

	if err := recordInfo.RecordDestroy(uintptr(unsafe.Pointer(&value[0]))); err != nil {
		t.Fatalf("RecordDestroy failed: %v", err)
	}
	if gotExisting != uintptr(unsafe.Pointer(&value[0])) {
		t.Fatalf("RecordDestroy received %#x, want %#x", gotExisting, uintptr(unsafe.Pointer(&value[0])))
	}
}

func TestIRecordInfoRecordDestroyInvalidArg(t *testing.T) {
	virtualTable := &IRecordInfoVirtualTable{
		RecordDestroy: syscall.NewCallback(func(this uintptr, existing uintptr) uintptr {
			return uintptr(windows.E_INVALIDARG)
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}

	if err := recordInfo.RecordDestroy(0); !errors.Is(err, InvalidArgumentComError) {
		t.Fatalf("RecordDestroy error = %v, want %v", err, InvalidArgumentComError)
	}
}

func TestIRecordInfoIsMatchingType(t *testing.T) {
	var gotArg uintptr
	other := &IRecordInfo{}
	virtualTable := &IRecordInfoVirtualTable{
		IsMatchingType: syscall.NewCallback(func(this uintptr, recordInfo uintptr) uintptr {
			gotArg = recordInfo
			return 1
		}),
	}
	recordInfo := &IRecordInfo{VirtualTable: virtualTable}

	if !recordInfo.IsMatchingType(other) {
		t.Fatal("IsMatchingType() = false, want true")
	}
	if gotArg != uintptr(unsafe.Pointer(other)) {
		t.Fatalf("IsMatchingType received %#x, want %#x", gotArg, uintptr(unsafe.Pointer(other)))
	}
	if !recordInfo.Equals(other) {
		t.Fatal("Equals() = false, want true")
	}
}
