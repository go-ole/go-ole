//go:build windows

package ole

import (
	"errors"
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestIEnumConnectionsClone(t *testing.T) {
	t.Run("success", func(t *testing.T) {
		want := &IEnumConnections{}
		virtualTable := &IEnumConnectionsVirtualTable{
			Clone: syscall.NewCallback(func(this uintptr, cloned uintptr) uintptr {
				*(**IEnumConnections)(unsafe.Pointer(cloned)) = want
				return uintptr(windows.S_OK)
			}),
		}
		enum := &IEnumConnections{VirtualTable: virtualTable}

		got, err := enum.Clone()
		if err != nil {
			t.Fatalf("Clone failed: %v", err)
		}
		if got != want {
			t.Fatalf("Clone() = %p, want %p", got, want)
		}
	})

	t.Run("out of memory", func(t *testing.T) {
		virtualTable := &IEnumConnectionsVirtualTable{
			Clone: syscall.NewCallback(func(this uintptr, cloned uintptr) uintptr {
				return uintptr(windows.E_OUTOFMEMORY)
			}),
		}
		enum := &IEnumConnections{VirtualTable: virtualTable}

		got, err := enum.Clone()
		if got != nil {
			t.Fatalf("Clone() = %p, want nil", got)
		}
		if !errors.Is(err, EnumOutOfMemoryError) {
			t.Fatalf("Clone() error = %v, want %v", err, EnumOutOfMemoryError)
		}
	})
}

func TestIEnumConnectionsResetAndSkip(t *testing.T) {
	t.Run("reset success", func(t *testing.T) {
		virtualTable := &IEnumConnectionsVirtualTable{
			Reset: syscall.NewCallback(func(this uintptr) uintptr {
				return uintptr(windows.S_OK)
			}),
		}
		enum := &IEnumConnections{VirtualTable: virtualTable}

		if !enum.Reset() {
			t.Fatal("Reset() = false, want true")
		}
	})

	t.Run("reset false", func(t *testing.T) {
		virtualTable := &IEnumConnectionsVirtualTable{
			Reset: syscall.NewCallback(func(this uintptr) uintptr {
				return uintptr(windows.S_FALSE)
			}),
		}
		enum := &IEnumConnections{VirtualTable: virtualTable}

		if enum.Reset() {
			t.Fatal("Reset() = true, want false")
		}
	})

	t.Run("skip success", func(t *testing.T) {
		var gotSkip uint
		virtualTable := &IEnumConnectionsVirtualTable{
			Skip: syscall.NewCallback(func(this uintptr, numSkip uintptr) uintptr {
				gotSkip = uint(numSkip)
				return uintptr(windows.S_OK)
			}),
		}
		enum := &IEnumConnections{VirtualTable: virtualTable}

		if !enum.Skip(3) {
			t.Fatal("Skip() = false, want true")
		}
		if gotSkip != 3 {
			t.Fatalf("Skip() captured %d, want 3", gotSkip)
		}
	})

	t.Run("skip false", func(t *testing.T) {
		virtualTable := &IEnumConnectionsVirtualTable{
			Skip: syscall.NewCallback(func(this uintptr, numSkip uintptr) uintptr {
				return uintptr(windows.S_FALSE)
			}),
		}
		enum := &IEnumConnections{VirtualTable: virtualTable}

		if enum.Skip(3) {
			t.Fatal("Skip() = true, want false")
		}
	})
}

func TestIEnumConnectionsNext(t *testing.T) {
	t.Run("zero retrieve", func(t *testing.T) {
		enum := &IEnumConnections{VirtualTable: &IEnumConnectionsVirtualTable{}}
		if got := enum.Next(0); got != nil {
			t.Fatalf("Next(0) = %#v, want nil", got)
		}
	})

	t.Run("returns retrieved items", func(t *testing.T) {
		virtualTable := &IEnumConnectionsVirtualTable{
			Next: syscall.NewCallback(func(this uintptr, requested uintptr, array uintptr, fetched uintptr) uintptr {
				items := unsafe.Slice((*ConnectData)(unsafe.Pointer(array)), int(requested))
				items[0] = ConnectData{unknown: 11, Cookie: 101}
				items[1] = ConnectData{unknown: 22, Cookie: 202}
				*(*uint32)(unsafe.Pointer(fetched)) = 2
				return uintptr(windows.S_OK)
			}),
		}
		enum := &IEnumConnections{VirtualTable: virtualTable}

		got := enum.Next(2)
		if len(got) != 2 {
			t.Fatalf("Next(2) len = %d, want 2", len(got))
		}
		if got[0].unknown != 11 || got[0].Cookie != 101 || got[1].unknown != 22 || got[1].Cookie != 202 {
			t.Fatalf("Next(2) = %#v, want filled items", got)
		}
	})

	t.Run("short read", func(t *testing.T) {
		virtualTable := &IEnumConnectionsVirtualTable{
			Next: syscall.NewCallback(func(this uintptr, requested uintptr, array uintptr, fetched uintptr) uintptr {
				items := unsafe.Slice((*ConnectData)(unsafe.Pointer(array)), int(requested))
				items[0] = ConnectData{unknown: 33, Cookie: 303}
				*(*uint32)(unsafe.Pointer(fetched)) = 1
				return uintptr(windows.S_FALSE)
			}),
		}
		enum := &IEnumConnections{VirtualTable: virtualTable}

		got := enum.Next(2)
		if len(got) != 1 {
			t.Fatalf("Next(2) len = %d, want 1", len(got))
		}
		if got[0].unknown != 33 || got[0].Cookie != 303 {
			t.Fatalf("Next(2) = %#v, want short read item", got)
		}
	})
}

func TestIEnumConnectionsForEach(t *testing.T) {
	type state struct {
		items []ConnectData
		index int
	}

	makeEnum := func(items []ConnectData) (*IEnumConnections, *state) {
		sharedState := &state{items: items}
		virtualTable := &IEnumConnectionsVirtualTable{
			Reset: syscall.NewCallback(func(this uintptr) uintptr {
				sharedState.index = 0
				return uintptr(windows.S_OK)
			}),
			Next: syscall.NewCallback(func(this uintptr, requested uintptr, array uintptr, fetched uintptr) uintptr {
				if sharedState.index >= len(sharedState.items) {
					*(*uint32)(unsafe.Pointer(fetched)) = 0
					return uintptr(windows.S_FALSE)
				}

				remaining := len(sharedState.items) - sharedState.index
				count := int(requested)
				if remaining < count {
					count = remaining
				}

				dst := unsafe.Slice((*ConnectData)(unsafe.Pointer(array)), count)
				copy(dst, sharedState.items[sharedState.index:sharedState.index+count])
				sharedState.index += count
				*(*uint32)(unsafe.Pointer(fetched)) = uint32(count)
				if sharedState.index >= len(sharedState.items) {
					return uintptr(windows.S_FALSE)
				}
				return uintptr(windows.S_OK)
			}),
		}
		return &IEnumConnections{VirtualTable: virtualTable}, sharedState
	}

	t.Run("iterates items", func(t *testing.T) {
		enum, _ := makeEnum([]ConnectData{
			{unknown: 11, Cookie: 101},
			{unknown: 22, Cookie: 202},
		})

		var got []ConnectData
		for item := range enum.ForEach {
			got = append(got, *item)
		}
		if len(got) != 2 {
			t.Fatalf("ForEach count = %d, want 2", len(got))
		}
		if got[0].unknown != 11 || got[0].Cookie != 101 || got[1].unknown != 22 || got[1].Cookie != 202 {
			t.Fatalf("ForEach items = %#v, want original items", got)
		}
	})

	t.Run("stops when range breaks", func(t *testing.T) {
		enum, _ := makeEnum([]ConnectData{
			{unknown: 11, Cookie: 101},
			{unknown: 22, Cookie: 202},
			{unknown: 33, Cookie: 303},
		})
		var calls int

		for item := range enum.ForEach {
			calls++
			if item.Cookie == 202 {
				break
			}
		}
		if calls != 2 {
			t.Fatalf("ForEach callback calls = %d, want 2", calls)
		}
	})
}
