//go:build windows

package ole

import (
	"syscall"
	"testing"
	"unsafe"

	"golang.org/x/sys/windows"
)

func TestIEnumVariantNext(t *testing.T) {
	t.Run("zero retrieve", func(t *testing.T) {
		enum := &IEnumVariant{}
		if got := enum.Next(0); got != nil {
			t.Fatalf("Next(0) = %#v, want nil", got)
		}
	})

	t.Run("reads returned variants", func(t *testing.T) {
		first := &VARIANT{VT: VT_I4, Val: 11}
		second := &VARIANT{VT: VT_I4, Val: 22}

		virtualTable := &IEnumVariantVirtualTable{
			Next: syscall.NewCallback(func(this uintptr, requested uintptr, array uintptr, fetched uintptr) uintptr {
				items := unsafe.Slice((**VARIANT)(unsafe.Pointer(array)), int(requested))
				items[0] = first
				items[1] = second
				*(*uint32)(unsafe.Pointer(fetched)) = 2
				return uintptr(windows.S_OK)
			}),
		}
		enum := &IEnumVariant{VirtualTable: virtualTable}

		got := enum.Next(2)
		if len(got) != 2 {
			t.Fatalf("Next(2) len = %d, want 2", len(got))
		}
		if got[0] != first || got[1] != second {
			t.Fatalf("Next(2) = %#v, want [%p %p]", got, first, second)
		}
	})
}

func TestIEnumVariantForEach(t *testing.T) {
	type state struct {
		items []*VARIANT
		index int
	}

	makeEnum := func(items []*VARIANT) *IEnumVariant {
		sharedState := &state{items: items}
		virtualTable := &IEnumVariantVirtualTable{
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

				dst := unsafe.Slice((**VARIANT)(unsafe.Pointer(array)), count)
				copy(dst, sharedState.items[sharedState.index:sharedState.index+count])
				sharedState.index += count
				*(*uint32)(unsafe.Pointer(fetched)) = uint32(count)
				if sharedState.index >= len(sharedState.items) {
					return uintptr(windows.S_FALSE)
				}
				return uintptr(windows.S_OK)
			}),
		}
		return &IEnumVariant{VirtualTable: virtualTable}
	}

	t.Run("iterates variants", func(t *testing.T) {
		enum := makeEnum([]*VARIANT{
			{VT: VT_I4, Val: 11},
			{VT: VT_I4, Val: 22},
		})

		var got []int64
		for item := range enum.ForEach {
			got = append(got, item.Val)
		}

		if len(got) != 2 {
			t.Fatalf("ForEach count = %d, want 2", len(got))
		}
		if got[0] != 11 || got[1] != 22 {
			t.Fatalf("ForEach values = %#v, want [11 22]", got)
		}
	})

	t.Run("stops when range breaks", func(t *testing.T) {
		enum := makeEnum([]*VARIANT{
			{VT: VT_I4, Val: 11},
			{VT: VT_I4, Val: 22},
			{VT: VT_I4, Val: 33},
		})

		var calls int
		for item := range enum.ForEach {
			calls++
			if item.Val == 22 {
				break
			}
		}

		if calls != 2 {
			t.Fatalf("ForEach callback calls = %d, want 2", calls)
		}
	})
}
