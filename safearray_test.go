//go:build windows

package ole

import (
	"reflect"
	"testing"
)

func TestCreateSupportsMultipleDimensions(t *testing.T) {
	array, err := Create(VT_I4, []SafeArrayBound{
		{Elements: 3, LowerBound: 0},
		{Elements: 2, LowerBound: 0},
	})
	if err != nil {
		t.Fatalf("Create() error = %v", err)
	}
	t.Cleanup(func() {
		_ = array.Destroy()
	})

	dimensions, err := array.GetDimensions()
	if err != nil {
		t.Fatalf("GetDimensions() error = %v", err)
	}
	if dimensions != 2 {
		t.Fatalf("GetDimensions() = %d, want 2", dimensions)
	}

	bounds, err := array.BoundsInfo()
	if err != nil {
		t.Fatalf("BoundsInfo() error = %v", err)
	}
	want := []SafeArrayBound{
		{Elements: 2, LowerBound: 0},
		{Elements: 3, LowerBound: 0},
	}
	if !reflect.DeepEqual(bounds, want) {
		t.Fatalf("BoundsInfo() = %#v, want %#v", bounds, want)
	}
}

func TestFromSliceRoundTripMultiDimensional(t *testing.T) {
	input := [][]int32{
		{1, 2, 3},
		{4, 5, 6},
	}

	array, err := FromSlice(input)
	if err != nil {
		t.Fatalf("FromSlice() error = %v", err)
	}
	t.Cleanup(func() {
		_ = array.Destroy()
	})

	got, err := ToSlice[[][]int32](array)
	if err != nil {
		t.Fatalf("ToSlice() error = %v", err)
	}
	if !reflect.DeepEqual(got, input) {
		t.Fatalf("ToSlice() = %#v, want %#v", got, input)
	}
}

func TestFromSliceRejectsJaggedSlices(t *testing.T) {
	_, err := FromSlice([][]int32{
		{1, 2},
		{3},
	})
	if err == nil {
		t.Fatal("FromSlice() error = nil, want shape error")
	}
}
