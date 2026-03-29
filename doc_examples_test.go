//go:build windows

package ole

import "fmt"

func ExampleInitializeApartmentThreaded() {
	result, err := InitializeApartmentThreaded()
	if err != nil {
		return
	}
	if result == SuccessfullyInitialized || result == AlreadyInitialized {
		defer Uninitialize()
	}
}

func ExampleClassIdFromString() {
	classID, err := ClassIdFromString("{00000000-0000-0000-C000-000000000046}")
	if err != nil {
		return
	}
	fmt.Println(classID == IID_IUnknown)
	// Output:
	// true
}

func ExampleInterfaceIdToString() {
	text, err := InterfaceIdToString(IID_IDispatch)
	if err != nil {
		return
	}
	fmt.Println(text)
	// Output:
	// {00020400-0000-0000-C000-000000000046}
}

func ExampleNewHString() {
	value, err := NewHString("go-ole")
	if err != nil {
		return
	}
	defer DeleteHString(value)

	fmt.Println(value.String())
	// Output:
	// go-ole
}

func ExampleWrapVariant() {
	variant, err := WrapVariant(int32(42))
	if err != nil {
		return
	}
	defer variant.Clear()

	fmt.Println(variant.VT, variant.Val)
	// Output:
	// 3 42
}

func ExampleBoolToVariant() {
	variant := BoolToVariant(true)
	fmt.Println(variant.VT == VT_BOOL)
	fmt.Println(VariantToBool(variant))
	// Output:
	// true
	// true
}

func ExampleMakeDisplayParams() {
	visible := BoolToVariant(true)
	params := MakeDisplayParams(DISPATCH_PROPERTYPUT, visible)

	fmt.Println(params.cArgs, params.cNamedArgs)
	// Output:
	// 1 1
}

func ExampleFromSlice() {
	array, err := FromSlice([][]int32{{1, 2}, {3, 4}})
	if err != nil {
		return
	}
	defer array.Destroy()

	bounds, err := array.BoundsInfo()
	if err != nil {
		return
	}

	fmt.Println(len(bounds), bounds[0].Elements, bounds[1].Elements)
	// Output:
	// 2 2 2
}

func ExampleToSlice() {
	array, err := FromSlice([]string{"red", "blue"})
	if err != nil {
		return
	}
	defer array.Destroy()

	values, err := ToSlice[[]string](array)
	if err != nil {
		return
	}

	fmt.Println(values[0], values[1])
	// Output:
	// red blue
}

func ExampleRoInitialize() {
	result, err := RoInitialize(RoSingleThreaded)
	if err != nil {
		return
	}
	if result == SuccessfullyInitialized || result == AlreadyInitialized {
		defer RoUninitialize()
	}
}
