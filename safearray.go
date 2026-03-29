//go:build windows

package ole

import (
	"errors"
	"fmt"
	"reflect"
	"unsafe"

	"golang.org/x/sys/windows"
)

const (
	FADF_AUTO        = 0x0001
	FADF_STATIC      = 0x0002
	FADF_EMBEDDED    = 0x0004
	FADF_FIXEDSIZE   = 0x0010
	FADF_RECORD      = 0x0020
	FADF_HAVEIID     = 0x0040
	FADF_HAVEVARTYPE = 0x0080
	FADF_BSTR        = 0x0100
	FADF_UNKNOWN     = 0x0200
	FADF_DISPATCH    = 0x0400
	FADF_VARIANT     = 0x0800
	FADF_RESERVED    = 0xF008
)

// SafeArrayBound describes one SAFEARRAY dimension in Go slice order.
type SafeArrayBound struct {
	Elements   uint32
	LowerBound int32
}

// SafeArray mirrors the Windows SAFEARRAY descriptor.
//
// Most callers should prefer the higher-level helpers such as FromSlice,
// ToSlice, MarshalSafeArray, and UnmarshalSafeArray instead of constructing or
// mutating this struct directly.
type SafeArray struct {
	Dimensions   uint16
	FeaturesFlag uint16
	ElementsSize uint32
	LocksAmount  uint32
	Data         uintptr
	Bounds       [1]SafeArrayBound
}

var (
	ArgumentNotSafeArrayError = errors.New("invalid safe array")
	BadIndexSafeArrayError    = errors.New("bad safe array index")
	OutOfMemoryError          = errors.New("out of memory")
	ArrayLockedError          = errors.New("safe array is locked")
	UnsupportedElementType    = errors.New("unsupported safe array element type")
	UnsupportedSliceShape     = errors.New("unsupported safe array slice shape")
)

var variantTypeType = reflect.TypeOf(VARIANT{})

// Create allocates a SAFEARRAY with the requested element type and bounds.
//
// Example:
//
//	sa, err := ole.Create(ole.VT_I4, []ole.SafeArrayBound{
//		{Elements: 3, LowerBound: 0},
//	})
//	if err != nil {
//		return err
//	}
//	defer sa.Destroy()
func Create(variantType VT, bounds []SafeArrayBound) (*SafeArray, error) {
	if len(bounds) == 0 {
		return nil, ArgumentNotSafeArrayError
	}

	ptr, err := SafeArrayCreate(variantType, uint32(len(bounds)), unsafe.Pointer(&bounds[0]))
	if ptr == 0 {
		if err != nil && err != windows.ERROR_SUCCESS {
			return nil, err
		}
		return nil, OutOfMemoryError
	}
	return (*SafeArray)(unsafe.Pointer(ptr)), nil
}

// CreateEx allocates a SAFEARRAY with extra type-specific metadata.
//
// This is primarily useful for record, interface, or custom element types that
// require additional runtime information.
func CreateEx(variantType VT, bounds []SafeArrayBound, extra uintptr) (*SafeArray, error) {
	if len(bounds) == 0 {
		return nil, ArgumentNotSafeArrayError
	}

	ptr, err := SafeArrayCreateEx(variantType, uint32(len(bounds)), unsafe.Pointer(&bounds[0]), extra)
	if ptr == 0 {
		if err != nil && err != windows.ERROR_SUCCESS {
			return nil, err
		}
		return nil, OutOfMemoryError
	}
	return (*SafeArray)(unsafe.Pointer(ptr)), nil
}

// CreateVector allocates a one-dimensional SAFEARRAY with a contiguous range.
func CreateVector(variantType VT, lowerBound int32, length uint32) (*SafeArray, error) {
	ptr, err := SafeArrayCreateVector(variantType, lowerBound, length)
	if ptr == 0 {
		if err != nil && err != windows.ERROR_SUCCESS {
			return nil, err
		}
		return nil, OutOfMemoryError
	}
	return (*SafeArray)(unsafe.Pointer(ptr)), nil
}

// CreateVectorEx allocates a one-dimensional SAFEARRAY with extra type-specific metadata.
func CreateVectorEx(variantType VT, lowerBound int32, length uint32, extra uintptr) (*SafeArray, error) {
	ptr, err := SafeArrayCreateVectorEx(variantType, lowerBound, length, extra)
	if ptr == 0 {
		if err != nil && err != windows.ERROR_SUCCESS {
			return nil, err
		}
		return nil, OutOfMemoryError
	}
	return (*SafeArray)(unsafe.Pointer(ptr)), nil
}

// AllocDescriptor allocates an empty SAFEARRAY descriptor for manual population.
//
// After calling AllocDescriptor you typically set bounds and then call
// AllocData, or destroy the descriptor if setup fails.
func AllocDescriptor(dimensions uint32) (*SafeArray, error) {
	var array *SafeArray
	hr := SafeArrayAllocDescriptor(dimensions, unsafe.Pointer(&array))
	if err := convertHRESULT(hr); err != nil {
		return nil, err
	}
	return array, nil
}

// AllocDescriptorEx allocates an empty SAFEARRAY descriptor with extended type metadata.
func AllocDescriptorEx(variantType VT, dimensions uint32) (*SafeArray, error) {
	var array *SafeArray
	hr := SafeArrayAllocDescriptorEx(variantType, dimensions, unsafe.Pointer(&array))
	if err := convertHRESULT(hr); err != nil {
		return nil, err
	}
	return array, nil
}

// MarshalSafeArray converts a rectangular Go slice into a SAFEARRAY.
//
// Example:
//
//	sa, err := ole.MarshalSafeArray([][]int32{{1, 2}, {3, 4}})
//	if err != nil {
//		return err
//	}
//	defer sa.Destroy()
func MarshalSafeArray[T any](value T) (*SafeArray, error) {
	return FromSlice(value)
}

// WrapSliceAsVariant converts a Go slice into a VT_ARRAY VARIANT.
//
// Example:
//
//	v, err := ole.WrapSliceAsVariant([]string{"alpha", "beta"})
//	if err != nil {
//		return err
//	}
//	defer v.Clear()
func WrapSliceAsVariant[T any](value T) (*VARIANT, error) {
	array, err := FromSlice(value)
	if err != nil {
		return nil, err
	}
	variantType, err := array.GetVarType()
	if err != nil {
		_ = array.Destroy()
		return nil, err
	}
	return &VARIANT{
		VT:  VT_ARRAY | variantType,
		Val: int64(uintptr(unsafe.Pointer(array))),
	}, nil
}

// UnmarshalSafeArray converts a SAFEARRAY into a Go slice of the requested type.
//
// Example:
//
//	values, err := ole.UnmarshalSafeArray[[]string](sa)
//	if err != nil {
//		return err
//	}
func UnmarshalSafeArray[T any](sa *SafeArray) (T, error) {
	return ToSlice[T](sa)
}

// FromSlice converts a rectangular Go slice into a SAFEARRAY.
//
// Multi-dimensional slices are supported as long as each nested slice has the
// same length.
//
// Example:
//
//	sa, err := ole.FromSlice([][]float64{{1.5, 2.5}, {3.5, 4.5}})
//	if err != nil {
//		return err
//	}
//	defer sa.Destroy()
func FromSlice[T any](value T) (*SafeArray, error) {
	root := reflect.ValueOf(value)
	if !root.IsValid() || root.Kind() != reflect.Slice {
		return nil, UnsupportedSliceShape
	}

	lengths, elementType, err := describeSlice(root)
	if err != nil {
		return nil, err
	}

	variantType, err := variantTypeForType(elementType)
	if err != nil {
		return nil, err
	}

	bounds := make([]SafeArrayBound, len(lengths))
	for index, length := range lengths {
		bounds[len(lengths)-1-index] = SafeArrayBound{Elements: uint32(length)}
	}

	sa, err := Create(variantType, bounds)
	if err != nil {
		return nil, err
	}

	if err := populateSafeArray(sa, root, variantType, make([]int32, len(lengths))); err != nil {
		_ = sa.Destroy()
		return nil, err
	}

	return sa, nil
}

// ToSlice converts a SAFEARRAY into a Go slice of the requested type.
//
// Example:
//
//	values, err := ole.ToSlice[[][]int32](sa)
//	if err != nil {
//		return err
//	}
func ToSlice[T any](sa *SafeArray) (T, error) {
	var zero T
	if sa == nil {
		return zero, ArgumentNotSafeArrayError
	}

	targetType := reflect.TypeFor[T]()
	if targetType.Kind() != reflect.Slice {
		return zero, UnsupportedSliceShape
	}

	bounds, err := sa.BoundsInfo()
	if err != nil {
		return zero, err
	}

	if len(bounds) != sliceDepth(targetType) {
		return zero, fmt.Errorf("%w: target depth %d does not match safe array dimensions %d", UnsupportedSliceShape, sliceDepth(targetType), len(bounds))
	}

	value, err := buildSliceValue(sa, targetType, bounds, 0, make([]int32, len(bounds)))
	if err != nil {
		return zero, err
	}

	return value.Interface().(T), nil
}

// Destroy releases a SAFEARRAY and any resources owned by it.
func (sa *SafeArray) Destroy() error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}

	hr := SafeArrayDestroy(unsafe.Pointer(sa))
	return convertHRESULT(hr)
}

// AccessData pins the SAFEARRAY and returns a pointer to its backing storage.
//
// Call UnaccessData after finishing direct memory access.
//
// Example:
//
//	ptr, err := sa.AccessData()
//	if err != nil {
//		return err
//	}
//	defer sa.UnaccessData()
func (sa *SafeArray) AccessData() (uintptr, error) {
	if sa == nil {
		return 0, ArgumentNotSafeArrayError
	}

	var data uintptr
	hr := SafeArrayAccessData(unsafe.Pointer(sa), unsafe.Pointer(&data))
	if err := convertHRESULT(hr); err != nil {
		return 0, err
	}
	return data, nil
}

// AddRef increments the SAFEARRAY pin count and returns release data for ReleaseData.
func (sa *SafeArray) AddRef() (uintptr, error) {
	if sa == nil {
		return 0, ArgumentNotSafeArrayError
	}
	var dataToRelease uintptr
	hr := SafeArrayAddRef(unsafe.Pointer(sa), unsafe.Pointer(&dataToRelease))
	if err := convertHRESULT(hr); err != nil {
		return 0, err
	}
	return dataToRelease, nil
}

// AllocData allocates storage for a manually allocated SAFEARRAY descriptor.
func (sa *SafeArray) AllocData() error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}
	hr := SafeArrayAllocData(unsafe.Pointer(sa))
	return convertHRESULT(hr)
}

// Copy duplicates the SAFEARRAY descriptor and data into a new SAFEARRAY.
func (sa *SafeArray) Copy() (*SafeArray, error) {
	if sa == nil {
		return nil, ArgumentNotSafeArrayError
	}

	var copy *SafeArray
	hr := SafeArrayCopy(unsafe.Pointer(sa), unsafe.Pointer(&copy))
	if err := convertHRESULT(hr); err != nil {
		return nil, err
	}
	return copy, nil
}

// CopyData copies element data into another SAFEARRAY with a compatible shape.
func (sa *SafeArray) CopyData(target *SafeArray) error {
	if sa == nil || target == nil {
		return ArgumentNotSafeArrayError
	}
	hr := SafeArrayCopyData(unsafe.Pointer(sa), unsafe.Pointer(target))
	return convertHRESULT(hr)
}

// DestroyData releases only the SAFEARRAY element storage.
func (sa *SafeArray) DestroyData() error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}
	hr := SafeArrayDestroyData(unsafe.Pointer(sa))
	return convertHRESULT(hr)
}

// DestroyDescriptor releases only the SAFEARRAY descriptor.
func (sa *SafeArray) DestroyDescriptor() error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}
	hr := SafeArrayDestroyDescriptor(unsafe.Pointer(sa))
	return convertHRESULT(hr)
}

// GetDimensions returns the number of dimensions in the SAFEARRAY.
func (sa *SafeArray) GetDimensions() (uint32, error) {
	if sa == nil {
		return 0, ArgumentNotSafeArrayError
	}

	ret := SafeArrayGetDim(unsafe.Pointer(sa))
	return uint32(ret), nil
}

// GetElementSize returns the size in bytes of each SAFEARRAY element.
func (sa *SafeArray) GetElementSize() (uint32, error) {
	if sa == nil {
		return 0, ArgumentNotSafeArrayError
	}

	ret := SafeArrayGetElemsize(unsafe.Pointer(sa))
	return uint32(ret), nil
}

// GetIID returns the interface identifier associated with SAFEARRAY interface elements.
func (sa *SafeArray) GetIID() (windows.GUID, error) {
	if sa == nil {
		return windows.GUID{}, ArgumentNotSafeArrayError
	}

	var guid windows.GUID
	hr := SafeArrayGetIID(unsafe.Pointer(sa), unsafe.Pointer(&guid))
	if err := convertHRESULT(hr); err != nil {
		return windows.GUID{}, err
	}
	return guid, nil
}

// GetVarType returns the VARTYPE stored by the SAFEARRAY.
func (sa *SafeArray) GetVarType() (VT, error) {
	if sa == nil {
		return 0, ArgumentNotSafeArrayError
	}

	var varType uint16
	hr := SafeArrayGetVartype(unsafe.Pointer(sa), unsafe.Pointer(&varType))
	if err := convertHRESULT(hr); err != nil {
		return 0, err
	}
	return VT(varType), nil
}

// GetRecordInfo returns the record metadata associated with SAFEARRAY record elements.
func (sa *SafeArray) GetRecordInfo() (*IRecordInfo, error) {
	if sa == nil {
		return nil, ArgumentNotSafeArrayError
	}

	var recordInfo *IRecordInfo
	hr := SafeArrayGetRecordInfo(unsafe.Pointer(sa), unsafe.Pointer(&recordInfo))
	if err := convertHRESULT(hr); err != nil {
		return nil, err
	}
	return recordInfo, nil
}

// BoundsInfo returns bounds in Go slice order, from outermost to innermost dimension.
//
// Example:
//
//	bounds, err := sa.BoundsInfo()
//	if err != nil {
//		return err
//	}
//	_ = bounds[0].Elements
func (sa *SafeArray) BoundsInfo() ([]SafeArrayBound, error) {
	dimensions, err := sa.GetDimensions()
	if err != nil {
		return nil, err
	}

	bounds := make([]SafeArrayBound, dimensions)
	for dimension := uint32(1); dimension <= dimensions; dimension++ {
		lower, err := safeArrayGetLBound(sa, dimension)
		if err != nil {
			return nil, err
		}
		upper, err := safeArrayGetUBound(sa, dimension)
		if err != nil {
			return nil, err
		}

		target := len(bounds) - int(dimension)
		bounds[target] = SafeArrayBound{
			Elements:   uint32(upper - lower + 1),
			LowerBound: lower,
		}
	}

	return bounds, nil
}

// PutElement writes a single element at the requested indices.
//
// Indices are passed in Go slice order.
//
// Example:
//
//	if err := sa.PutElement([]int32{1, 2}, int32(42)); err != nil {
//		return err
//	}
func (sa *SafeArray) PutElement(indices []int32, value any) error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}

	if len(indices) == 0 {
		return BadIndexSafeArrayError
	}

	comIndices := reverseIndices(indices)
	return putElementValue(sa, comIndices, reflect.ValueOf(value))
}

// Lock increments the SAFEARRAY lock count.
func (sa *SafeArray) Lock() error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}
	hr := SafeArrayLock(unsafe.Pointer(sa))
	return convertHRESULT(hr)
}

// PtrOfIndex returns a direct pointer to the element at the requested indices.
//
// Indices are passed in Go slice order.
func (sa *SafeArray) PtrOfIndex(indices []int32) (uintptr, error) {
	if sa == nil || len(indices) == 0 {
		return 0, ArgumentNotSafeArrayError
	}
	comIndices := reverseIndices(indices)
	var element uintptr
	hr := SafeArrayPtrOfIndex(unsafe.Pointer(sa), unsafe.Pointer(&comIndices[0]), unsafe.Pointer(&element))
	if err := convertHRESULT(hr); err != nil {
		return 0, err
	}
	return element, nil
}

// Redim resizes the right-most SAFEARRAY dimension.
func (sa *SafeArray) Redim(bound SafeArrayBound) error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}
	hr := SafeArrayRedim(unsafe.Pointer(sa), unsafe.Pointer(&bound))
	return convertHRESULT(hr)
}

// ReleaseData releases the token returned by AddRef.
func (sa *SafeArray) ReleaseData(data uintptr) error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}
	SafeArrayReleaseData(data)
	return nil
}

// ReleaseDescriptor releases a descriptor previously pinned for manual lifetime management.
func (sa *SafeArray) ReleaseDescriptor() error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}
	SafeArrayReleaseDescriptor(unsafe.Pointer(sa))
	return nil
}

// SetIID associates interface element metadata with the SAFEARRAY.
func (sa *SafeArray) SetIID(guid *windows.GUID) error {
	if sa == nil || guid == nil {
		return ArgumentNotSafeArrayError
	}
	hr := SafeArraySetIID(unsafe.Pointer(sa), unsafe.Pointer(guid))
	return convertHRESULT(hr)
}

// SetRecordInfo associates record metadata with the SAFEARRAY.
func (sa *SafeArray) SetRecordInfo(recordInfo *IRecordInfo) error {
	if sa == nil || recordInfo == nil {
		return ArgumentNotSafeArrayError
	}
	hr := SafeArraySetRecordInfo(unsafe.Pointer(sa), unsafe.Pointer(recordInfo))
	return convertHRESULT(hr)
}

// UnaccessData releases a prior AccessData call.
func (sa *SafeArray) UnaccessData() error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}
	hr := SafeArrayUnaccessData(unsafe.Pointer(sa))
	return convertHRESULT(hr)
}

// Unlock decrements the SAFEARRAY lock count.
func (sa *SafeArray) Unlock() error {
	if sa == nil {
		return ArgumentNotSafeArrayError
	}
	hr := SafeArrayUnlock(unsafe.Pointer(sa))
	return convertHRESULT(hr)
}

// InterfaceSlice converts a one-dimensional SAFEARRAY into a []any.
//
// Example:
//
//	values, err := sa.InterfaceSlice()
//	if err != nil {
//		return err
//	}
func (sa *SafeArray) InterfaceSlice() ([]any, error) {
	if sa == nil {
		return nil, ArgumentNotSafeArrayError
	}

	bounds, err := sa.BoundsInfo()
	if err != nil {
		return nil, err
	}
	if len(bounds) != 1 {
		return nil, fmt.Errorf("%w: InterfaceSlice only supports one-dimensional arrays", UnsupportedSliceShape)
	}

	values := make([]any, bounds[0].Elements)
	for index := range values {
		actualIndex := []int32{bounds[0].LowerBound + int32(index)}
		value, err := valueAtIndices(sa, actualIndex)
		if err != nil {
			return nil, err
		}
		values[index] = value
	}

	return values, nil
}

func safeArrayCreate(variantType VT, dimensions uint32, bounds *SafeArrayBound) (*SafeArray, error) {
	if bounds == nil || dimensions == 0 {
		return nil, ArgumentNotSafeArrayError
	}
	return Create(variantType, unsafe.Slice(bounds, dimensions))
}

func safeArrayCreateEx(variantType VT, dimensions uint32, bounds *SafeArrayBound, extra uintptr) (*SafeArray, error) {
	if bounds == nil || dimensions == 0 {
		return nil, ArgumentNotSafeArrayError
	}
	return CreateEx(variantType, unsafe.Slice(bounds, dimensions), extra)
}

func safeArrayCreateVector(variantType VT, lowerBound int32, length uint32) (*SafeArray, error) {
	return CreateVector(variantType, lowerBound, length)
}

func safeArrayCreateVectorEx(variantType VT, lowerBound int32, length uint32, extra uintptr) (*SafeArray, error) {
	return CreateVectorEx(variantType, lowerBound, length, extra)
}

func safeArrayDestroy(array *SafeArray) error {
	return array.Destroy()
}

func safeArrayGetDim(array *SafeArray) (uint32, error) {
	return array.GetDimensions()
}

func safeArrayGetElementSize(array *SafeArray) (uint32, error) {
	return array.GetElementSize()
}

func safeArrayGetVartype(array *SafeArray) (uint16, error) {
	vt, err := array.GetVarType()
	return uint16(vt), err
}

func safeArrayGetLBound(array *SafeArray, dimension uint32) (int32, error) {
	var lowerBound int32
	hr := SafeArrayGetLBound(unsafe.Pointer(array), dimension, unsafe.Pointer(&lowerBound))
	if err := convertHRESULT(hr); err != nil {
		return 0, err
	}
	return lowerBound, nil
}

func safeArrayGetUBound(array *SafeArray, dimension uint32) (int32, error) {
	var upperBound int32
	hr := SafeArrayGetUBound(unsafe.Pointer(array), dimension, unsafe.Pointer(&upperBound))
	if err := convertHRESULT(hr); err != nil {
		return 0, err
	}
	return upperBound, nil
}

func safeArrayPutElement(array *SafeArray, index int64, element uintptr) error {
	indices := []int32{int32(index)}
	hr := SafeArrayPutElement(unsafe.Pointer(array), unsafe.Pointer(&indices[0]), element)
	return convertHRESULT(hr)
}

func safeArrayGetElement(array *SafeArray, index int32, element unsafe.Pointer) error {
	indices := []int32{index}
	hr := SafeArrayGetElement(unsafe.Pointer(array), unsafe.Pointer(&indices[0]), uintptr(element))
	return convertHRESULT(hr)
}

func safeArrayGetElementString(array *SafeArray, index int32) (string, error) {
	var bstr *uint16
	if err := safeArrayGetElement(array, index, unsafe.Pointer(&bstr)); err != nil {
		return "", err
	}
	if bstr == nil {
		return "", nil
	}
	defer SysFreeString(bstr)
	return windows.UTF16PtrToString(bstr), nil
}

func convertHRESULT(hr uintptr) error {
	switch windows.Handle(hr) {
	case windows.S_OK:
		return nil
	case windows.DISP_E_BADINDEX:
		return BadIndexSafeArrayError
	case windows.E_INVALIDARG:
		return ArgumentNotSafeArrayError
	case windows.E_OUTOFMEMORY:
		return OutOfMemoryError
	case windows.DISP_E_ARRAYISLOCKED:
		return ArrayLockedError
	default:
		if hr == 0 {
			return nil
		}
		return windows.Errno(hr)
	}
}

func sliceDepth(t reflect.Type) int {
	depth := 0
	for t.Kind() == reflect.Slice {
		depth++
		t = t.Elem()
	}
	return depth
}

func describeSlice(value reflect.Value) ([]int, reflect.Type, error) {
	if value.Kind() != reflect.Slice {
		return nil, nil, UnsupportedSliceShape
	}

	lengths := make([]int, 0, sliceDepth(value.Type()))
	currentType := value.Type()
	currentValue := value
	for currentType.Kind() == reflect.Slice {
		lengths = append(lengths, currentValue.Len())
		if currentValue.Len() == 0 {
			currentType = currentType.Elem()
			currentValue = reflect.Zero(currentType)
			continue
		}

		first := currentValue.Index(0)
		if first.Kind() == reflect.Slice {
			for index := 1; index < currentValue.Len(); index++ {
				if currentValue.Index(index).Len() != first.Len() {
					return nil, nil, fmt.Errorf("%w: jagged slices are not supported", UnsupportedSliceShape)
				}
			}
		}

		currentType = currentType.Elem()
		currentValue = first
	}

	return lengths, currentType, nil
}

func variantTypeForType(t reflect.Type) (VT, error) {
	switch t.Kind() {
	case reflect.Bool:
		return VT_BOOL, nil
	case reflect.Int8:
		return VT_I1, nil
	case reflect.Int16:
		return VT_I2, nil
	case reflect.Int32:
		return VT_I4, nil
	case reflect.Int64:
		return VT_I8, nil
	case reflect.Uint8:
		return VT_UI1, nil
	case reflect.Uint16:
		return VT_UI2, nil
	case reflect.Uint32:
		return VT_UI4, nil
	case reflect.Uint64:
		return VT_UI8, nil
	case reflect.Float32:
		return VT_R4, nil
	case reflect.Float64:
		return VT_R8, nil
	case reflect.String:
		return VT_BSTR, nil
	}

	if t == variantTypeType {
		return VT_VARIANT, nil
	}

	return 0, fmt.Errorf("%w: %s", UnsupportedElementType, t)
}

func populateSafeArray(sa *SafeArray, value reflect.Value, variantType VT, indices []int32) error {
	if value.Kind() != reflect.Slice {
		return putReflectValue(sa, indices, value, variantType)
	}

	for index := 0; index < value.Len(); index++ {
		indices[len(indices)-sliceDepth(value.Type())] = int32(index)
		if err := populateSafeArray(sa, value.Index(index), variantType, indices); err != nil {
			return err
		}
	}
	return nil
}

func putReflectValue(sa *SafeArray, indices []int32, value reflect.Value, variantType VT) error {
	comIndices := reverseIndices(indices)

	switch variantType {
	case VT_BOOL:
		var raw int16
		if value.Bool() {
			raw = -1
		}
		return putElementRaw(sa, comIndices, uintptr(unsafe.Pointer(&raw)))
	case VT_BSTR:
		bstr := SysAllocStringLen(value.String())
		err := putElementRaw(sa, comIndices, uintptr(unsafe.Pointer(bstr)))
		if err != nil {
			_ = SysFreeString(bstr)
		}
		return err
	case VT_VARIANT:
		if value.Type() != variantTypeType {
			return fmt.Errorf("%w: expected VARIANT, got %s", UnsupportedElementType, value.Type())
		}
		v := value.Interface().(VARIANT)
		return putElementRaw(sa, comIndices, uintptr(unsafe.Pointer(&v)))
	default:
		if !value.CanAddr() {
			copyValue := reflect.New(value.Type()).Elem()
			copyValue.Set(value)
			value = copyValue
		}
		return putElementRaw(sa, comIndices, value.Addr().Pointer())
	}
}

func putElementValue(sa *SafeArray, indices []int32, value reflect.Value) error {
	variantType, err := variantTypeForType(value.Type())
	if err != nil {
		return err
	}
	return putReflectValue(sa, reverseIndices(indices), value, variantType)
}

func putElementRaw(sa *SafeArray, indices []int32, element uintptr) error {
	hr := SafeArrayPutElement(unsafe.Pointer(sa), unsafe.Pointer(&indices[0]), element)
	return convertHRESULT(hr)
}

func reverseIndices(indices []int32) []int32 {
	reversed := make([]int32, len(indices))
	for index := range indices {
		reversed[index] = indices[len(indices)-1-index]
	}
	return reversed
}

func buildSliceValue(sa *SafeArray, targetType reflect.Type, bounds []SafeArrayBound, depth int, indices []int32) (reflect.Value, error) {
	length := int(bounds[depth].Elements)
	result := reflect.MakeSlice(targetType, length, length)

	if targetType.Elem().Kind() != reflect.Slice {
		for index := 0; index < length; index++ {
			indices[depth] = bounds[depth].LowerBound + int32(index)
			value, err := reflectValueAt(sa, indices, targetType.Elem())
			if err != nil {
				return reflect.Value{}, err
			}
			result.Index(index).Set(value)
		}
		return result, nil
	}

	for index := 0; index < length; index++ {
		indices[depth] = bounds[depth].LowerBound + int32(index)
		value, err := buildSliceValue(sa, targetType.Elem(), bounds, depth+1, indices)
		if err != nil {
			return reflect.Value{}, err
		}
		result.Index(index).Set(value)
	}

	return result, nil
}

func reflectValueAt(sa *SafeArray, indices []int32, targetType reflect.Type) (reflect.Value, error) {
	comIndices := reverseIndices(indices)

	switch targetType.Kind() {
	case reflect.Bool:
		var raw int16
		if err := getElementRaw(sa, comIndices, uintptr(unsafe.Pointer(&raw))); err != nil {
			return reflect.Value{}, err
		}
		return reflect.ValueOf(raw != 0).Convert(targetType), nil
	case reflect.String:
		var bstr *uint16
		if err := getElementRaw(sa, comIndices, uintptr(unsafe.Pointer(&bstr))); err != nil {
			return reflect.Value{}, err
		}
		defer SysFreeString(bstr)
		value := reflect.New(targetType).Elem()
		value.SetString(windows.UTF16PtrToString(bstr))
		return value, nil
	default:
		if targetType == variantTypeType {
			var value VARIANT
			if err := getElementRaw(sa, comIndices, uintptr(unsafe.Pointer(&value))); err != nil {
				return reflect.Value{}, err
			}
			return reflect.ValueOf(value), nil
		}

		ptr := reflect.New(targetType)
		if err := getElementRaw(sa, comIndices, ptr.Pointer()); err != nil {
			return reflect.Value{}, err
		}
		return ptr.Elem(), nil
	}
}

func valueAtIndices(sa *SafeArray, indices []int32) (any, error) {
	vt, err := sa.GetVarType()
	if err != nil {
		return nil, err
	}

	switch vt {
	case VT_BOOL:
		var value int16
		if err := getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value))); err != nil {
			return nil, err
		}
		return value != 0, nil
	case VT_I1:
		var value int8
		return value, getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value)))
	case VT_I2:
		var value int16
		return value, getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value)))
	case VT_I4:
		var value int32
		return value, getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value)))
	case VT_I8:
		var value int64
		return value, getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value)))
	case VT_UI1:
		var value uint8
		return value, getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value)))
	case VT_UI2:
		var value uint16
		return value, getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value)))
	case VT_UI4:
		var value uint32
		return value, getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value)))
	case VT_UI8:
		var value uint64
		return value, getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value)))
	case VT_R4:
		var value float32
		return value, getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value)))
	case VT_R8:
		var value float64
		return value, getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value)))
	case VT_BSTR:
		var bstr *uint16
		if err := getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&bstr))); err != nil {
			return nil, err
		}
		defer SysFreeString(bstr)
		return windows.UTF16PtrToString(bstr), nil
	case VT_VARIANT:
		var value VARIANT
		if err := getElementRaw(sa, reverseIndices(indices), uintptr(unsafe.Pointer(&value))); err != nil {
			return nil, err
		}
		defer value.Clear()
		conversions.lock.RLock()
		callback, ok := conversions.from[value.VT]
		conversions.lock.RUnlock()
		if !ok {
			return nil, fmt.Errorf("%w: inner vartype %d", UnsupportedElementType, value.VT)
		}
		return callback(&value), nil
	default:
		return nil, fmt.Errorf("%w: vartype %d", UnsupportedElementType, vt)
	}
}

func getElementRaw(sa *SafeArray, indices []int32, element uintptr) error {
	hr := SafeArrayGetElement(unsafe.Pointer(sa), unsafe.Pointer(&indices[0]), element)
	return convertHRESULT(hr)
}
