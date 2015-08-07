// +build !windows

package ole

// safeArrayAccessData returns raw array pointer.
//
// AKA: SafeArrayAccessData in Windows API.
func safeArrayAccessData(safearray *SafeArray) (uintptr, error) {
	return uintptr(0), NewError(E_NOTIMPL)
}

// safeArrayUnaccessData releases raw array.
//
// AKA: SafeArrayUnaccessData in Windows API.
func safeArrayUnaccessData(safearray *SafeArray) error {
	return NewError(E_NOTIMPL)
}

// safeArrayAllocData allocates SafeArray.
//
// AKA: SafeArrayAllocData in Windows API.
func safeArrayAllocData(safearray *SafeArray) error {
	return NewError(E_NOTIMPL)
}

// safeArrayAllocDescriptor allocates SafeArray.
//
// AKA: SafeArrayAllocDescriptor in Windows API.
func safeArrayAllocDescriptor(dimensions uint32) (*SafeArray, error) {
	return nil, NewError(E_NOTIMPL)
}

// safeArrayAllocDescriptorEx allocates SafeArray.
//
// AKA: SafeArrayAllocDescriptorEx in Windows API.
func safeArrayAllocDescriptorEx(variantType VT, dimensions uint32) (*SafeArray, error) {
	return nil, NewError(E_NOTIMPL)
}

// safeArrayCopy returns copy of SafeArray.
//
// AKA: SafeArrayCopy in Windows API.
func safeArrayCopy(original *SafeArray) (*SafeArray, error) {
	return nil, NewError(E_NOTIMPL)
}

// safeArrayCopyData duplicates SafeArray into another SafeArray object.
//
// AKA: SafeArrayCopyData in Windows API.
func safeArrayCopyData(original *SafeArray, duplicate *SafeArray) error {
	return NewError(E_NOTIMPL)
}

// safeArrayCreate creates SafeArray.
//
// AKA: SafeArrayCreate in Windows API.
func safeArrayCreate(variantType VT, dimensions uint32, bounds *SafeArrayBound) (*SafeArray, error) {
	return nil, NewError(E_NOTIMPL)
}

// safeArrayCreateEx creates SafeArray.
//
// AKA: SafeArrayCreateEx in Windows API.
func safeArrayCreateEx(variantType VT, dimensions uint32, bounds *SafeArrayBound, extra uintptr) (*SafeArray, error) {
	return nil, NewError(E_NOTIMPL)
}

// safeArrayCreateVector creates SafeArray.
//
// AKA: SafeArrayCreateVector in Windows API.
func safeArrayCreateVector(variantType VT, lowerBound int32, length uint32) (*SafeArray, error) {
	return nil, NewError(E_NOTIMPL)
}

// safeArrayCreateVectorEx creates SafeArray.
//
// AKA: SafeArrayCreateVectorEx in Windows API.
func safeArrayCreateVectorEx(variantType VT, lowerBound int32, length uint32, extra uintptr) (*SafeArray, error) {
	return nil, NewError(E_NOTIMPL)
}

// safeArrayDestroy destroys SafeArray object.
//
// AKA: SafeArrayDestroy in Windows API.
func safeArrayDestroy(safearray *SafeArray) error {
	return NewError(E_NOTIMPL)
}

// safeArrayDestroyData destroys SafeArray object.
//
// AKA: SafeArrayDestroyData in Windows API.
func safeArrayDestroyData(safearray *SafeArray) error {
	return NewError(E_NOTIMPL)
}

// safeArrayDestroyDescriptor destroys SafeArray object.
//
// AKA: SafeArrayDestroyDescriptor in Windows API.
func safeArrayDestroyDescriptor(safearray *SafeArray) error {
	return NewError(E_NOTIMPL)
}

// safeArrayGetDim is the amount of dimensions in the SafeArray.
//
// SafeArrays may have multiple dimensions. Meaning, it could be
// multidimensional array.
//
// AKA: SafeArrayGetDim in Windows API.
func safeArrayGetDim(safearray *SafeArray) (*uint32, error) {
	u := uint32(0)
	return &u, NewError(E_NOTIMPL)
}

// safeArrayGetElementSize is the element size in bytes.
//
// AKA: SafeArrayGetElemsize in Windows API.
func safeArrayGetElementSize(safearray *SafeArray) (*uint32, error) {
	u := uint32(0)
	return &u, NewError(E_NOTIMPL)
}

// safeArrayGetElement retrieves element at given index.
func safeArrayGetElement(safearray *SafeArray, index int64) (uintptr, error) {
	return uintptr(0), NewError(E_NOTIMPL)
}

// safeArrayGetElement retrieves element at given index and converts to string.
func safeArrayGetElementString(safearray *SafeArray, index int64) (string, error) {
	return "", NewError(E_NOTIMPL)
}

// safeArrayGetIID is the InterfaceID of the elements in the SafeArray.
//
// AKA: SafeArrayGetIID in Windows API.
func safeArrayGetIID(safearray *SafeArray) (*GUID, error) {
	return nil, NewError(E_NOTIMPL)
}

func safeArrayGetLBound(safearray *SafeArray, dimension uint32) (int64, error) {
	return int64(0), NewError(E_NOTIMPL)
}

func safeArrayGetUBound(safearray *SafeArray, dimension uint32) (int64, error) {
	return int64(0), NewError(E_NOTIMPL)
}

func safeArrayGetVartype(safearray *SafeArray) (uint16, error) {
	return uint16(0), NewError(E_NOTIMPL)
}

func safeArrayLock(safearray *SafeArray) error {
	return NewError(E_NOTIMPL)
}

func safeArrayUnlock(safearray *SafeArray) error {
	return NewError(E_NOTIMPL)
}

func safeArrayPutElement(safearray *SafeArray, index int64, element uintptr) error {
	return NewError(E_NOTIMPL)
}

func safeArrayGetRecordInfo(safearray *SafeArray) (interface{}, error) {
	return nil, NewError(E_NOTIMPL)
}

func safeArraySetRecordInfo(safearray *SafeArray, recordInfo interface{}) error {
	return NewError(E_NOTIMPL)
}
