// Helper for converting SafeArray to array of objects.

package ole

import (
	"unsafe"
)

type SafeArrayConversion struct {
	Array *SafeArray
}

func (sac *SafeArrayConversion) ToStringArray() (strings []string) {
	totalElements, _ := sac.TotalElements(0)
	strings = make([]string, totalElements)

	for i := int64(0); i < totalElements; i++ {
		strings[int32(i)], _ = safeArrayGetElementString(sac.Array, i)
	}

	return
}

func (sac *SafeArrayConversion) ToByteArray() (bytes []byte) {
	totalElements, _ := sac.TotalElements(0)
	bytes = make([]byte, totalElements)

	for i := int64(0); i < totalElements; i++ {
		safeArrayGetElement(sac.Array, i, unsafe.Pointer(&bytes[int32(i)]))
	}

	return
}

func (sac *SafeArrayConversion) ToValueArray() (values []interface{}) {
	totalElements, _ := sac.TotalElements(0)
	values = make([]interface{}, totalElements)
	vt, _ := safeArrayGetVartype(sac.Array)

	for i := 0; i < int(totalElements); i++ {
		switch VT(vt) {
		case VT_BOOL:
			var v bool
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_I1:
			var v int8
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_I2:
			var v int16
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_I4:
			var v int32
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_I8:
			var v int64
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_UI1:
			var v uint8
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_UI2:
			var v uint16
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_UI4:
			var v uint32
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_UI8:
			var v uint64
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_R4:
			var v float32
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_R8:
			var v float64
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_BSTR:
			var v string
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v
		case VT_VARIANT:
			var v VARIANT
			safeArrayGetElement(sac.Array, int64(i), unsafe.Pointer(&v))
			values[i] = v.Value()
		default:
			// TODO
		}
	}

	return
}

func ToValueArray2(sac *SafeArrayConversion) (values [][]interface{}) {
	totalElements1, _ := sac.TotalElements(1)
	totalElements2, _ := sac.TotalElements(2)
	te1, te2 := int(totalElements1), int(totalElements2)

	values = make([][]interface{}, te1)
	for i := 0; i < te1; i ++ {
		row := make([]interface{}, te2)
		for j := 0; j < te2; j ++ {
			var v VARIANT
			SafeArrayGetElement2(sac.Array, int32(i)+1, int32(j)+1, unsafe.Pointer(&v))
			row[j] = v.Value()
		}
		values[i] = row
	}

	return
}

func (sac *SafeArrayConversion) GetType() (varType uint16, err error) {
	return safeArrayGetVartype(sac.Array)
}

func (sac *SafeArrayConversion) GetDimensions() (dimensions *uint32, err error) {
	return safeArrayGetDim(sac.Array)
}

func (sac *SafeArrayConversion) GetSize() (length *uint32, err error) {
	return safeArrayGetElementSize(sac.Array)
}

func (sac *SafeArrayConversion) TotalElements(index uint32) (totalElements int64, err error) {
	if index < 1 {
		index = 1
	}

	// Get array bounds
	var LowerBounds int64
	var UpperBounds int64

	LowerBounds, err = safeArrayGetLBound(sac.Array, index)
	if err != nil {
		return
	}

	UpperBounds, err = safeArrayGetUBound(sac.Array, index)
	if err != nil {
		return
	}

	totalElements = UpperBounds - LowerBounds + 1
	return
}

// Release Safe Array memory
func (sac *SafeArrayConversion) Release() {
	safeArrayDestroy(sac.Array)
}
