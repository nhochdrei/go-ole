// Helper for converting SafeArray to array of objects.

package ole

import (
	"fmt"
	"unsafe"
)

type SafeArrayConversion struct {
	Array *SafeArray
}

func (sac *SafeArrayConversion) ToStringArray() (strings []string) {
	totalElements, _ := sac.TotalElements(0)
	strings = make([]string, totalElements)

	for i := int32(0); i < totalElements; i++ {
		strings[int32(i)], _ = safeArrayGetElementString(sac.Array, i)
	}

	return
}

func (sac *SafeArrayConversion) ToByteArray() (bytes []byte) {
	totalElements, _ := sac.TotalElements(0)
	bytes = make([]byte, totalElements)

	for i := int32(0); i < totalElements; i++ {
		safeArrayGetElement(sac.Array, i, unsafe.Pointer(&bytes[int32(i)]))
	}

	return
}

const sizeOfUintPtr = unsafe.Sizeof(uintptr(0))

func uintptrToBytes(u *uintptr) []byte {
	return (*[sizeOfUintPtr]byte)(unsafe.Pointer(u))[:]
}

func printValueAtMemoryLocation(location uintptr, next int) {
	var v byte
	p := unsafe.Pointer(location)
	fmt.Println("8 Bit Memory \n")
	for i := 1; i < next; i++ {
		p = unsafe.Pointer(location)
		v = *((*byte)(p))
		fmt.Print(v, " ")
		//fmt.Println("Loc : ", loc, " --- Val : ", v)
		location++
	}
	fmt.Println("\n")
}

func printValueAtMemoryLocation16(location uintptr, next int) {
	var v int16
	p := unsafe.Pointer(location)
	fmt.Println("16 Bit Memory \n")
	for i := 1; i < next; i++ {
		p = unsafe.Pointer(location)
		v = *((*int16)(p))
		fmt.Print(v, " ")
		//fmt.Println("Loc : ", loc, " --- Val : ", v)
		location += 2
	}
	fmt.Println("\n")
}
func printValueAtMemoryLocation32(location uintptr, next int) {
	var v int32
	p := unsafe.Pointer(location)
	fmt.Println("32 Bit Memory \n")
	for i := 1; i < next; i++ {
		p = unsafe.Pointer(location)
		v = *((*int32)(p))
		fmt.Print(v, " ")
		//fmt.Println("Loc : ", loc, " --- Val : ", v)
		location += 4
	}
	fmt.Println("\n")
}

func (sac *SafeArrayConversion) GetData() ([]byte, error) {
	bytes := make([]byte, 200)
	data, err := safeArrayAccessData(sac.Array)
	if err != nil {
		return nil, err
	}

	// d := (*byte)(unsafe.Pointer(data))

	// fmt.Println(data, uintptrToBytes(&data))
	printValueAtMemoryLocation(data, 200)
	printValueAtMemoryLocation16(data, 100)
	printValueAtMemoryLocation32(data, 50)

	err = safeArrayUnaccessData(sac.Array)
	if err != nil {
		return nil, err
	}
	return bytes, nil
}

func (sac *SafeArrayConversion) GetVarType() VT {
	vt, _ := safeArrayGetVartype(sac.Array)
	return VT(vt)
}

func (sac *SafeArrayConversion) ToValueArray() (values []interface{}) {
	vt, _ := safeArrayGetVartype(sac.Array)
	return sac.ToValueArrayWithType(VT(vt))
}

func (sac *SafeArrayConversion) ToValueArrayWithType(vt VT) (values []interface{}) {
	totalElements, _ := sac.TotalElements(0)
	values = make([]interface{}, totalElements)

	for i := int32(0); i < totalElements; i++ {
		switch VT(vt) {
		case VT_BOOL:
			var v bool
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_I1:
			var v int8
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_I2:
			var v int16
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_I4:
			var v int32
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_I8:
			var v int64
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_UI1:
			var v uint8
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_UI2:
			var v uint16
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_UI4:
			var v uint32
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_UI8:
			var v uint64
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_R4:
			var v float32
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_R8:
			var v float64
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_BSTR:
			v, _ := safeArrayGetElementString(sac.Array, i)
			values[i] = v
		case VT_VARIANT:
			var v VARIANT
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v.Value()
			v.Clear()
		default:
			// TODO
		}
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

func (sac *SafeArrayConversion) TotalElements(index uint32) (totalElements int32, err error) {
	if index < 1 {
		index = 1
	}

	// Get array bounds
	var LowerBounds int32
	var UpperBounds int32

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
