//go:build windows
// +build windows

package ole

import (
	"unsafe"
)

func safeArrayFromByteSlice(slice []byte) *SafeArray {
	array, _ := safeArrayCreateVector(VT_UI1, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []byte to SAFEARRAY")
	}

	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return array
}

func safeArrayFromInt16Slice(slice []int16) *SafeArray {
	array, _ := safeArrayCreateVector(VT_I2, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []int16 to SAFEARRAY")
	}

	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return array
}

func safeArrayFromUInt16Slice(slice []uint16) *SafeArray {
	array, _ := safeArrayCreateVector(VT_UI2, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []uint16 to SAFEARRAY")
	}

	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return array
}

func safeArrayFromInt32Slice(slice []int32) *SafeArray {
	array, _ := safeArrayCreateVector(VT_I4, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []int32 to SAFEARRAY")
	}

	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return array
}

func safeArrayFromUInt32Slice(slice []uint32) *SafeArray {
	array, _ := safeArrayCreateVector(VT_UI4, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []uint32 to SAFEARRAY")
	}

	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return array
}

func safeArrayFromInt64Slice(slice []int64) *SafeArray {
	array, _ := safeArrayCreateVector(VT_I8, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []int64 to SAFEARRAY")
	}

	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return array
}

func safeArrayFromUInt64Slice(slice []uint64) *SafeArray {
	array, _ := safeArrayCreateVector(VT_UI8, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []uint64 to SAFEARRAY")
	}

	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return array
}

func safeArrayFromStringSlice(slice []string) *SafeArray {
	array, _ := safeArrayCreateVector(VT_BSTR, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []string to SAFEARRAY")
	}
	// SysAllocStringLen(s)
	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(SysAllocStringLen(v))))
	}
	return array
}
