package ole

import (
	"fmt"
	"syscall"
	"unicode/utf16"
	"unsafe"
)

type SAFEARRAYBOUND struct {
	CElements uint32
	LLbound   int32
}

type SAFEARRAY struct {
	CDims      uint16
	FFeatures  uint16
	CbElements uint32
	CLocks     uint32
	PvData     uint32
	RgsaBound  SAFEARRAYBOUND
}