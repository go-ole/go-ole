//go:build windows

// Package is meant to retrieve and process safe array data returned from COM.

package safearray

// Safe Array Feature Flags

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

// SafeArrayBound defines the SafeArray boundaries.
type SafeArrayBound struct {
	Elements   uint32
	LowerBound int32
}

// SafeArray is how COM handles arrays.
type SafeArray struct {
	Dimensions   uint16
	FeaturesFlag uint16
	ElementsSize uint32
	LocksAmount  uint32
	Data         uint32
	Bounds       [16]byte
}
