package ole

type VT uint16

const (
	VT_EMPTY            VT = 0x0
	VT_NULL                = 0x1
	VT_I2                  = 0x2
	VT_I4                  = 0x3
	VT_R4                  = 0x4
	VT_R8                  = 0x5
	VT_CY                  = 0x6
	VT_DATE                = 0x7
	VT_BSTR                = 0x8
	VT_DISPATCH            = 0x9
	VT_ERROR               = 0xa
	VT_BOOL                = 0xb
	VT_VARIANT             = 0xc
	VT_UNKNOWN             = 0xd
	VT_DECIMAL             = 0xe
	VT_I1                  = 0x10
	VT_UI1                 = 0x11
	VT_UI2                 = 0x12
	VT_UI4                 = 0x13
	VT_I8                  = 0x14
	VT_UI8                 = 0x15
	VT_INT                 = 0x16
	VT_UINT                = 0x17
	VT_VOID                = 0x18
	VT_HRESULT             = 0x19
	VT_PTR                 = 0x1a
	VT_SAFEARRAY           = 0x1b
	VT_CARRAY              = 0x1c
	VT_USERDEFINED         = 0x1d
	VT_LPSTR               = 0x1e
	VT_LPWSTR              = 0x1f
	VT_RECORD              = 0x24
	VT_INT_PTR             = 0x25
	VT_UINT_PTR            = 0x26
	VT_FILETIME            = 0x40
	VT_BLOB                = 0x41
	VT_STREAM              = 0x42
	VT_STORAGE             = 0x43
	VT_STREAMED_OBJECT     = 0x44
	VT_STORED_OBJECT       = 0x45
	VT_BLOB_OBJECT         = 0x46
	VT_CF                  = 0x47
	VT_CLSID               = 0x48
	VT_VERSIONED_STREAM    = 0x49
	// Reserved for system use.
	VT_BSTR_BLOB     = 0xfff
	VT_VECTOR        = 0x1000
	VT_ARRAY         = 0x2000
	VT_BYREF         = 0x4000
	VT_RESERVED      = 0x8000
	VT_ILLEGAL       = 0xffff
	VT_ILLEGALMASKED = 0xfff
	VT_TYPEMASK      = 0xfff
)
