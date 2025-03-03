package ole

import "time"

type VT uint16

const (
	// These do not have values
	VT_EMPTY VT = 0x0
	VT_NULL     = 0x1

	// Numbers
	VT_I2       = 0x2
	VT_I4       = 0x3
	VT_R4       = 0x4
	VT_R8       = 0x5
	VT_I1       = 0x10
	VT_UI1      = 0x11
	VT_UI2      = 0x12
	VT_UI4      = 0x13
	VT_I8       = 0x14
	VT_UI8      = 0x15
	VT_INT      = 0x16
	VT_UINT     = 0x17
	VT_INT_PTR  = 0x25
	VT_UINT_PTR = 0x26

	// VT_CY is a Currency data type which is 8-byte integer with scale of 10,000. $5.25 would be 52500
	VT_CY = 0x6

	// VT_DECIMAL - Please note that this is not a floating point, it is a special structure that allows for Big
	// numbers.
	VT_DECIMAL = 0xe

	// Date
	// The value is a double starting from December 30, 1899
	VT_DATE = 0x7

	// Strings
	VT_BSTR   = 0x8
	VT_LPSTR  = 0x1e
	VT_LPWSTR = 0x1f

	// OLE Automation Objects
	VT_DISPATCH = 0x9
	VT_UNKNOWN  = 0xd

	// Arrays
	VT_SAFEARRAY = 0x1b // Use VT_ARRAY
	VT_CARRAY    = 0x1c
	VT_VECTOR    = 0x1000
	VT_ARRAY     = 0x2000

	// Other types
	VT_ERROR         = 0xa
	VT_BOOL          = 0xb
	VT_VARIANT       = 0xc
	VT_VOID          = 0x18
	VT_HRESULT       = 0x19
	VT_PTR           = 0x1a
	VT_FILETIME      = 0x40
	VT_BLOB          = 0x41
	VT_STORAGE       = 0x43
	VT_STORED_OBJECT = 0x45
	VT_BLOB_OBJECT   = 0x46
	VT_CLSID         = 0x48 // Class IDs are GUIDs.

	// This is not used by its own and must not be used with some other VT values. Not every VT can be by reference or
	// be a pointer.
	VT_BYREF = 0x4000

	//
	// Unsupported types for various reasons.
	//

	// Unsupported, This is a clipboard format
	VT_CF = 0x47

	// Unsupported, do not currently support streams.
	VT_STREAM           = 0x42
	VT_VERSIONED_STREAM = 0x49
	VT_STREAMED_OBJECT  = 0x44

	// Unsupported, because of user-defined types. This COM library is meant to be used as a client and is therefore an
	// edge case. COM servers will likely define the objects and use primitives instead of accepting user-defined types.
	// If you discover a use case and wish to provide the implementation, then please create a fork and submit a pull
	// request.
	VT_USERDEFINED = 0x1d
	VT_RECORD      = 0x24

	// These do not need to be handled.
	VT_RESERVED      = 0x8000
	VT_ILLEGAL       = 0xffff
	VT_ILLEGALMASKED = 0xfff
	VT_TYPEMASK      = 0xfff

	// Reserved for system use.
	//VT_BSTR_BLOB = 0xfff
)

// Date is a Microsoft Date that has an underlying data type of float64 from December 30, 1899. The days are the whole
// number and the floating point within a single day. An hour then is 1 divided by 24 to get 0.04166667 and to get 6am,
// you would multiply 6 by 0.04166667 to get .25. To get minutes, you would then divide 1 by 60 minutes times 24 hours.
// To get seconds, you would multiply by another 60.
//
// You should then be able to multiply hours by DateSingleHour, multiply minutes by DateSingleMinute, and seconds by
// DateSingleSecond to get the fraction.
//
// To get the hours from a date, you would then do the opposite operation. Divide by the below. To get hours, divide
// the fraction by DateSingleHour.
type Date float64

const (
	DateSingleHour        Date = 1 / 24
	DateSingleMinute           = 1 / (24 * 60)
	DateSingleSecond           = 1 / (24 * 60 * 60)
	DateSingleMilliSecond      = 1 / (24 * 60 * 60 * 1000)
)

const (
	// dates taken from microsoft docs
	// https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=net-8.0
	MinOleDate float64 = -657434.0        // Represents January 1, 100
	MaxOleDate float64 = 2958465.99999999 // Represents December 31, 9999
)

var (
	// Technically, The start is local timezone.
	// For simplicity sake, we will convert to UTC. Please be aware that this may mangle displays if this is used.
	// Understand that it is not UTC time, it is local time. UTC is used to make calculations easier but you should not
	// convert the date to local time.
	//
	// NOTE: This hasn't been verified.
	// TODO: Test this by sending dates.
	DateEpoch = time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)

	FileTimeEpoch = time.Date(1601, 1, 1, 0, 0, 0, 0, time.UTC)
)

// Currency data type is an 8-byte integer with scale of 10,000. $5.25 would be 52500
//
// Using this data type to pass into Invoke will automatically convert to a VT_CY VARIANT.
type Currency int64
