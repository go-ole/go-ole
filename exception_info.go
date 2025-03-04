//go:build windows

package ole

import (
	"fmt"
	"strings"
	"unsafe"
)

// EXCEPINFO defines exception info.
type EXCEPINFO struct {
	wCode             uint16
	wReserved         uint16
	bstrSource        *uint16
	bstrDescription   *uint16
	bstrHelpFile      *uint16
	dwHelpContext     uint32
	pvReserved        uintptr
	pfnDeferredFillIn uintptr
	sCode             uint32

	// Go-specific part. Don't move upper cos it'll break structure layout for native code.
	rendered    bool
	source      string
	description string
	helpFile    string
}

// renderStrings translates BSTR strings to Go ones so `.Error` and `.String`
// could be safely called after `.Clear`. We need this when we can't rely on
// a caller to call `.Clear`.
func (e *EXCEPINFO) renderStrings() {
	e.rendered = true
	if e.bstrSource == nil {
		e.source = "<nil>"
	} else {
		e.source = BstrToString(e.bstrSource)
	}
	if e.bstrDescription == nil {
		e.description = "<nil>"
	} else {
		e.description = BstrToString(e.bstrDescription)
	}
	if e.bstrHelpFile == nil {
		e.helpFile = "<nil>"
	} else {
		e.helpFile = BstrToString(e.bstrHelpFile)
	}
}

// Clear frees BSTR strings inside an EXCEPINFO and set it to NULL.
func (e *EXCEPINFO) Clear() {
	freeBSTR := func(s *uint16) {
		// SysFreeString don't return errors and is safe for call's on NULL.
		// https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysfreestring
		_ = SysFreeString((*uint16)(unsafe.Pointer(s)))
	}

	if e.bstrSource != nil {
		freeBSTR(e.bstrSource)
		e.bstrSource = nil
	}
	if e.bstrDescription != nil {
		freeBSTR(e.bstrDescription)
		e.bstrDescription = nil
	}
	if e.bstrHelpFile != nil {
		freeBSTR(e.bstrHelpFile)
		e.bstrHelpFile = nil
	}
}

// WCode return wCode in EXCEPINFO.
func (e EXCEPINFO) WCode() uint16 {
	return e.wCode
}

// SCode return sCode in EXCEPINFO.
func (e EXCEPINFO) SCode() uint32 {
	return e.sCode
}

// String convert EXCEPINFO to string.
func (e EXCEPINFO) String() string {
	if !e.rendered {
		e.renderStrings()
	}
	return fmt.Sprintf(
		"wCode: %#x, bstrSource: %v, bstrDescription: %v, bstrHelpFile: %v, dwHelpContext: %#x, scode: %#x",
		e.wCode, e.source, e.description, e.helpFile, e.dwHelpContext, e.sCode,
	)
}

// Error implements error interface and returns error string.
func (e EXCEPINFO) Error() string {
	if !e.rendered {
		e.renderStrings()
	}

	if e.description != "<nil>" {
		return strings.TrimSpace(e.description)
	}

	code := e.sCode
	if e.wCode != 0 {
		code = uint32(e.wCode)
	}
	return fmt.Sprintf("%v: %#x", e.source, code)
}
