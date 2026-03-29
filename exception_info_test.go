//go:build windows

package ole

import (
	"fmt"
	"testing"
)

func TestEXCEPINFOStringRendersBSTRFields(t *testing.T) {
	excepInfo := EXCEPINFO{
		wCode:           0x1234,
		bstrSource:      SysAllocString("UnitTest.Source"),
		bstrDescription: SysAllocString("something happened"),
		bstrHelpFile:    SysAllocString("help.chm"),
		dwHelpContext:   0x4321,
		sCode:           0x80004005,
	}
	defer excepInfo.Clear()

	want := "wCode: 0x1234, bstrSource: UnitTest.Source, bstrDescription: something happened, bstrHelpFile: help.chm, dwHelpContext: 0x4321, scode: 0x80004005"

	if got := excepInfo.String(); got != want {
		t.Fatalf("EXCEPINFO.String() = %q, want %q", got, want)
	}
}

func TestEXCEPINFOErrorUsesTrimmedDescription(t *testing.T) {
	excepInfo := EXCEPINFO{
		bstrSource:      SysAllocString("UnitTest.Source"),
		bstrDescription: SysAllocString("  detailed message \r\n"),
		sCode:           0x80004005,
	}
	defer excepInfo.Clear()

	if got := excepInfo.Error(); got != "detailed message" {
		t.Fatalf("EXCEPINFO.Error() = %q, want %q", got, "detailed message")
	}
}

func TestEXCEPINFOErrorUsesWCodeWhenDescriptionMissing(t *testing.T) {
	excepInfo := EXCEPINFO{
		wCode:      0x1234,
		bstrSource: SysAllocString("UnitTest.Source"),
		sCode:      0x80004005,
	}
	defer excepInfo.Clear()

	want := "UnitTest.Source: 0x1234"
	if got := excepInfo.Error(); got != want {
		t.Fatalf("EXCEPINFO.Error() = %q, want %q", got, want)
	}
}

func TestEXCEPINFOErrorUsesSCodeWhenWCodeMissing(t *testing.T) {
	excepInfo := EXCEPINFO{
		bstrSource: SysAllocString("UnitTest.Source"),
		sCode:      0x80004005,
	}
	defer excepInfo.Clear()

	want := "UnitTest.Source: 0x80004005"
	if got := excepInfo.Error(); got != want {
		t.Fatalf("EXCEPINFO.Error() = %q, want %q", got, want)
	}
}

func TestEXCEPINFOClearPreservesRenderedStrings(t *testing.T) {
	excepInfo := EXCEPINFO{
		wCode:           0x1234,
		bstrSource:      SysAllocString("UnitTest.Source"),
		bstrDescription: SysAllocString("message"),
		bstrHelpFile:    SysAllocString("help.chm"),
		dwHelpContext:   0x4321,
		sCode:           0x80004005,
	}

	wantString := excepInfo.String()
	wantError := excepInfo.Error()

	excepInfo.Clear()

	if excepInfo.bstrSource != nil || excepInfo.bstrDescription != nil || excepInfo.bstrHelpFile != nil {
		t.Fatal("EXCEPINFO.Clear() did not nil all BSTR fields")
	}

	if got := excepInfo.String(); got != wantString {
		t.Fatalf("EXCEPINFO.String() after Clear() = %q, want %q", got, wantString)
	}

	if got := excepInfo.Error(); got != wantError {
		t.Fatalf("EXCEPINFO.Error() after Clear() = %q, want %q", got, wantError)
	}
}

func TestEXCEPINFOStringWithNilBSTRs(t *testing.T) {
	excepInfo := EXCEPINFO{
		wCode:         0x1234,
		dwHelpContext: 0x4321,
		sCode:         0x80004005,
	}

	want := fmt.Sprintf(
		"wCode: %#x, bstrSource: <nil>, bstrDescription: <nil>, bstrHelpFile: <nil>, dwHelpContext: %#x, scode: %#x",
		excepInfo.wCode, excepInfo.dwHelpContext, excepInfo.sCode,
	)

	if got := excepInfo.String(); got != want {
		t.Fatalf("EXCEPINFO.String() = %q, want %q", got, want)
	}
}
