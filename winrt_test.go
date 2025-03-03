//go:build windows

package ole

import (
	"testing"
)

func TestWinRT_XMLDocument(t *testing.T) {
	// IXmlDocumentIO is ABI.Windows.Data.Xml.Dom.IXmlDocumentIO
	IXmlDocumentIO, err := windows.GUIDFromString("{6cd0e74e-ee65-4489-9ebf-ca43e87ba637}")
	if err != nil {
		t.Error(err)
		return
	}

	RoInitialize(RoMultithreaded)
	defer RoUninitialize()

	inspectable, err := RoActivateInstance("Windows.Data.Xml.Dom.XmlDocument")
	if err != nil {
		t.Error(err)
		return
	}
	defer inspectable.Release()

	xmldoc, err := QueryInterfaceOnIUnknown[IDispatch](inspectable, IXmlDocumentIO)
	if err != nil {
		t.Error(err)
		return
	}
	defer xmldoc.Release()

	hString, err := NewHString("<test></test>")
	if err != nil {
		t.Error(err)
		return
	}
	defer DeleteHString(hString)

	// panics with "unknown type"
	xmldoc.CallMethod("LoadXml", HStringToVariant(hString))
}
