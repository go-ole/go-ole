//go:build windows

package safearray

import (
	"github.com/go-ole/go-ole"
	"golang.org/x/sys/windows"
)

// This tests more than one function. It tests all of the functions needed in
// order to retrieve an SafeArray populated with Strings.
func Example_safeArrayGetElementString() {
	ole.Initialize(0)
	defer ole.Uninitialize()

	clsid, err := ole.CLSIDFromProgID("QBXMLRP2.RequestProcessor.1")
	if err != nil {
		if err == windows.CO_E_CLASSSTRING {
			return
		}
	}

	unknown, err := CreateInstance(clsid, IID_IUnknown)
	if err != nil {
		return
	}
	defer unknown.Release()

	dispatch, err := unknown.QueryInterface(IID_IDispatch)
	if err != nil {
		return
	}

	var dispid []int32
	dispid, err = dispatch.GetIDsOfName([]string{"OpenConnection2"})
	if err != nil {
		return
	}

	var result *VARIANT
	_, err = dispatch.Invoke(dispid[0], ole.DISPATCH_METHOD, "", "Test Application 1", 1)
	if err != nil {
		return
	}

	dispid, err = dispatch.GetIDsOfName([]string{"BeginSession"})
	if err != nil {
		return
	}

	result, err = dispatch.Invoke(dispid[0], ole.DISPATCH_METHOD, "", 2)
	if err != nil {
		return
	}

	ticket := result.ToString()

	dispid, err = dispatch.GetIDsOfName([]string{"QBXMLVersionsForSession"})
	if err != nil {
		return
	}

	result, err = dispatch.Invoke(dispid[0], ole.DISPATCH_PROPERTYGET, ticket)
	if err != nil {
		return
	}

	// Where the real tests begin.
	var qbXMLVersions *SafeArray
	var qbXmlVersionStrings []string
	qbXMLVersions = result.ToArray().Array

	// Get array bounds
	var LowerBounds int32
	var UpperBounds int32
	LowerBounds, err = safeArrayGetLBound(qbXMLVersions, 1)
	if err != nil {
		return
	}

	UpperBounds, err = safeArrayGetUBound(qbXMLVersions, 1)
	if err != nil {
		return
	}

	totalElements := UpperBounds - LowerBounds + 1
	qbXmlVersionStrings = make([]string, totalElements)

	for i := int32(0); i < totalElements; i++ {
		qbXmlVersionStrings[i], _ = safeArrayGetElementString(qbXMLVersions, i)
	}

	// Release Safe Array memory
	safeArrayDestroy(qbXMLVersions)

	dispid, err = dispatch.GetIDsOfName([]string{"EndSession"})
	if err != nil {
		return
	}

	_, err = dispatch.Invoke(dispid[0], ole.DISPATCH_METHOD, ticket)
	if err != nil {
		return
	}

	dispid, err = dispatch.GetIDsOfName([]string{"CloseConnection"})
	if err != nil {
		return
	}

	_, err = dispatch.Invoke(dispid[0], ole.DISPATCH_METHOD)
	if err != nil {
		return
	}
}
