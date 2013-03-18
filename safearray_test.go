package ole

import (
	"testing"
)

// This tests more than one function. It tests all of the functions needed in order to retrieve an
// SafeArray populated with Strings.
func TestGetSafeArrayString(t *testing.T) {
	CoInitialize(0)
	defer CoUninitialize()

	clsid, err := CLSIDFromProgID("QBXMLRP2.RequestProcessor.1")
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	unknown, err := ole.CreateInstance(clsid, IID_IUnknown)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}
	defer unknown.Release()

	dispatch, err := unknown.QueryInterface(IID_IDispatch)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	var dispid []int32
	dispid, err = dispatch.GetIDsOfName([]string{"OpenConnection2"})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	var result *VARIANT
	_, err = d.dispatch.Invoke(dispid[0], DISPATCH_METHOD, "", "Test Application 1", 1)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	dispid, err = dispatch.GetIDsOfName([]string{"BeginSession"})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	result, err = d.dispatch.Invoke(dispid[0], DISPATCH_METHOD, "", 2)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	ticket := result.ToString()

	dispid, err = dispatch.GetIDsOfName([]string{"QBXMLVersionsForSession"})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	result, err = d.dispatch.Invoke(dispid[0], DISPATCH_METHOD, ticket)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}
	
	// Where the real tests begin.
	var qbXMLVersions *SafeArray
	qbXMLVersions = result.ToSafeArray()

	// Increment reference count
	safeArrayDataPointer, err := safeArrayAccessData(qbXMLVersions)
	if err != nil {
		t.Log("Safe Array Access Data")
		t.Log(err)
		t.FailNow()
	}
	
	// Get array bounds
	var LowerBounds int64
	var UpperBounds int64
	LowerBounds, err = safeArrayGetLBound(qbXMLVersions, 1)
	if err != nil {
		t.Log("Safe Array Get Lower Bound")
		t.Log(err)
		t.FailNow()
	}
	t.Log("Lower Bounds:")
	t.Log(LowerBounds)
	
	UpperBounds, err = safeArrayGetUBound(qbXMLVersions, 1)
	if err != nil {
		t.Log("Safe Array Get Lower Bound")
		t.Log(err)
		t.FailNow()
	}
	t.Log("Upper Bounds:")
	t.Log(UpperBounds)
	
	// Decrement reference count
	safeArrayUnaccessData(qbXMLVersions)
	
	// Release Safe Array memory
	safeArrayDestroy(qbXMLVersions)

	dispid, err = dispatch.GetIDsOfName([]string{"EndSession"})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	_, err = d.dispatch.Invoke(dispid[0], DISPATCH_METHOD, ticket)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	dispid, err = dispatch.GetIDsOfName([]string{"CloseConnection"})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	_, err = d.dispatch.Invoke(dispid[0], DISPATCH_METHOD)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}
}

