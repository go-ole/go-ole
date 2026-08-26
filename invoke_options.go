package ole

// FreeByRefBSTR controls whether invoke releases the BSTR a server stores in a
// *string out-parameter after the text has been copied into the Go string.
//
// The Automation rules make the caller of IDispatch::Invoke responsible for
// every string referred to by rgvarg, so with FreeByRefBSTR false (the default,
// and the historical behaviour) each [out] BSTR* call leaks one BSTR on the OLE
// task allocator.
//
// It is opt-in because not every server follows the rule: some hand back a BSTR
// they still own (ZKTeco's zkemkeeper is a measured example) and freeing it makes
// every later out-parameter come back empty. Set it to true only for servers
// known to allocate a fresh BSTR for the caller.
var FreeByRefBSTR = false
