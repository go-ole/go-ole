package legacy

// ClassIDFrom retrieves class ID whether given is program ID or application string.
//
// Helper that provides check against both Class ID from Program ID and Class ID from string. It is
// faster, if you know which you are using, to use the individual functions, but this will check
// against available functions for you.
func ClassIDFrom(programID string) (classID *GUID, err error) {
	classID, err = CLSIDFromProgID(programID)
	if err != nil {
		classID, err = CLSIDFromString(programID)
		if err != nil {
			return
		}
	}
	return
}

// convertHresultToError converts syscall to error, if call is unsuccessful.
func convertHresultToError(hr uintptr, r2 uintptr, ignore error) (err error) {
	if hr != 0 {
		err = NewError(hr)
	}
	return
}
