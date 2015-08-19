// build +windows

package ole

import "github.com/go-ole/go-ole/oleutil"

func (enum *IEnumVARIANT) Clone() (*IEnumVARIANT, error) {
	r, err := oleutil.CallMethod(enum, "Clone")
	if err != nil {
		return nil, err
	}
	return &IEnumVARIANT{*r.ToIDispatch()}, nil
}

func (enum *IEnumVARIANT) Reset() error {
	_, err := oleutil.CallMethod(enum, "Reset")
	return err
}

func (enum *IEnumVARIANT) Skip(celt int) error {
	_, err := oleutil.CallMethod(enum, "Skip", celt)
	return err
}

func (enum *IEnumVARIANT) Next(celt int, int q) error {
	_, err := oleutil.CallMethod(enum, "Next", celt)
	return err
}
