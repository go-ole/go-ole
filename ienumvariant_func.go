// +build !windows

package ole

func (enum *IEnumVARIANT) Clone() (*IEnumVARIANT, error) {
	return nil, NewError(E_NOTIMPL)
}

func (enum *IEnumVARIANT) Reset() error {
	return NewError(E_NOTIMPL)
}

func (enum *IEnumVARIANT) Skip(celt int) error {
	return NewError(E_NOTIMPL)
}

func (enum *IEnumVARIANT) Next(celt int, int q) error {
	return NewError(E_NOTIMPL)
}
