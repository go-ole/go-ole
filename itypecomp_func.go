// build !windows

package ole

func (v *ITypeComp) Bind(name string, hash uint, flags short) (*ITypeInfo, int, error) {
	return nil, 0, NewError(E_NOTIMPL)
}

func (v *ITypeComp) BindType(name string, hash uint) (*ITypeInfo, *ITypeComp, error) {
	return nil, nil, NewError(E_NOTIMPL)
}
