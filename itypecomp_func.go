// build !windows

package ole

func (v *ITypeComp) Bind(name string, hash uint, flags int16) (*ITypeInfo, int32, error) {
	return nil, 0, NewError(E_NOTIMPL)
}

func (v *ITypeComp) BindType(name string, hash uint) (*ITypeInfo, *ITypeComp, error) {
	return nil, nil, NewError(E_NOTIMPL)
}
