//go:build windows

package ole

import (
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

type MEMBERID int32
type HREFTYPE uint32

type ITypeComp struct {
	VirtualTable *ITypeCompVirtualTable
}

type ITypeCompVirtualTable struct {
	// IUnknown
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
	// ITypeComp
	Bind            uintptr
	BindType        uintptr
}

func (obj *ITypeComp) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *ITypeComp) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

func (obj *ITypeComp) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

func (obj *ITypeComp) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *ITypeComp) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

type ITypeLib struct {
	VirtualTable *ITypeLibVirtualTable
}

type ITypeLibVirtualTable struct {
	// IUnknown
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
	// ITypeLib
	GetTypeInfoCount    uintptr
	GetTypeInfo         uintptr
	GetTypeInfoType     uintptr
	GetTypeInfoOfGuid   uintptr
	GetLibAttr          uintptr
	GetTypeComp         uintptr
	GetDocumentation    uintptr
	IsName              uintptr
	FindName            uintptr
	ReleaseTLibAttr     uintptr
}

func (obj *ITypeLib) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *ITypeLib) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

func (obj *ITypeLib) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

func (obj *ITypeLib) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *ITypeLib) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

type PARAMDESC struct {
	Pptrparam uintptr
	WParamFlags uint16
}

type ELEMDESC struct {
	Tdesc TYPEDESC
	Paramdesc PARAMDESC
}

type FUNCDESC struct {
	Memid             MEMBERID
	Lprgscode         *int32
	LprgelemdescParam *ELEMDESC
	Funckind          int32
	Invkind           int32
	Callconv          int32
	CParams           int16
	CParamsOpt        int16
	OVft              int16
	CScodes           int16
	ElemdescFunc      ELEMDESC
	WFuncFlags        uint16
}

type VARDESC struct {
	Memid       MEMBERID
	LpstrSchema *uint16
	LpvarValue  *VARIANT
	ElemdescVar ELEMDESC
	WVarFlags   uint16
	Varkind     int32
}

type ITypeInfoAddresses interface {
	IsIUnknown
	GetTypeAttrAddress() uintptr
	GetTypeCompAddress() uintptr
	GetFuncDescAddress() uintptr
	GetVarDescAddress() uintptr
	GetNamesAddress() uintptr
	GetRefTypeOfImplTypeAddress() uintptr
	GetImplTypeFlagsAddress() uintptr
	GetIDsOfNamesAddress() uintptr
	InvokeAddress() uintptr
	GetDocumentationAddress() uintptr
	GetDllEntryAddress() uintptr
	GetRefTypeInfoAddress() uintptr
	AddressOfMemberAddress() uintptr
	CreateInstanceAddress() uintptr
	GetMopsAddress() uintptr
	GetContainingTypeLibAddress() uintptr
	ReleaseTypeAttrAddress() uintptr
	ReleaseFuncDescAddress() uintptr
	ReleaseVarDescAddress() uintptr
}

type ITypeInfo struct {
	VirtualTable *ITypeInfoVirtualTable
}

type ITypeInfoVirtualTable struct {
	// IUnknown
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
	// ITypeInfo
	GetTypeAttr          uintptr
	GetTypeComp          uintptr
	GetFuncDesc          uintptr
	GetVarDesc           uintptr
	GetNames             uintptr
	GetRefTypeOfImplType uintptr
	GetImplTypeFlags     uintptr
	GetIDsOfNames        uintptr
	Invoke               uintptr
	GetDocumentation     uintptr
	GetDllEntry          uintptr
	GetRefTypeInfo       uintptr
	AddressOfMember      uintptr
	CreateInstance       uintptr
	GetMops              uintptr
	GetContainingTypeLib uintptr
	ReleaseTypeAttr      uintptr
	ReleaseFuncDesc      uintptr
	ReleaseVarDesc       uintptr
}

func (obj *ITypeInfo) QueryInterfaceAddress() uintptr {
	return obj.VirtualTable.QueryInterface
}

func (obj *ITypeInfo) AddRefAddress() uintptr {
	return obj.VirtualTable.AddRef
}

func (obj *ITypeInfo) ReleaseAddress() uintptr {
	return obj.VirtualTable.Release
}

func (obj *ITypeInfo) AddRef() uint32 {
	return AddRefOnIUnknown(obj)
}

func (obj *ITypeInfo) Release() uint32 {
	return ReleaseOnIUnknown(obj)
}

func (obj *ITypeInfo) GetTypeAttrAddress() uintptr {
	return obj.VirtualTable.GetTypeAttr
}

func (obj *ITypeInfo) GetTypeCompAddress() uintptr {
	return obj.VirtualTable.GetTypeComp
}

func (obj *ITypeInfo) GetFuncDescAddress() uintptr {
	return obj.VirtualTable.GetFuncDesc
}

func (obj *ITypeInfo) GetVarDescAddress() uintptr {
	return obj.VirtualTable.GetVarDesc
}

func (obj *ITypeInfo) GetNamesAddress() uintptr {
	return obj.VirtualTable.GetNames
}

func (obj *ITypeInfo) GetRefTypeOfImplTypeAddress() uintptr {
	return obj.VirtualTable.GetRefTypeOfImplType
}

func (obj *ITypeInfo) GetImplTypeFlagsAddress() uintptr {
	return obj.VirtualTable.GetImplTypeFlags
}

func (obj *ITypeInfo) GetIDsOfNamesAddress() uintptr {
	return obj.VirtualTable.GetIDsOfNames
}

func (obj *ITypeInfo) InvokeAddress() uintptr {
	return obj.VirtualTable.Invoke
}

func (obj *ITypeInfo) GetDocumentationAddress() uintptr {
	return obj.VirtualTable.GetDocumentation
}

func (obj *ITypeInfo) GetDllEntryAddress() uintptr {
	return obj.VirtualTable.GetDllEntry
}

func (obj *ITypeInfo) GetRefTypeInfoAddress() uintptr {
	return obj.VirtualTable.GetRefTypeInfo
}

func (obj *ITypeInfo) AddressOfMemberAddress() uintptr {
	return obj.VirtualTable.AddressOfMember
}

func (obj *ITypeInfo) CreateInstanceAddress() uintptr {
	return obj.VirtualTable.CreateInstance
}

func (obj *ITypeInfo) GetMopsAddress() uintptr {
	return obj.VirtualTable.GetMops
}

func (obj *ITypeInfo) GetContainingTypeLibAddress() uintptr {
	return obj.VirtualTable.GetContainingTypeLib
}

func (obj *ITypeInfo) ReleaseTypeAttrAddress() uintptr {
	return obj.VirtualTable.ReleaseTypeAttr
}

func (obj *ITypeInfo) ReleaseFuncDescAddress() uintptr {
	return obj.VirtualTable.ReleaseFuncDesc
}

func (obj *ITypeInfo) ReleaseVarDescAddress() uintptr {
	return obj.VirtualTable.ReleaseVarDesc
}

// GetTypeComp retrieves the ITypeComp interface for the type description.
func (obj *ITypeInfo) GetTypeComp() (tcomp *ITypeComp, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetTypeComp,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&tcomp)),
		0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// GetFuncDesc retrieves the FUNCDESC structure that contains information about a specified function.
func (obj *ITypeInfo) GetFuncDesc(index uint32) (funcdesc *FUNCDESC, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetFuncDesc,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(index),
		uintptr(unsafe.Pointer(&funcdesc)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// GetVarDesc retrieves a VARDESC structure that describes the specified variable.
func (obj *ITypeInfo) GetVarDesc(index uint32) (vardesc *VARDESC, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetVarDesc,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(index),
		uintptr(unsafe.Pointer(&vardesc)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// GetNames retrieves the variable with the specified member ID or the name of the property or method and its parameters.
func (obj *ITypeInfo) GetNames(memid MEMBERID, names []*uint16, maxNames uint32) (count uint32, err error) {
	hr, _, _ := syscall.Syscall6(
		obj.VirtualTable.GetNames,
		5,
		uintptr(unsafe.Pointer(obj)),
		uintptr(memid),
		uintptr(unsafe.Pointer(&names[0])),
		uintptr(maxNames),
		uintptr(unsafe.Pointer(&count)),
		0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// GetRefTypeOfImplType retrieves the type description of the implemented interface types if a type description describes a COM class.
func (obj *ITypeInfo) GetRefTypeOfImplType(index uint32) (reftype HREFTYPE, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetRefTypeOfImplType,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(index),
		uintptr(unsafe.Pointer(&reftype)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// GetImplTypeFlags retrieves the IMPLTYPEFLAGS enumeration for one implemented interface or base interface in a type description.
func (obj *ITypeInfo) GetImplTypeFlags(index uint32) (flags int32, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetImplTypeFlags,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(index),
		uintptr(unsafe.Pointer(&flags)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// GetIDsOfNames maps between member names and member IDs, and parameter names and parameter IDs.
func (obj *ITypeInfo) GetIDsOfNames(names []*uint16, count uint32) (memids []MEMBERID, err error) {
	memids = make([]MEMBERID, count)
	hr, _, _ := syscall.Syscall6(
		obj.VirtualTable.GetIDsOfNames,
		4,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&names[0])),
		uintptr(count),
		uintptr(unsafe.Pointer(&memids[0])),
		0, 0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// Invoke invokes a method, or accesses a property of an object, that implements the interface described by the type description.
func (obj *ITypeInfo) Invoke(instance uintptr, memid MEMBERID, flags uint16, params *DISPPARAMS, result *VARIANT, excepInfo *EXCEPINFO, argErr *uint32) (err error) {
	hr, _, _ := syscall.Syscall9(
		obj.VirtualTable.Invoke,
		8,
		uintptr(unsafe.Pointer(obj)),
		instance,
		uintptr(memid),
		uintptr(flags),
		uintptr(unsafe.Pointer(params)),
		uintptr(unsafe.Pointer(result)),
		uintptr(unsafe.Pointer(excepInfo)),
		uintptr(unsafe.Pointer(argErr)),
		0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// GetDocumentation retrieves the documentation string, the complete Help file name and path, and the context ID for the Help topic for a specified type description.
func (obj *ITypeInfo) GetDocumentation(memid MEMBERID) (name, docString, helpFile string, helpContext uint32, err error) {
	var bstrName, bstrDocString, bstrHelpFile *uint16
	hr, _, _ := syscall.Syscall6(
		obj.VirtualTable.GetDocumentation,
		6,
		uintptr(unsafe.Pointer(obj)),
		uintptr(memid),
		uintptr(unsafe.Pointer(&bstrName)),
		uintptr(unsafe.Pointer(&bstrDocString)),
		uintptr(unsafe.Pointer(&helpContext)),
		uintptr(unsafe.Pointer(&bstrHelpFile)))
	if hr != 0 {
		err = windows.Errno(hr)
		return
	}
	if bstrName != nil {
		name = windows.UTF16PtrToString(bstrName)
		SysFreeString(bstrName)
	}
	if bstrDocString != nil {
		docString = windows.UTF16PtrToString(bstrDocString)
		SysFreeString(bstrDocString)
	}
	if bstrHelpFile != nil {
		helpFile = windows.UTF16PtrToString(bstrHelpFile)
		SysFreeString(bstrHelpFile)
	}
	return
}

// GetDllEntry retrieves a description or specification of an entry point for a function in a DLL.
func (obj *ITypeInfo) GetDllEntry(memid MEMBERID, invkind int32) (dllName, name string, entry uint16, err error) {
	var bstrDllName, bstrName *uint16
	var wOrdinal uint16
	hr, _, _ := syscall.Syscall6(
		obj.VirtualTable.GetDllEntry,
		6,
		uintptr(unsafe.Pointer(obj)),
		uintptr(memid),
		uintptr(invkind),
		uintptr(unsafe.Pointer(&bstrDllName)),
		uintptr(unsafe.Pointer(&bstrName)),
		uintptr(unsafe.Pointer(&wOrdinal)))
	if hr != 0 {
		err = windows.Errno(hr)
		return
	}
	if bstrDllName != nil {
		dllName = windows.UTF16PtrToString(bstrDllName)
		SysFreeString(bstrDllName)
	}
	if bstrName != nil {
		name = windows.UTF16PtrToString(bstrName)
		SysFreeString(bstrName)
	}
	entry = wOrdinal
	return
}

// GetRefTypeInfo retrieves the type description that is referenced by other type descriptions.
func (obj *ITypeInfo) GetRefTypeInfo(reftype HREFTYPE) (typeInfo *ITypeInfo, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetRefTypeInfo,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(reftype),
		uintptr(unsafe.Pointer(&typeInfo)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// AddressOfMember retrieves the addresses of static functions or variables, such as those defined in a DLL.
func (obj *ITypeInfo) AddressOfMember(memid MEMBERID, invkind int32) (address uintptr, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.AddressOfMember,
		4,
		uintptr(unsafe.Pointer(obj)),
		uintptr(memid),
		uintptr(invkind),
		uintptr(unsafe.Pointer(&address)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// CreateInstance creates a new instance of a type that describes a component object class (coclass).
func (obj *ITypeInfo) CreateInstance(outer *IUnknown, riid *windows.GUID) (instance uintptr, err error) {
	hr, _, _ := syscall.Syscall6(
		obj.VirtualTable.CreateInstance,
		4,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(outer)),
		uintptr(unsafe.Pointer(riid)),
		uintptr(unsafe.Pointer(&instance)),
		0, 0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// GetMops retrieves marshaling information.
func (obj *ITypeInfo) GetMops(memid MEMBERID) (mops string, err error) {
	var bstrMops *uint16
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetMops,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(memid),
		uintptr(unsafe.Pointer(&bstrMops)))
	if hr != 0 {
		err = windows.Errno(hr)
		return
	}
	if bstrMops != nil {
		mops = windows.UTF16PtrToString(bstrMops)
		SysFreeString(bstrMops)
	}
	return
}

// GetContainingTypeLib retrieves the containing type library and the index of the type description within that type library.
func (obj *ITypeInfo) GetContainingTypeLib() (typeLib *ITypeLib, index uint32, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetContainingTypeLib,
		3,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&typeLib)),
		uintptr(unsafe.Pointer(&index)))
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}

// ReleaseTypeAttr releases a TYPEATTR previously returned by GetTypeAttr.
func (obj *ITypeInfo) ReleaseTypeAttr(typeAttr *TYPEATTR) {
	syscall.Syscall(
		obj.VirtualTable.ReleaseTypeAttr,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(typeAttr)),
		0)
}

// ReleaseFuncDesc releases a FUNCDESC previously returned by GetFuncDesc.
func (obj *ITypeInfo) ReleaseFuncDesc(funcDesc *FUNCDESC) {
	syscall.Syscall(
		obj.VirtualTable.ReleaseFuncDesc,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(funcDesc)),
		0)
}

// ReleaseVarDesc releases a VARDESC previously returned by GetVarDesc.
func (obj *ITypeInfo) ReleaseVarDesc(varDesc *VARDESC) {
	syscall.Syscall(
		obj.VirtualTable.ReleaseVarDesc,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(varDesc)),
		0)
}

// GetTypeAttr retrieves a TYPEATTR structure that contains the attributes of the type description.
func (obj *ITypeInfo) GetTypeAttr() (tattr *TYPEATTR, err error) {
	hr, _, _ := syscall.Syscall(
		obj.VirtualTable.GetTypeAttr,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&tattr)),
		0)
	if hr != 0 {
		err = windows.Errno(hr)
	}
	return
}
