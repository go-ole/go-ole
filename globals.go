package ole

import (
	"golang.org/x/sys/windows"
)

var (
	modcombase  = windows.NewLazySystemDLL("combase.dll")
	modkernel32 = windows.NewLazySystemDLL("kernel32.dll")
	modole32    = windows.NewLazySystemDLL("ole32.dll")
	modoleaut32 = windows.NewLazySystemDLL("oleaut32.dll")
	moduser32   = windows.NewLazySystemDLL("user32.dll")
)

var (
	procCoCreateInstance = modole32.NewProc("CoCreateInstance")
	procCoGetObject      = modole32.NewProc("CoGetObject")

	procCopyMemory              = modkernel32.NewProc("RtlMoveMemory")
	procVariantInit             = modoleaut32.NewProc("VariantInit")
	procVariantClear            = modoleaut32.NewProc("VariantClear")
	procVariantTimeToSystemTime = modoleaut32.NewProc("VariantTimeToSystemTime")
	procSysAllocString          = modoleaut32.NewProc("SysAllocString")
	procSysAllocStringLen       = modoleaut32.NewProc("SysAllocStringLen")
	procSysFreeString           = modoleaut32.NewProc("SysFreeString")
	procCreateDispTypeInfo      = modoleaut32.NewProc("CreateDispTypeInfo")
	procCreateStdDispatch       = modoleaut32.NewProc("CreateStdDispatch")
	procGetActiveObject         = modoleaut32.NewProc("GetActiveObject")

	procGetMessageW      = moduser32.NewProc("GetMessageW")
	procDispatchMessageW = moduser32.NewProc("DispatchMessageW")
)
