//go:build !windows
// +build !windows

package oleutil

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/legacy"
)

// ConnectObject creates a connection point between two services for communication.
func ConnectObject(disp *ole.IDispatch, iid *legacy.GUID, idisp interface{}) (uint32, error) {
	return 0, legacy.NewError(legacy.E_NOTIMPL)
}
