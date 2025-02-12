//go:build windows
// +build windows

package legacy

import (
	"github.com/go-ole/go-ole"
	"reflect"
	"syscall"
	"unsafe"
)
