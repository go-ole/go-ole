# Go OLE

[![Build status](https://ci.appveyor.com/api/projects/status/qr0u2sf7q43us9fj?svg=true)](https://ci.appveyor.com/project/jacobsantos/go-ole-jgs28)
[![Build Status](https://travis-ci.org/go-ole/go-ole.svg?branch=master)](https://travis-ci.org/go-ole/go-ole)
[![GoDoc](https://godoc.org/github.com/go-ole/go-ole?status.svg)](https://godoc.org/github.com/go-ole/go-ole)

Go bindings for Windows COM using shared libraries instead of cgo.

By Yasuhiro Matsumoto.

## Install

To experiment with go-ole, you can just compile and run the example program:

```
go get github.com/go-ole/go-ole
cd /path/to/go-ole/
go test

cd /path/to/go-ole/example/excel
go run excel.go
```

## Guide

The library provides an utility for accessing the COM and working around its unsafe execution. When designing your API
on top of this library, it is recommended that you use native types and convert to the `ole.VARIANT` type in the wrapper
function or function pointer.

### Upgrading from 1.x

#### VARIANT

- There is a new API for registering and unregistering conversions for `ole.VT` and native go types.
- `ole.VARIANT` no longer has function pointers.
  - Use `ole.WrapVariant` to convert supported native Go types to `ole.VARIANT`.
  - Use `ole.UnwrapVariant` to convert supported `ole.VT` to native Go types.
  - The API uses Generics to directly cast to the native type instead of requiring additional redirection.
-`CoInitializeEx` is now `Initialize`.
-`CoUninitialize` is now `Uninitialize`.
- `IUnknown` uses generics to convert directly to the OLE Automation object. COM/OLE functions will directly return the
  provided OLE Automation Object instead of requiring indirection starting with `ole.IUnknown` to `ole.IDispatch`. You
  may now directly go to `ole.IDispatch`.

### VARIANT Types

Support for converting to and from `ole.VARIANT` type is provided through the `ole.RegisterVariantConverters()` function.

```go
package main

import "github.com/go-ole/go-ole"

func main() {
	ole.RegisterVariantConverters()
}
```

Not all native types will be known how to convert to a `ole.VARIANT` and a function is provided to help with the
conversion.

```go
package main

import (
   "github.com/go-ole/go-ole"
   "unsafe"
)

func main() {
   ole.RegisterVariantConverter[bool](
      ole.VT_BOOL,
      func(a any) *ole.VARIANT {
         var val int64
         if a.(bool) {
            val = -1
         } else {
            val = 0
         }
         return &ole.VARIANT{VT: ole.VT_BOOL, Val: val}
      },
      func(variant *ole.VARIANT) any {
         return variant.Val != 0
      },
   )

   ole.RegisterVariantConverter[*bool](
      ole.VT_BOOL | ole.VT_BYREF,
      func(a any) *ole.VARIANT {
         var val int64
         if a.(bool) {
            val = -1
         } else {
            val = 0
         }
         return &ole.VARIANT{VT: ole.VT_BOOL | ole.VT_BYREF, Val: int64(uintptr(unsafe.Pointer(val)))}
      },
      func(variant *ole.VARIANT) any {
         return variant.Val != 0
      },
   )
}
```

You may also remove registered or supported variant types and native types.

```go
package main

import "github.com/go-ole/go-ole"

func main() {
	ole.DeregisterToVariantConverter[bool]()
    ole.DeregisterFromVariantConverter(ole.VT_BOOL)
	
    ole.DeregisterToVariantConverter[*bool]()
    ole.DeregisterFromVariantConverter(ole.VT_BOOL | ole.VT_BYREF)
}
```

### VARIANT Conversion

There are two functions that make converting to `ole.VARIANT` and from `ole.VARIANT`.

```go
package main

import "github.com/go-ole/go-ole"

func main() {
	var val int64
	var unwrapped int64
	val = 1024
	variant, err := ole.WrapVariant[int64](val)
	
	// If the type is known, then you could get away with not providing the generic type.
	variant, err = ole.WrapVariant(val)
	
	if err != nil {
		// Do something
    }
	
	unwrapped = ole.UnwrapVariant[int64](variant)
	
	if unwrapped != val {
		panic("VARIANT conversion did not work!")
    }
}
```

If performance is important, then please note that you may want to call to the native conversion functions `ole.[Type]ToVariant`
when dispatching.

## Testing

1. Download a release from https://github.com/go-ole/test-com-server
2. Register the COM server
   - On Windows 32-bit: `c:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe /codebase /nologo c:\path\to\TestCOMServer.dll`
   - On Windows 64-bit: `c:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe /codebase /nologo c:\path\to\TestCOMServer.dll`
3. `go test`

## Multithreading

You have two solutions for handling gothreads and multithreading.

1. You may lock the function or gothread to a single thread. 
    ```go
    runtime.LockOSThread()
    defer runtime.UnlockOSThread()
    ```
2. Use [scjalliance/comshim](https://github.com/scjalliance/comshim)

The key to any solution is that you must call `CoUninitialize()` or `Uninitialize()` for every `CoInitialize()` or `Initialize()`.

## Continuous Integration

Continuous integration configuration has been added for both Travis-CI and AppVeyor. You will have to add these to your own account for your fork in order for it to run.

**Travis-CI**

Travis-CI was added to check builds on Linux to ensure that `go get` works when cross building. Currently, Travis-CI is not used to test cross-building, but this may be changed in the future. It is also not currently possible to test the library on Linux, since COM API is specific to Windows and it is not currently possible to run a COM server on Linux or even connect to a remote COM server.

**AppVeyor**

AppVeyor is used to build on Windows using the (in-development) test COM server. It is currently only used to test the build and ensure that the code works on Windows. It will be used to register a COM server and then run the test cases based on the test COM server.

The tests currently do run and do pass and this should be maintained with commits.

## Versioning

Go OLE uses [semantic versioning](http://semver.org) for version numbers, which is similar to the version contract of the Go language. Which means that the major version will always maintain backwards compatibility with minor versions. Minor versions will only add new additions and changes. Fixes will always be in patch. 

This contract should allow you to upgrade to new minor and patch versions without breakage or modifications to your existing code. Leave a ticket, if there is breakage, so that it could be fixed.

## LICENSE

Under the MIT License: http://mattn.mit-license.org/2013
