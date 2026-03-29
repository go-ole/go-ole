# Version 2.0.0-alpha

This release is a major API and implementation overhaul. The v2 branch moves the Windows interop layer onto `golang.org/x/sys/windows`, introduces generic helpers for COM and `VARIANT` handling, restructures packages, and expands WinRT and SafeArray support.

## User-Facing Changes

### New and Expanded APIs

* `Initialize` now returns `InitializeResult` plus `error`, making COM initialization state explicit.
  * Added `InitializeMultithreaded` and `InitializeApartmentThreaded` convenience helpers.
  * Added `ConcurrencyModel` to constrain `Initialize` inputs.
* Added generic COM querying helpers so callers can request concrete interface types directly from `IUnknown`.
* Added a new generic `VARIANT` conversion model.
  * `WrapVariant` and `UnwrapVariant` convert between native Go values and `*VARIANT`.
  * Added registration and deregistration APIs for custom `VT` and Go-type conversions.
* Added or expanded wrappers for WinRT and COM interfaces including `IInspectable`, `IActivationFactory`, `IEnumConnections`, `IRecordInfo`, `IConnectionPoint`, and `IConnectionPointContainer`.
* Added WinRT helpers such as `RoInitialize`, `RoUninitialize`, `RoActivateInstance`, `RoGetActivationFactory`, and `HString`.
* Added native Go date conversion improvements and broader scalar/variant coverage.

### Package and Layout Changes

* SafeArray support moved into the `safearray` package and gained broader conversion helpers.
* Connection-point helpers moved into the `server` package.
* Added lookup helpers such as `ClassIdFromProgramId`, `ClassIdFromGuidString`, `ClassIdFromString`, `ClassIdToString`, `InterfaceIdFromString`, and `InterfaceIdToString`.

### Documentation and Tooling

* Refreshed examples and README guidance for the v2 APIs, including multithreading guidance.
* Added GitHub Actions workflows for x86 and x64 Windows testing and updated the COM test server download steps.

## Breaking Changes

### Renamed

* `CoInitializeEx` is now `Initialize`.
* `CoUninitialize` is now `Uninitialize`.
* `CLSIDFromProgID` is now `ClassIdFromProgramId`.
* `CLSIDFromString` is now `ClassIdFromGuidString`.
* `IIDFromString` is now `InterfaceIdFromString`.
* `StringFromCLSID` is now `ClassIdToString`.
* `StringFromIID` is now `InterfaceIdToString`.

### Behavioral and Structural Changes

* Many COM helper APIs now return strongly typed interface pointers directly instead of requiring intermediate `IUnknown` casts.
* `IDispatch` helpers now operate on `*VARIANT` parameters directly instead of converting `interface{}` values internally.
* Package layout changed substantially: helpers moved between the repository root, `safearray`, and `server`.
* The project now targets the modern Go module/toolchain used by the v2 branch.

### Removed

* The custom `GUID` type was removed in favor of `golang.org/x/sys/windows`.`GUID`.
* `CoInitialize` was removed; use `Initialize()` instead.
* `OleError` and the older custom error plumbing were removed in favor of `golang.org/x/sys/windows` error values.
* The previous `oleutil` layout and many legacy compatibility/helper files were removed during the refactor.

## Internal Refactors

* Reworked the Windows interop layer to use `golang.org/x/sys/windows` types and APIs throughout.
* Replaced many older helper/shim files (`*_func.go`, `*_windows.go`, legacy error wrappers, and related glue code) with direct interface wrappers and syscall-based implementations.
* Removed or consolidated a large number of older wrapper files while moving functionality into the new interface layout.
* Expanded tests around WinRT, variants, SafeArray support, and interface wrappers during the overhaul.

# Version 1.x.x

* **Add more test cases and reference new test COM server project.** (Placeholder for future additions)

# Version 1.2.0-alphaX

**Minimum supported version is now Go 1.4. Go 1.1 support is deprecated, but should still build.**

 * Added CI configuration for Travis-CI and AppVeyor.
 * Added test InterfaceID and ClassID for the COM Test Server project.
 * Added more inline documentation (#83).
 * Added IEnumVARIANT implementation (#88).
 * Added IEnumVARIANT test cases (#99, #100, #101).
 * Added support for retrieving `time.Time` from VARIANT (#92).
 * Added test case for IUnknown (#64).
 * Added test case for IDispatch (#64).
 * Added test cases for scalar variants (#64, #76).

# Version 1.1.1

 * Fixes for Linux build.
 * Fixes for Windows build.

# Version 1.1.0

The change to provide building on all platforms is a new feature. The increase in minor version reflects that and allows those who wish to stay on 1.0.x to continue to do so. Support for 1.0.x will be limited to bug fixes.

 * Move GUID out of variables.go into its own file to make new documentation available.
 * Move OleError out of ole.go into its own file to make new documentation available.
 * Add documentation to utility functions.
 * Add documentation to variant receiver functions.
 * Add documentation to ole structures.
 * Make variant available to other systems outside of Windows.
 * Make OLE structures available to other systems outside of Windows.

## New Features

 * Library should now be built on all platforms supported by Go. Library will NOOP on any platform that is not Windows.
 * More functions are now documented and available on godoc.org.

# Version 1.0.1

 1. Fix package references from repository location change.

# Version 1.0.0

This version is stable enough for use. The COM API is still incomplete, but provides enough functionality for accessing COM servers using IDispatch interface.

There is no changelog for this version. Check commits for history.
