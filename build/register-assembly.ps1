# Register TestCOMServer v2 COM classes by writing registry entries directly.
# .NET 9 comhost DLLs do not export DllRegisterServer, so regsvr32 cannot be used.
#
# Usage: .\register-assembly.ps1 -DllPath <path-to-TestCOMServer.comhost.dll> [-Unregister]

param(
    [Parameter(Mandatory=$true)]
    [string]$DllPath,
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"

$DllPath = (Resolve-Path $DllPath).Path

if (-not (Test-Path $DllPath)) {
    Write-Error "File not found: $DllPath"
    exit 1
}

# All CLSIDs from test-com-server v2.0.0
$clsids = @(
    # Types
    "{A1B2C3D4-1111-1111-1111-000000000002}",  # COMTestInt8
    "{A1B2C3D4-1111-1111-1111-000000000004}",  # COMTestInt16
    "{A1B2C3D4-1111-1111-1111-000000000006}",  # COMTestInt32
    "{A1B2C3D4-1111-1111-1111-000000000008}",  # COMTestInt64
    "{A1B2C3D4-1111-1111-1111-00000000000A}",  # COMTestFloat32
    "{A1B2C3D4-1111-1111-1111-00000000000C}",  # COMTestFloat64
    "{A1B2C3D4-1111-1111-1111-00000000000E}",  # COMTestString
    "{A1B2C3D4-1111-1111-1111-000000000010}",  # COMTestBoolean
    "{A1B2C3D4-1111-1111-1111-000000000012}",  # COMTestCurrency
    "{A1B2C3D4-1111-1111-1111-000000000014}",  # COMTestDate
    "{A1B2C3D4-1111-1111-1111-000000000016}",  # COMTestDecimal
    "{A1B2C3D4-1111-1111-1111-000000000018}",  # COMTestError
    "{A1B2C3D4-1111-1111-1111-00000000001A}",  # COMTestVariant
    "{A1B2C3D4-1111-1111-1111-00000000001C}",  # COMTestUnknown
    "{A1B2C3D4-1111-1111-1111-00000000001E}",  # COMTestDispatch
    "{A1B2C3D4-1111-1111-1111-000000000020}",  # COMTestEmpty
    "{A1B2C3D4-1111-1111-1111-000000000022}",  # COMTestClsid
    "{A1B2C3D4-1111-1111-1111-000000000024}",  # COMTestHResult
    "{A1B2C3D4-1111-1111-1111-000000000026}",  # COMTestFileTime
    "{A1B2C3D4-1111-1111-1111-000000000028}",  # COMTestStream
    "{A1B2C3D4-1111-1111-1111-000000000029}",  # ManagedStream
    "{A1B2C3D4-1111-1111-1111-00000000002B}",  # COMTestBlob
    "{A1B2C3D4-1111-1111-1111-00000000002D}",  # COMTestPtr
    # SafeArrays
    "{A1B2C3D4-2222-2222-2222-000000000002}",  # COMSafeArrayInt8
    "{A1B2C3D4-2222-2222-2222-000000000004}",  # COMSafeArrayInt16
    "{A1B2C3D4-2222-2222-2222-000000000006}",  # COMSafeArrayInt32
    "{A1B2C3D4-2222-2222-2222-000000000008}",  # COMSafeArrayInt64
    "{A1B2C3D4-2222-2222-2222-00000000000A}",  # COMSafeArrayFloat32
    "{A1B2C3D4-2222-2222-2222-00000000000C}",  # COMSafeArrayFloat64
    "{A1B2C3D4-2222-2222-2222-00000000000E}",  # COMSafeArrayString
    "{A1B2C3D4-2222-2222-2222-000000000010}",  # COMSafeArrayBoolean
    "{A1B2C3D4-2222-2222-2222-000000000012}",  # COMSafeArrayCurrency
    "{A1B2C3D4-2222-2222-2222-000000000014}",  # COMSafeArrayDate
    "{A1B2C3D4-2222-2222-2222-000000000016}",  # COMSafeArrayDecimal
    "{A1B2C3D4-2222-2222-2222-000000000018}",  # COMSafeArrayVariant
    # Interfaces
    "{A1B2C3D4-3333-3333-3333-000000000002}",  # COMTestDualInterface
    "{A1B2C3D4-3333-3333-3333-000000000004}",  # COMTestUnknownOnly
    "{A1B2C3D4-3333-3333-3333-000000000006}",  # COMTestDispatchOnly
    "{A1B2C3D4-3333-3333-3333-00000000000A}",  # COMTestMultipleInterfaces
    "{A1B2C3D4-3333-3333-3333-00000000000D}",  # COMTestInheritedInterface
    "{A1B2C3D4-3333-3333-3333-000000000010}"   # COMTestConnectionPoint
)

foreach ($clsid in $clsids) {
    $keyPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"
    $inprocPath = "$keyPath\InprocServer32"

    if ($Unregister) {
        if (Test-Path $keyPath) {
            Remove-Item -Path $keyPath -Recurse -Force
            Write-Host "Removed $clsid"
        }
    } else {
        New-Item -Path $inprocPath -Force | Out-Null
        Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value $DllPath
        Set-ItemProperty -Path $inprocPath -Name "ThreadingModel" -Value "Both"
        Write-Host "Registered $clsid"
    }
}

# Verify one CLSID
if (-not $Unregister) {
    $testKey = "HKLM:\SOFTWARE\Classes\CLSID\{A1B2C3D4-1111-1111-1111-00000000000E}\InprocServer32"
    if (Test-Path $testKey) {
        $val = (Get-ItemProperty -Path $testKey).'(Default)'
        Write-Host "Verification: COMTestString InprocServer32 = $val"
    } else {
        Write-Error "Verification failed: COMTestString CLSID not found in registry"
        exit 1
    }
}

Write-Host "Done. $($clsids.Count) CLSIDs processed."
