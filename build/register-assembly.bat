@ECHO OFF
REM Legacy wrapper - use register-assembly.ps1 instead.
REM Usage: register-assembly.bat <path-to-TestCOMServer.comhost.dll>

powershell -ExecutionPolicy Bypass -File "%~dp0register-assembly.ps1" -DllPath "%~1"
