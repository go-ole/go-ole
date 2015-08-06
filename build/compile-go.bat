@ECHO OFF

ECHO "BUILD GOLANG"
cd "%GOROOT%\src"
./make.bat --dist-tool
