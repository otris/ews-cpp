@echo off

set VC_INSTALL_DIR=%ProgramFiles(x86)%\Microsoft Visual Studio\2017\BuildTools\VC
set VC_VARSALL_BAT="%VC_INSTALL_DIR%\Auxiliary\Build\vcvarsall.bat"

echo Start build using %VC_VARSALL_BAT%
cmd.exe /c "%VC_VARSALL_BAT% x64 && .travis\init-vcpkg.bat"
cmd.exe /c "%VC_VARSALL_BAT% x64 && .travis\build-win.bat"
