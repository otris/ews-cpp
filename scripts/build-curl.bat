:: Downloads and builds cURL and libcurl on Windows
::
:: Builds DLLs for x64 and amd64 in Release- and Debug-config including .pdb-files
::
:: Requirements:
::   - Visual Studio 2015 installed
::   - 7z.exe in PATH

@setlocal
@set curl_version=7.49.1
@set curl_download_url="http://curl.haxx.se/download/curl-%curl_version%.zip"

@if "%VS140COMNTOOLS%"=="" goto error_no_vscomntools

@where /q 7z.exe
@if %errorlevel% equ 1 goto error_no_7zip

@if not exist curl-%curl_version%.zip (
@bitsadmin /transfer downloadjob /download /priority normal ^
  %curl_download_url% ^
  %cd%\curl-%curl_version%.zip
)
@if not exist curl-%curl_version%.zip goto error_download_failed

@7z.exe x curl-%curl_version%.zip

@pushd curl-%curl_version%\winbuild
@call "%VS140COMNTOOLS%VCVarsQueryRegistry.bat" 32bit 64bit
@cmd.exe /c ""%VSINSTALLDIR%VC\vcvarsall.bat" amd64 && nmake /f Makefile.vc mode=dll MACHINE=x64 VC=14 GEN_PDB=yes"
@cmd.exe /c ""%VSINSTALLDIR%VC\vcvarsall.bat" amd64 && nmake /f Makefile.vc mode=dll MACHINE=x64 VC=14 GEN_PDB=yes DEBUG=yes"
@cmd.exe /c ""%VSINSTALLDIR%VC\vcvarsall.bat" x86 && nmake /f Makefile.vc mode=dll MACHINE=x86 VC=14 GEN_PDB=yes"
@cmd.exe /c ""%VSINSTALLDIR%VC\vcvarsall.bat" x86 && nmake /f Makefile.vc mode=dll MACHINE=x86 VC=14 GEN_PDB=yes DEBUG=yes"
@popd

@robocopy ^
  curl-%curl_version%\builds\libcurl-vc14-x64-release-dll-ipv6-sspi-winssl ^
  win64 /e /move

@robocopy ^
  curl-%curl_version%\builds\libcurl-vc14-x64-debug-dll-ipv6-sspi-winssl ^
  win64-debug /e /move

@robocopy ^
  curl-%curl_version%\builds\libcurl-vc14-x86-release-dll-ipv6-sspi-winssl ^
  win32 /e /move

@robocopy ^
  curl-%curl_version%\builds\libcurl-vc14-x86-debug-dll-ipv6-sspi-winssl ^
  win32-debug /e /move

@goto end

:error_no_vscomntools
@echo ERROR: Cannot determine location of Visual Studio Common Tools folder
@goto end

:error_no_7zip
@echo ERROR: Cannot find 7z.exe in PATH
@goto end

:error_download_failed
@echo ERROR: Downloading curl-%curl_version%.zip failed
@goto end

:end
@if exist curl-%curl_version% ( @rd /s /q curl-%curl_version% )
@endlocal
