:: Download and build curl.exe and libcurl on Windows.
::
:: Builds DLLs and curl.exe for x86 and amd64 in Release- and Debug
:: configuration including .pdb files.
::
:: Requirements:
::   - Visual Studio 2017 installed (Tested with Professional edition but
::     Community should work fine, too.)
::   - 7z.exe in PATH

@setlocal
@set curl_version=7.59.0
@set curl_download_url="https://curl.haxx.se/download/curl-%curl_version%.zip"

@if exist "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\Common7\Tools\VsDevCmd.bat" (
  @set vsdevcmd="C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\Common7\Tools\VsDevCmd.bat"
) else (
  @set vsdevcmd="C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\Tools\VsDevCmd.bat"
)
@call %vsdevcmd%
@if "%VS150COMNTOOLS%"=="" goto error_no_vscomntools

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
@cmd.exe /c "@call %vsdevcmd% -no_logo -arch=amd64 && nmake /f Makefile.vc mode=dll MACHINE=x64 VC=15 GEN_PDB=yes"
@cmd.exe /c "@call %vsdevcmd% -no_logo -arch=amd64 && nmake /f Makefile.vc mode=dll MACHINE=x64 VC=15 GEN_PDB=yes DEBUG=yes"
@cmd.exe /c "@call %vsdevcmd% -no_logo -arch=x86 && nmake /f Makefile.vc mode=dll MACHINE=x86 VC=15 GEN_PDB=yes"
@cmd.exe /c "@call %vsdevcmd% -no_logo -arch=x86 && nmake /f Makefile.vc mode=dll MACHINE=x86 VC=15 GEN_PDB=yes DEBUG=yes"
@popd

@robocopy ^
  curl-%curl_version%\builds\libcurl-vc15-x64-release-dll-ipv6-sspi-winssl ^
  win64 /e /move

@robocopy ^
  curl-%curl_version%\builds\libcurl-vc15-x64-debug-dll-ipv6-sspi-winssl ^
  win64-debug /e /move

@robocopy ^
  curl-%curl_version%\builds\libcurl-vc15-x86-release-dll-ipv6-sspi-winssl ^
  win32 /e /move

@robocopy ^
  curl-%curl_version%\builds\libcurl-vc15-x86-debug-dll-ipv6-sspi-winssl ^
  win32-debug /e /move

@goto end

:error_no_vscomntools
@echo ERROR: Cannot determine location of Visual Studio 2017 Common Tools folder
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
exit /b 0
