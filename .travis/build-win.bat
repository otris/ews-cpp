@echo off

cmake -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Release -DCURL_DIR=.\vcpkg\installed\x64-windows\share\curl -DZLIB_ROOT=.\vcpkg\installed\x64-windows .
nmake add_delegate
REM ./tests --assets=../tests/assets --gtest_filter=Offline*.*"
