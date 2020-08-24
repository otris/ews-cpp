@echo off

cd vcpkg
call bootstrap-vcpkg.bat
vcpkg --triplet x64-windows install curl[winssl]
cd ..