#!/bin/bash

mkdir build
cd build

if [ $TRAVIS_OS_NAME = 'windows' ]; then
	cmd.exe /c build-win.bat
else
	cmake -DCMAKE_BUILD_TYPE=Release .. && make && make doc && ./tests --assets=../tests/assets --gtest_filter=Offline*.*
fi