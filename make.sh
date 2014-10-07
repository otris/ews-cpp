#!/bin/bash

clang++-3.5 -O0 -g -std=c++14 -stdlib=libc++ \
    -Wall -Wextra -Wshadow \
    test.cpp -o test \
    -lcurl
