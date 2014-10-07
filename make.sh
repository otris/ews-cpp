#!/bin/bash

# Clang/libc++:

clang++-3.5 -O0 -g -std=c++11 -stdlib=libc++ \
    -Wall -Wextra -Wshadow \
    test.cpp -o test \
    -lcurl

# GCC/libstdc++:

# g++ -O0 -g -std=c++11 \
#     -Wall -Wextra -pedantic \
#     test.cpp -o test \
#     -lcurl
