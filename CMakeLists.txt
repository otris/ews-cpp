cmake_minimum_required(VERSION 2.8)
project(ews CXX)

include(CheckCXXCompilerFlag)

option(ENABLE_ASSERTS "Enable run-time checks." OFF)
option(ENABLE_ASAN "Enable address sanitizer." OFF)

set(SANITIZE_CXXFLAGS)
set(SANITIZE_LDFLAGS)
if(ENABLE_ASAN)
    if("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang" OR
       "${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
        list(APPEND SANITIZE_CXXFLAGS
            "-fsanitize=address -fno-omit-frame-pointer")
        list(APPEND SANITIZE_LDFLAGS
            "-fsanitize=address")
    else()
        message(WARNING "Option ENABLE_ASAN only supported with clang and gcc.")
    endif()
endif()

if("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang")
    set(CMAKE_CXX_FLAGS
        "${CMAKE_CXX_FLAGS} -std=c++11 -stdlib=libc++ -Wall -Wextra -Wshadow")
elseif("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
    set(CMAKE_CXX_FLAGS
        "${CMAKE_CXX_FLAGS} -std=c++11 -Wall -Wextra -pedantic")

    # Use colorized output on terminal if supported (GCC 4.9 onwards)
    CHECK_CXX_COMPILER_FLAG("-fdiagnostics-color=auto" GCC_HAS_COLOR)
    if("${GCC_HAS_COLOR}")
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fdiagnostics-color=auto")
    endif()
elseif("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Intel")
    message(WARNING "Intel C++ not configured.")
elseif("${CMAKE_CXX_COMPILER_ID}" STREQUAL "MSVC")
    if(CMAKE_CXX_FLAGS MATCHES "/W[0-4]")
        string(REGEX REPLACE "/W[0-4]" "/W4" CMAKE_CXX_FLAGS
            "${CMAKE_CXX_FLAGS}")
    else()
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /W4")
    endif()
endif()

set(ews_INCLUDE_DIR ${CMAKE_CURRENT_SOURCE_DIR}/include)
set(ews_SOURCES
    ${ews_INCLUDE_DIR}/ews/plumbing.hpp
    ${ews_INCLUDE_DIR}/ews/porcelain.hpp)

# libcurl 7.22 is packed for Ubuntu 12.04 LTS
find_package(CURL 7.22 REQUIRED)
include_directories(${ews_INCLUDE_DIR} ${CURL_INCLUDE_DIRS})

add_subdirectory(3rdparty/gtest)
set(GTEST_ROOT ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/gtest)
set(GTEST_INCLUDE_DIR ${GTEST_ROOT}/include)
set(GTEST_LIBRARIES gtest gtest_main)
include_directories(${GTEST_INCLUDE_DIR})

if(ENABLE_ASSERTS)
    add_definitions(-DEWS_ENABLE_ASSERTS)
endif()

add_executable(tests
    ${ews_SOURCES}
    tests/test_fu.cpp)

target_link_libraries(tests ${GTEST_LIBRARIES} ${CURL_LIBRARIES})

set_target_properties(tests PROPERTIES
    LINKER_LANGUAGE CXX
    COMPILE_FLAGS "${CMAKE_CXX_FLAGS} ${SANITIZE_CXXFLAGS}"
    LINK_FLAGS "${SANITIZE_LDFLAGS}")

# vim:et sw=4 ts=4
