cmake_minimum_required(VERSION 2.8)
project(ews CXX)

include(CheckCXXCompilerFlag)

option(ENABLE_ASSERTS "Enable run-time checks." OFF)
option(ENABLE_ASAN "Enable address sanitizer." OFF)
option(ENABLE_VERBOSE "Enable verbose output to stderr." OFF)
if("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang")
    option(ENABLE_LIBCXX "Link against libc++ instead of libstdc++." ON)
endif()

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

set(GNUCXX_MINIMUM_VERSION "4.8")
set(CLANGCXX_MINIMUM_VERSION "3.5")
set(CXX_STANDARD_TAG "c++14")

# Clang
if("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang")
    if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS CLANGCXX_MINIMUM_VERSION)
        message(FATAL_ERROR
            "Minimum required clang++ version: ${CLANGCXX_MINIMUM_VERSION}")
    endif()
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -std=${CXX_STANDARD_TAG}")
    if(ENABLE_LIBCXX)
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -stdlib=libc++")
    endif()
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -Wextra -Wshadow")

# GCC
elseif("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
    if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS GNUCXX_MINIMUM_VERSION)
        message(FATAL_ERROR
            "Minimum required g++ version: ${GNUCXX_MINIMUM_VERSION}")
    endif()

    if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS "4.9")
        set(CMAKE_CXX_FLAGS
            "${CMAKE_CXX_FLAGS} -std=c++1y")
    else()
        set(CMAKE_CXX_FLAGS
            "${CMAKE_CXX_FLAGS} -std=${CXX_STANDARD_TAG}")
    endif()
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -Wextra -pedantic")

    # Use colorized output on terminal if supported (GCC 4.9 onwards)
    CHECK_CXX_COMPILER_FLAG("-fdiagnostics-color=auto" GCC_HAS_COLOR)
    if("${GCC_HAS_COLOR}")
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fdiagnostics-color=auto")
    endif()

# Intel
elseif("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Intel")
    message(WARNING "Intel C++ not supported.")

# MSVC
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
    ${ews_INCLUDE_DIR}/ews/ews.hpp
    ${ews_INCLUDE_DIR}/ews/ews_test_support.hpp)
set(rapidxml_SOURCES
    ${ews_INCLUDE_DIR}/ews/rapidxml/rapidxml.hpp
    ${ews_INCLUDE_DIR}/ews/rapidxml/rapidxml_iterators.hpp
    ${ews_INCLUDE_DIR}/ews/rapidxml/rapidxml_print.hpp
    ${ews_INCLUDE_DIR}/ews/rapidxml/rapidxml_utils.hpp)

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

if(ENABLE_VERBOSE)
    add_definitions(-DEWS_ENABLE_VERBOSE)
endif()

add_executable(tests
    ${ews_SOURCES}
    ${rapidxml_SOURCES}
    tests/test_constants.cpp
    tests/test_item_id.cpp
    tests/test_tasks.cpp)
target_link_libraries(tests ${GTEST_LIBRARIES} ${CURL_LIBRARIES})
set_target_properties(tests PROPERTIES
    LINKER_LANGUAGE CXX
    COMPILE_FLAGS "${CMAKE_CXX_FLAGS} ${SANITIZE_CXXFLAGS}"
    LINK_FLAGS "${SANITIZE_LDFLAGS}")

# Helper function for all those example executables
function(add_example EXAMPLE_NAME)
    add_executable(${EXAMPLE_NAME}
        ${ews_SOURCES}
        ${rapidxml_SOURCES}
        examples/${EXAMPLE_NAME}.cpp)
    target_link_libraries(${EXAMPLE_NAME} ${CURL_LIBRARIES})
    set_target_properties(${EXAMPLE_NAME} PROPERTIES
        LINKER_LANGUAGE CXX
        COMPILE_FLAGS "${CMAKE_CXX_FLAGS} ${SANITIZE_CXXFLAGS}"
        LINK_FLAGS "${SANITIZE_LDFLAGS}")
endfunction()

add_example(raw_soap_request)
add_example(create_task)

install(FILES ${ews_SOURCES} DESTINATION include/ews)
install(FILES ${rapidxml_SOURCES} DESTINATION include/ews/rapidxml)

# vim:et sw=4 ts=4
