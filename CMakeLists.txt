cmake_minimum_required(VERSION 2.8.12)
project(ews CXX)

list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")
include(GetGitRevisionDescription)
git_describe(GIT_TAG --tag --candidates=30 --abbrev=0)
message(STATUS "Git tag: ${GIT_TAG}")
get_git_head_revision(GIT_REFSPEC GIT_SHA1)
message(STATUS "Git SHA-1: ${GIT_SHA1}")

include(CheckCXXCompilerFlag)

option(DISABLE_TLS_CERT_VERIFICATION
    "Turn-off verification of the server's authenticity." OFF)
if(NOT ${CMAKE_CXX_COMPILER_ID} STREQUAL MSVC)
    option(ENABLE_SANITIZERS "Enable ASAN and UBSAN sanitizers." OFF)
endif()
option(ENABLE_VERBOSE
    "Enable verbose output to stderr (You don't want this set in production)."
    OFF)
if(${CMAKE_CXX_COMPILER_ID} STREQUAL "Clang")
    set(_ews_option_text "Link against libc++ instead of libstdc++.")
    if(${CMAKE_HOST_SYSTEM_NAME} STREQUAL "Darwin")
        option(ENABLE_LIBCXX ${_ews_option_text} ON)
    else()
        option(ENABLE_LIBCXX ${_ews_option_text} OFF)
    endif()
endif()
if(${CMAKE_CXX_COMPILER_ID} STREQUAL "GNU")
    option(ENABLE_PROFILE "Generate extra code to write profile information." OFF)
endif()

set(SANITIZE_CXXFLAGS)
set(SANITIZE_LDFLAGS)
if(ENABLE_SANITIZERS)
    if("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang" AND
       "${CMAKE_CXX_COMPILER_VERSION}" VERSION_GREATER "4.0")
        set(_enable_ubsan TRUE)
    elseif("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU" AND
           "${CMAKE_CXX_COMPILER_VERSION}" VERSION_GREATER "7.1")
        set(_enable_ubsan TRUE)
    else()
        set(_enable_ubsan FALSE)
    endif()

    if("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang" OR
       "${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
        if(_enable_ubsan)
            list(APPEND SANITIZE_CXXFLAGS
                "-fsanitize=address,undefined -fno-omit-frame-pointer")
            list(APPEND SANITIZE_LDFLAGS
                "-fsanitize=address,undefined")
        else()
            list(APPEND SANITIZE_CXXFLAGS
                "-fsanitize=address -fno-omit-frame-pointer")
            list(APPEND SANITIZE_LDFLAGS
                "-fsanitize=address")
        endif()

    else()
        message(WARNING "Option ENABLE_SANITIZERS only supported with clang and gcc.")
    endif()
endif()

set(GNUCXX_MINIMUM_VERSION "4.8")
set(CLANGCXX_MINIMUM_VERSION "3.5")
set(CXX_STANDARD_TAG "c++17")

# Helper macro. Add -O0 in Debug configuration to prevent any optimization.
# Makes gdb users much happier.
macro(patch_cmake_cxx_debug_flags)
    if("${CMAKE_BUILD_TYPE}" STREQUAL "Debug")
        set(CMAKE_CXX_FLAGS_DEBUG "-O0 ${CMAKE_CXX_FLAGS_DEBUG}")
    endif()
endmacro()

if(${CMAKE_CXX_COMPILER_ID} STREQUAL "Clang")
    patch_cmake_cxx_debug_flags()
    if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS CLANGCXX_MINIMUM_VERSION)
        message(FATAL_ERROR
            "Minimum required clang++ version: ${CLANGCXX_MINIMUM_VERSION}")
    endif()
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -std=${CXX_STANDARD_TAG}")
    if(ENABLE_LIBCXX)
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -stdlib=libc++")
    endif()

    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -Wextra -fcolor-diagnostics")
elseif(${CMAKE_CXX_COMPILER_ID} STREQUAL "GNU")
    patch_cmake_cxx_debug_flags()
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

    if(ENABLE_PROFILE)
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -pg")
    endif()

    # https://gcc.gnu.org/onlinedocs/gcc/Warning-Options.html
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -Wextra -Wpedantic")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wcast-qual -Wcast-align -Wdisabled-optimization -Wformat=2 -Wlogical-op -Wmissing-include-dirs -Wshadow -Wredundant-decls -Wdouble-promotion -Wold-style-cast")

    if("${CMAKE_CXX_COMPILER_VERSION}" VERSION_GREATER "6.0")
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wduplicated-cond")
    endif()

    if("${CMAKE_CXX_COMPILER_VERSION}" VERSION_GREATER "7.0")
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wrestrict -Wduplicated-branches")
    endif()

    # Use colorized output on terminal if supported (GCC 4.9 onwards)
    CHECK_CXX_COMPILER_FLAG("-fdiagnostics-color=auto" GCC_HAS_COLOR)
    if("${GCC_HAS_COLOR}")
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fdiagnostics-color=auto")
    endif()
elseif(${CMAKE_CXX_COMPILER_ID} STREQUAL "Intel")
    message(WARNING "Intel C++ not supported.")
elseif(${CMAKE_CXX_COMPILER_ID} STREQUAL MSVC)
    if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS 17.0)
        message(FATAL_ERROR "Visual Studio must be at least 2012")
    endif()

    if(CMAKE_CXX_FLAGS MATCHES "/W[0-4]")
        string(REGEX REPLACE "/W[0-4]" "/W4" CMAKE_CXX_FLAGS
            "${CMAKE_CXX_FLAGS}")
    else()
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /W4")
    endif()
    set(CMAKE_CXX_FLAGS
        "${CMAKE_CXX_FLAGS} /D_SCL_SECURE_NO_WARNINGS /D_CRT_SECURE_NO_WARNINGS /DNOMINMAX")

    if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS 18.0)
        # Make Google Test compile with Visual Studio 2012
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /D_VARIADIC_MAX=10")
    endif()

    if(CMAKE_CXX_COMPILER_VERSION VERSION_GREATER 18.0)
        # Explicitly enable C++17 language features on Visual Studio 2017
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /std:c++17")
    endif()
else()
    message(AUTHOR_WARNING
        "Could not determine compiler ID: ${CMAKE_CXX_COMPILER_ID}")
endif()

# Uncomment to print out compiler settings for different configuration types
#message(STATUS "Debug: ${CMAKE_CXX_FLAGS_DEBUG}")
#message(STATUS "Release: ${CMAKE_CXX_FLAGS_RELEASE}")
#message(STATUS "RelWithDebInfo: ${CMAKE_CXX_FLAGS_RELWITHDEBINFO}")
#message(STATUS "MinSizeRel: ${CMAKE_CXX_FLAGS_MINSIZEREL}")

set(ews_INCLUDE_DIR ${CMAKE_CURRENT_SOURCE_DIR}/include)
set(ews_SOURCES
    ${ews_INCLUDE_DIR}/ews/ews_fwd.hpp
    ${ews_INCLUDE_DIR}/ews/ews.hpp
    ${ews_INCLUDE_DIR}/ews/ews_test_support.hpp)
set(rapidxml_SOURCES
    ${ews_INCLUDE_DIR}/ews/rapidxml/rapidxml.hpp
    ${ews_INCLUDE_DIR}/ews/rapidxml/rapidxml_iterators.hpp
    ${ews_INCLUDE_DIR}/ews/rapidxml/rapidxml_print.hpp
    ${ews_INCLUDE_DIR}/ews/rapidxml/rapidxml_utils.hpp)

# libcurl 7.35 is packed for Ubuntu 14.04 LTS, 7.29 for RHEL7
find_package(CURL 7.29 REQUIRED)
include_directories(${ews_INCLUDE_DIR} ${CURL_INCLUDE_DIRS})

if(DISABLE_TLS_CERT_VERIFICATION)
    add_definitions(-DEWS_DISABLE_TLS_CERT_VERIFICATION)
endif()

if(ENABLE_VERBOSE)
    add_definitions(-DEWS_ENABLE_VERBOSE)
endif()

# Feature test macros
include(CheckCXXSourceCompiles)

CHECK_CXX_SOURCE_COMPILES("
#include <optional>
int main() { std::optional<int> a; }
" HAS_OPTIONAL)
if(HAS_OPTIONAL)
    add_definitions(-DEWS_HAS_OPTIONAL)
endif()

# Test support for noexcept specifier - negative: VS 2013
CHECK_CXX_SOURCE_COMPILES("
void func() noexcept { }
int main() { }
" HAS_NOEXCEPT_SPECIFIER)
if(HAS_NOEXCEPT_SPECIFIER)
    add_definitions(-DEWS_HAS_NOEXCEPT_SPECIFIER)
endif()

# N3656 - negative: libstdc++ prior 4.9.0
CHECK_CXX_SOURCE_COMPILES("
#include <memory>
int main() { std::unique_ptr<int> foo = std::make_unique<int>(4); }
" HAS_MAKE_UNIQUE)
if(HAS_MAKE_UNIQUE)
    add_definitions(-DEWS_HAS_MAKE_UNIQUE)
endif()

# Type alias declarations
# http://en.cppreference.com/w/cpp/language/type_alias
# negative: VS 2012
CHECK_CXX_SOURCE_COMPILES("
#include <ios>
using flags = std::ios_base::fmtflags;
flags f = std::ios_base::dec;
template <typename T> using ptr = T*;
ptr<int> x;
int main() { }
" HAS_TYPE_ALIAS)
if(HAS_TYPE_ALIAS)
    add_definitions(-DEWS_HAS_TYPE_ALIAS)
endif()

# Default template arguments for function templates - negative: VS 2012
CHECK_CXX_SOURCE_COMPILES("
template <typename Fu = int> void func() { }
int main() { }
" HAS_DEFAULT_TEMPLATE_ARGS_FOR_FUNCTIONS)
if(HAS_DEFAULT_TEMPLATE_ARGS_FOR_FUNCTIONS)
    add_definitions(-DEWS_HAS_DEFAULT_TEMPLATE_ARGS_FOR_FUNCTIONS)
endif()

# Thread local storage - partial: VS 2012
CHECK_CXX_SOURCE_COMPILES("
int main() { thread_local int i; }
" HAS_THREAD_LOCAL_STORAGE)
if(HAS_THREAD_LOCAL_STORAGE)
    add_definitions(-DEWS_HAS_THREAD_LOCAL_STORAGE)
endif()

# N2326 defaulted and deleted functions - negative: VS 2012 of course
CHECK_CXX_SOURCE_COMPILES("
struct Fu { Fu() = default; Fu(const Fu&) = delete; };
int main() { }
" HAS_DEFAULT_AND_DELETE)
if(HAS_DEFAULT_AND_DELETE)
    add_definitions(-DEWS_HAS_DEFAULT_AND_DELETE)
endif()

# Variadic templates - negative: VS 2012
CHECK_CXX_SOURCE_COMPILES("
template <typename... T> void func(T...) { }
int main() { }
" HAS_VARIADIC_TEMPLATES)
if(HAS_VARIADIC_TEMPLATES)
    add_definitions(-DEWS_HAS_VARIADIC_TEMPLATES)
endif()

# Non-buggy <type_traits> implementation - negative VS 2013
CHECK_CXX_SOURCE_COMPILES("
#include <type_traits>
struct Fu { Fu(const Fu&) = delete; };
static_assert(!std::is_copy_constructible<Fu>::value, \"\");
int main() { }
" HAS_NON_BUGGY_TYPE_TRAITS)
if(HAS_NON_BUGGY_TYPE_TRAITS)
    add_definitions(-DEWS_HAS_NON_BUGGY_TYPE_TRAITS)
endif()

# Message in static_assert can be omitted - negative GCC 5 on Ubuntu 16.04
CHECK_CXX_SOURCE_COMPILES("
static_assert(true);
int main() { }
" HAS_CXX17_STATIC_ASSERT)
if(HAS_CXX17_STATIC_ASSERT)
    add_definitions(-DEWS_HAS_CXX17_STATIC_ASSERT)
endif()

# N3671 - negative: _MSC_VER <= 1800, libstdc++ prior 4.9.0
CHECK_CXX_SOURCE_COMPILES("
#include <algorithm>
#include <cstring>
static const char source[] = \"foo\";
static const char dest[] = \"bar\";
int main() { std::equal(source, source + std::strlen(source),
                        dest, dest + std::strlen(dest)); }
" HAS_ROBUST_NONMODIFYING_SEQ_OPS)
if(HAS_ROBUST_NONMODIFYING_SEQ_OPS)
    add_definitions(-DEWS_HAS_ROBUST_NONMODIFYING_SEQ_OPS)
endif()

# Test support for initializer lists - negative: VS 2012
CHECK_CXX_SOURCE_COMPILES("
#include <vector>
int main() { std::vector<int> vec{1, 2, 3, 4, 5}; }
" HAS_INITIALIZER_LISTS)
if(HAS_INITIALIZER_LISTS)
    add_definitions(-DEWS_HAS_INITIALIZER_LISTS)
endif()

# Test support for std::variant
CHECK_CXX_SOURCE_COMPILES("
#include <variant>
int main() { std::variant<int, float> v; }
" HAS_VARIANT)
if(HAS_VARIANT)
    add_definitions(-DEWS_HAS_VARIANT)
endif()

# Test support for C++17 <filesystem> header
CHECK_CXX_SOURCE_COMPILES("
#include <filesystem>
int main() { std::filesystem::path p; }
" HAS_FILESYSTEM_HEADER)
if(HAS_FILESYSTEM_HEADER)
    add_definitions(-DEWS_HAS_FILESYSTEM_HEADER)
endif()

# Helper function for all those example executables
function(add_example EXAMPLE_NAME)
    add_executable(${EXAMPLE_NAME}
        ${ews_SOURCES}
        ${rapidxml_SOURCES}
        examples/${EXAMPLE_NAME}.cpp)
    target_link_libraries(${EXAMPLE_NAME} ${CURL_LIBRARIES})
    set_target_properties(${EXAMPLE_NAME} PROPERTIES
        LINKER_LANGUAGE CXX
        COMPILE_FLAGS "${SANITIZE_CXXFLAGS}"
        LINK_FLAGS "${SANITIZE_LDFLAGS}")
endfunction()

add_example(add_delegate)
add_example(autodiscover)
add_example(create_calendar_item_with_delegate_access)
add_example(create_contact)
add_example(create_folder)
add_example(create_task)
add_example(create_task_with_delegate_access)
add_example(create_task_with_impersonation)
add_example(delete_folder)
add_example(find_calendar_items)
add_example(find_folder)
add_example(find_messages)
add_example(find_tasks)
add_example(find_unread_messages)
add_example(get_delegates)
add_example(get_folder)
add_example(get_meeting_room_calendar)
add_example(get_rooms)
add_example(move_folder)
add_example(move_item)
add_example(office365)
add_example(paging_view)
add_example(raw_soap_request)
add_example(remove_delegate)
add_example(resolve_names)
add_example(save_attachment)
add_example(send_message)
add_example(sync_folder_hierarchy)
add_example(sync_folder_items)
if(HAS_VARIANT AND NOT ${CMAKE_CXX_COMPILER_ID} STREQUAL MSVC)
    add_example(subscribe)
endif()
add_example(update_contact)
add_example(update_folder)

# Python is optional as it is only used by TimeoutTest
find_package(PythonInterp QUIET)
if(PYTHONINTERP_FOUND)
    add_definitions(-DEWS_HAS_PYTHON_EXECUTABLE)
endif()

# Google Test
set(GTEST_ROOT ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/gtest/googletest)
set(GTEST_INCLUDE_DIR ${GTEST_ROOT}/include)
set(GTEST_LIBRARIES gtest)
if(${CMAKE_HOST_SYSTEM_NAME} STREQUAL "Windows")
    set(gtest_force_shared_crt ON CACHE BOOL
        "Use shared (DLL) run-time lib even when Google Test is built as static lib.")
endif()
add_subdirectory(${GTEST_ROOT} EXCLUDE_FROM_ALL)
include_directories(${GTEST_INCLUDE_DIR} ${CMAKE_CURRENT_SOURCE_DIR}/tests)

set(ews_TESTS
    tests/fixtures.hpp
    tests/main.cpp
    tests/regressions/91.cpp
    tests/test_attachments.cpp
    tests/test_autodiscover.cpp
    tests/test_calendar_items.cpp
    tests/test_contacts.cpp
    tests/test_folders.cpp
    tests/test_internals.cpp
    tests/test_items.cpp
    tests/test_messages.cpp
    tests/test_resolve_names.cpp
    tests/test_restrictions.cpp
    tests/test_service.cpp
    tests/test_tasks.cpp
    tests/test_timeout.cpp)

if(HAS_VARIANT)
set(ews_TESTS
    ${ews_TESTS}
    tests/test_subscribe.cpp)
endif()

add_executable(tests
    ${ews_SOURCES}
    ${rapidxml_SOURCES}
    ${ews_TESTS})

set(_libs ${GTEST_LIBRARIES} ${CURL_LIBRARIES})
if(HAS_FILESYSTEM_HEADER AND "${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
    set(_libs ${_libs} stdc++fs)
endif()
target_link_libraries(tests ${_libs})
set_target_properties(tests PROPERTIES
    LINKER_LANGUAGE CXX
    COMPILE_FLAGS "${SANITIZE_CXXFLAGS}"
    LINK_FLAGS "${SANITIZE_LDFLAGS}")

# Target to generate API documentation with Doxygen
find_package(Doxygen)
if(DOXYGEN_FOUND)
    configure_file(${CMAKE_CURRENT_SOURCE_DIR}/Doxyfile.in
        ${CMAKE_CURRENT_BINARY_DIR}/Doxyfile @ONLY)
    add_custom_target(doc
        ${DOXYGEN_EXECUTABLE} ${CMAKE_CURRENT_BINARY_DIR}/Doxyfile
        WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
        COMMENT "Generating API documentation with Doxygen" VERBATIM)
endif()

configure_file("${CMAKE_CURRENT_SOURCE_DIR}/ews_version.hpp.in"
    "${CMAKE_CURRENT_BINARY_DIR}/ews_version.hpp" @ONLY)
list(APPEND ews_SOURCES "${CMAKE_CURRENT_BINARY_DIR}/ews_version.hpp")

install(FILES ${ews_SOURCES} DESTINATION ews)
install(FILES ${rapidxml_SOURCES} DESTINATION ews/rapidxml)

# vim:et sw=4 ts=4
