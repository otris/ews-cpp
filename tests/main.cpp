
//   Copyright 2016 otris software AG
//
//   Licensed under the Apache License, Version 2.0 (the "License");
//   you may not use this file except in compliance with the License.
//   You may obtain a copy of the License at
//
//       http://www.apache.org/licenses/LICENSE-2.0
//
//   Unless required by applicable law or agreed to in writing, software
//   distributed under the License is distributed on an "AS IS" BASIS,
//   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//   See the License for the specific language governing permissions and
//   limitations under the License.
//
//   This project is hosted at https://github.com/otris

#include <ews/ews_test_support.hpp>
#include <iostream>
#include <ostream>
#include <stdexcept>
#include <stdlib.h>
#include <string.h>
#include <string>
#include <unordered_map>
#include <vector>

#ifdef _WIN32
#    include <direct.h>
#    define getcwd _getcwd
#else
#    include <limits.h>
#    include <unistd.h>
#endif

#include <gtest/gtest.h>

typedef std::unordered_map<std::string, std::string> argument_map;

namespace
{
bool starts_with(const char* prefix, const std::string& str)
{
    return !str.compare(0, strlen(prefix), prefix);
}

bool is(const char* prefix, const std::string& arg, argument_map& map)
{
    if (starts_with(prefix, arg))
    {
        const auto value = arg.substr(strlen(prefix));
        map[prefix] = value;
        return true;
    }
    return false;
}

std::string pwd()
{
    enum
    {
        size = 4096U
    };
    char buf[size];
    auto ret = getcwd(buf, size);
    return ret != nullptr ? std::string(buf)
                          : throw std::runtime_error("getcwd failed");
}

// Parse and interpret command line flags that we recognize. Remove them
// from argv. Leave the rest untouched.
void init_from_args(int* argc, char** argv)
{
    argument_map args;

    for (int i = 1; i < *argc; i++)
    {
        const auto arg = std::string(argv[i]);

        // Do we see a flag that we recognize?
        if (is("--assets=", arg, args))
        {
            // Finally, shift remainder of argv left by one. Note that argv
            // has (*argc + 1) elements, the last one always being null.
            // The following loop moves the trailing null element as well.
            for (int j = i; j != *argc; j++)
            {
                argv[j] = argv[j + 1];
            }

            // Decrements the argument count.
            (*argc)--;

            // We also need to decrement the iterator as we just removed an
            // element.
            i--;
        }
        else if (arg == "--help" || arg == "-h" || arg == "-?" || arg == "/?")
        {
            std::cout << "usage: tests [--assets=PATH]\n"
                         "\n"
                         "--assets=PATH  path to the test assets "
                         "directory, default\n"
                         "               $PWD/tests/assets\n"
                         "\n"
                         "invoke with --gtest_help to see Google Test flags"
                      << std::endl;

            std::exit(EXIT_SUCCESS);
        }
        else if (is("--gtest_help", arg, args))
        {
            return;
        }
    }

#ifdef _WIN32
    auto assets_dir = pwd() + "\\tests\\assets";
#else
    auto assets_dir = pwd() + "/tests/assets";
#endif
    auto it = args.find("--assets=");
    if (it != args.end())
    {
        assets_dir = it->second;

        // If a word begins with an unquoted tilde character replace it
        // with value of $HOME or %USERPROFILE%, respectively

        if (starts_with("~", assets_dir))
        {
#ifdef _WIN32
#    ifdef _MSC_VER
#        pragma warning(push)
#        pragma warning(disable : 4996)
#    endif

            const char* const env = getenv("USERPROFILE");

#    ifdef _MSC_VER
#        pragma warning(pop)
#    endif
#else
            const char* const env = getenv("HOME");
#endif

            assets_dir.replace(0, 1, env);
        }

        // Expand relative paths

        char* res = nullptr;
#ifdef _WIN32
        std::vector<char> abspath(_MAX_PATH);
        res = _fullpath(&abspath[0], assets_dir.c_str(), 1000);
#else
        std::vector<char> abspath(FILENAME_MAX);
        res = realpath(assets_dir.c_str(), &abspath[0]);
#endif
        if (res == nullptr)
        {
            std::cout << "No such directory: '" << assets_dir << "'\n";
            std::exit(EXIT_FAILURE);
        }

        assets_dir = std::string(&abspath[0]);
    }

    ews::test::global_data::instance().assets_dir = assets_dir;
    std::cout << "Loading assets from: '" << assets_dir << "'\n";
}
} // namespace

int main(int argc, char** argv)
{
    init_from_args(&argc, argv);
    std::cout << "Running tests from main.cpp\n";
    testing::InitGoogleTest(&argc, argv);
    return RUN_ALL_TESTS();
}
