
//   Copyright 2016-2020 otris software AG
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

#include "fixtures.hpp"

#include <ews/ews.hpp>

namespace tests
{
TEST_F(CallbackTest, Default)
{
    auto response =
        service().resolve_names("name", ews::search_scope::active_directory);
    EXPECT_EQ(response.resolutions.size(), 0);
    std::string output = get_callback_output();
    EXPECT_EQ(output, "");
}
TEST_F(CallbackTest, SetActive)
{
    service().set_debug_callback(
        std::bind(&::tests::CallbackTest_SetActive_Test::test_callback, this,
                  std::placeholders::_1));
    auto response =
        service().resolve_names("name", ews::search_scope::active_directory);
    EXPECT_EQ(response.resolutions.size(), 0);

    std::string output = get_callback_output();
    EXPECT_NE(output, "");
    EXPECT_NE(output.find("Info: "), std::string::npos);
}
TEST_F(CallbackTest, SetActiveAndDeactive)
{
    service().set_debug_callback(std::bind(
        &::tests::CallbackTest_SetActiveAndDeactive_Test::test_callback, this,
        std::placeholders::_1));
    auto response =
        service().resolve_names("name", ews::search_scope::active_directory);
    EXPECT_EQ(response.resolutions.size(), 0);

    std::string output_on = get_callback_output();
    EXPECT_NE(output_on, "");

    service().set_debug_callback(nullptr);
    response =
        service().resolve_names("name", ews::search_scope::active_directory);
    EXPECT_EQ(response.resolutions.size(), 0);

    std::string output_off = get_callback_output();
    EXPECT_EQ(output_on, output_off);
}
} // namespace tests
