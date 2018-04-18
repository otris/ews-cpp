//   Copyright 2017 otris software AG
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

#include <cstring>
#include <string>
#include <vector>

#ifdef EWS_HAS_FILESYSTEM_HEADER
namespace tests
{
TEST_F(ResolveNamesTest, NoUserFound)
{
    set_next_fake_response(read_file(assets_dir() / "resolve_names_error.xml"));

    auto result =
        service().resolve_names("name", ews::search_scope::active_directory);
    EXPECT_TRUE(result.empty());
}
TEST_F(ResolveNamesTest, UserFound)
{
    set_next_fake_response(
        read_file(assets_dir() / "resolve_names_response.xml"));

    auto response =
        service().resolve_names("name", ews::search_scope::active_directory);
    auto resolution_mailbox = response.resolutions[0].mailbox;
    auto resolution_id = response.resolutions[0].directory_id;
    EXPECT_EQ(resolution_mailbox.name(), "User2");
    EXPECT_EQ(resolution_mailbox.value(), "User2@example.com");
    EXPECT_EQ(resolution_id.get_id(), "<GUID=abc-123-foo-bar>");
}

TEST_F(ResolveNamesTest, SendCorrectRequest)
{
    auto response =
        service().resolve_names("name", ews::search_scope::active_directory);
    EXPECT_NE(get_last_request().request_string().find(
                  "<soap:Body>"
                  "<m:ResolveNames ReturnFullContactData=\"true\" "
                  "SearchScope=\"ActiveDirectory\" ContactDataShape=\"IdOnly\">"
                  "<m:UnresolvedEntry>name</m:UnresolvedEntry>"
                  "</m:ResolveNames>"),
              std::string::npos);
}
} // namespace tests

#endif // EWS_HAS_FILESYSTEM_HEADER
