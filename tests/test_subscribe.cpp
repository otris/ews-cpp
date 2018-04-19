//   Copyright 2018 otris software AG
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
TEST_F(SubscribeTest, SubscriptionSuccessful)
{
    set_next_fake_response(read_file(assets_dir() / "subscribe_response.xml"));

    ews::distinguished_folder_id folder(ews::standard_folder::inbox);
    auto type = ews::event_type::created_event;
    std::vector<ews::distinguished_folder_id> folders = {folder};
    std::vector<ews::event_type> types = {type};

    auto response = service().subscribe(folders, types, 10);
    EXPECT_EQ(response.get_subscription_id(), "abc-foo-123-456");
    EXPECT_EQ(response.get_watermark(), "AAAAAHgGAAAAAAAAAQ==");
}

TEST_F(SubscribeTest, SendCorrectRequest)
{
    ews::distinguished_folder_id folder(ews::standard_folder::inbox);
    auto type = ews::event_type::new_mail_event;
    std::vector<ews::distinguished_folder_id> folders = {folder};
    std::vector<ews::event_type> types = {type};

    auto response = service().subscribe(folders, types, 10);
    EXPECT_NE(get_last_request().request_string().find(
                  "<soap:Body>"
                  "<m:Subscribe>"
                  "<m:PullSubscriptionRequest>"
                  "<t:FolderIds>"
                  "<t:DistinguishedFolderId Id=\"inbox\"/>"
                  "</t:FolderIds>"
                  "<t:EventTypes>"
                  "<t:EventType>NewMailEvent</t:EventType>"
                  "</t:EventTypes>"
                  "<t:Timeout>10</t:Timeout>"
                  "</m:PullSubscriptionRequest>"),
              std::string::npos);
}
} // namespace tests

#endif // EWS_HAS_FILESYSTEM_HEADER
