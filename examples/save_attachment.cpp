
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

#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <exception>
#include <iostream>
#include <ostream>
#include <stdlib.h>
#include <string>

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri, env.domain, env.username,
                                    env.password);

        auto search_expression =
            ews::is_equal_to(ews::item_property_path::has_attachments, true);
        ews::distinguished_folder_id drafts = ews::standard_folder::drafts;
        auto ids = service.find_item(drafts, search_expression);

        if (ids.empty())
        {
            std::cout << "No messages with attachment found!\n";
        }
        else
        {
            // Get and save attachments of first message we just found. We
            // assume these are PNG images

            auto msg = service.get_message(ids[0]);

            auto i = 0;
            auto attachments = msg.get_attachments();
            for (const auto& attachment : attachments)
            {
                auto target_path = "test" + std::to_string(i++) + ".png";
                const auto bytes_written =
                    service.get_attachment(attachment.id())
                        .write_content_to_file(target_path);
                std::cout << bytes_written << std::endl;
            }
        }
    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}
