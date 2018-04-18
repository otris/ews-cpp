
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

        // First create a draft message
        ews::distinguished_folder_id drafts = ews::standard_folder::drafts;
        auto message = ews::message();
        message.set_subject("This is an e-mail message for our Contains query");
        std::vector<ews::mailbox> recipients;
        recipients.push_back(ews::mailbox("donald.duck@duckburg.com"));
        message.set_to_recipients(recipients);
        auto item_id =
            service.create_item(message, ews::message_disposition::save_only);

        // Then search for it
        auto search_expression =
            ews::contains(ews::item_property_path::subject, "ess",
                          ews::containment_mode::substring,
                          ews::containment_comparison::ignore_case);

        auto item_ids = service.find_item(drafts, search_expression);

        if (item_ids.empty())
        {
            std::cout << "No messages found!\n";
        }
        else
        {
            for (const auto& id : item_ids)
            {
                auto msg = service.get_message(id);
                std::cout << msg.get_subject() << std::endl;
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
