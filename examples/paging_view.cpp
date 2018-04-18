
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

        // First create some draft message
        ews::distinguished_folder_id drafts = ews::standard_folder::drafts;
        for (int i = 0; i < 20; i++)
        {
            auto message = ews::message();
            message.set_subject(
                "This is an e-mail message for our paging view");
            std::vector<ews::mailbox> recipients;
            recipients.push_back(ews::mailbox("donald.duck@duckburg.com"));
            message.set_to_recipients(recipients);
            auto item_id = service.create_item(
                message, ews::message_disposition::save_only);
        }

        // Now iterate over all items in the folder
        // Starting at the end with an offset of 10
        // Returning a number of 5 items per iteration
        // Until no more items are returned
        ews::paging_view view(5, 10, ews::paging_base_point::end);
        while (true)
        {
            auto item_ids = service.find_item(drafts, view);
            if (item_ids.empty())
            {
                std::cout << "No more messages found!\n";
                break;
            }

            for (const auto& id : item_ids)
            {
                auto msg = service.get_message(id);
                std::cout << msg.get_subject() << std::endl;
            }

            view.advance();
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
