
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

        // Get all unread messages from inbox
        ews::distinguished_folder_id inbox = ews::standard_folder::inbox;
        auto search_expression =
            ews::is_equal_to(ews::message_property_path::is_read, false);
        auto item_ids = service.find_item(inbox, search_expression);

        if (item_ids.empty())
        {
            std::cout << "No unread messages found!\n";
        }
        else
        {
            for (const auto& id : item_ids)
            {
                auto msg = service.get_message(id);
                std::cout << "Marking " << msg.get_subject() << " as read\n";
                auto prop =
                    ews::property(ews::message_property_path::is_read, true);
                service.update_item(msg.get_item_id(), prop);
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
