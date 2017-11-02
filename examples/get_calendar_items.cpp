
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

#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>

#include <cstdlib>
#include <exception>
#include <iostream>
#include <ostream>
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

        // We need to be able to acces the rooms mailbox
        auto owner = ews::mailbox("meetingroom1@example.com");

        // Using the calendar_view we can limit the returned appointments
        // to a specific range of time
        // If we want all appointments, we leave out the calendar_view
        ews::distinguished_folder_id calendar_folder(
            ews::standard_folder::calendar, owner);
        std::string start_date("2017-03-01T00:00:00-07:00");
        std::string end_date("2017-03-31T23:59:59-07:00");

        // By using ews::base_shape::all_properties we get the full calendar
        // item, not only the item_id
        const auto found_items =
            service.find_item(ews::calendar_view(start_date, end_date),
                              calendar_folder, ews::base_shape::all_properties);
        std::cout << "# calender items found: " << found_items.size()
                  << std::endl;
    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}

// vim:et ts=4 sw=4

