
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

        // We need to be able to access the room's mailbox
        auto owner = ews::mailbox("meetingroom1@otris.de");

        // Using the calendar_view we can limit the returned appointments
        // to a specific range of time
        // If we want all appointments, we leave out the calendar_view
        ews::distinguished_folder_id calendar_folder(
            ews::standard_folder::calendar, owner);
        std::string start_date("2017-11-05T00:00:00-07:00");
        std::string end_date("2017-11-10T23:00:00-07:00");

        // By using ews::base_shape::all_properties we get the full calendar
        // item, not only the item_id
        ews::item_shape shape;
        shape.set_base_shape(ews::base_shape::all_properties);
        const auto found_items = service.find_item(
            ews::calendar_view(start_date, end_date), calendar_folder, shape);
        std::cout << "# calendar items found: " << found_items.size()
                  << std::endl;

        for (const auto& item : found_items)
        {
            std::cout << item.get_subject();
            std::cout << ", start: " << item.get_start().to_string();
            std::cout << ", organizer: " << item.get_organizer().name();
            std::cout << ", possible attendees: " << item.get_display_to()
                      << std::endl;
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
