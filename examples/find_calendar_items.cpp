
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

        ews::distinguished_folder_id calendar = ews::standard_folder::calendar;
        std::string start_date_time("2017-03-01T00:00:00");
        std::string end_date_time("2017-03-31T23:59:59");
        ews::calendar_view cal_view(start_date_time, end_date_time);
        auto cal_items = service.find_item(cal_view, calendar);
        std::cout << "# calender items found: " << cal_items.size()
                  << std::endl;

        for (const auto& cal_item : cal_items)
        {
            std::cout << cal_item.get_start().to_string() << " "
                      << cal_item.get_end().to_string() << ": "
                      << cal_item.get_subject() << ": "
                      << cal_item.get_location() << std::endl;
            const auto& body = cal_item.get_body().content();
            if (!body.empty())
            {
                std::cout << body << std::endl;
            }

            auto resources = cal_item.get_resources();
            for (const auto& resource : resources)
            {
                auto mailbox = resource.get_mailbox();
                std::cout << "  R " << mailbox.name() << std::endl;
            }

            auto req_attendees = cal_item.get_required_attendees();
            for (const auto& att : req_attendees)
            {
                auto mailbox = att.get_mailbox();
                std::cout << " AR " << mailbox.name() << std::endl;
            }

            auto opt_attendees = cal_item.get_optional_attendees();
            for (const auto& att : opt_attendees)
            {
                auto mailbox = att.get_mailbox();
                std::cout << " AO " << mailbox.name() << std::endl;
            }

            auto organizer = cal_item.get_organizer();
            std::cout << "  O " << organizer.name() << std::endl;
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

// vim:et ts=4 sw=4
