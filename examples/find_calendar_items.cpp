
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

#include <algorithm>
#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <exception>
#include <iostream>
#include <iterator>
#include <ostream>
#include <stdlib.h>
#include <string>

// This example demonstrates how to retrieve all occurrences, exceptions, and
// single calendar items on an entire month. It shows a FindItem request with a
// CalendarView instance.
//
// CalendarView makes it easy to get all events from a calendar in a set as
// they appear in a calendar.
//
// It expands all occurrences of a recurring calendar item automatically and
// provides some basic pagination features.
//
// If we wouldn't use CalendarView here we would get a list of single calendar
// items and recurring master calendar items. This means we'd have to expand
// all occurrences of the recurring master ourselves, paying attention to
// exceptions.
//
// The FindItem operation with a CalendarView is a pretty quick query. Note
// however that it cannot return all properties, e.g. an calendar item's body.
// We use a subsequent GetItem operation to get all properties.
int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri, env.domain, env.username,
                                    env.password);

        // First we get all calendar entries in the specified range by using
        // the FindItem operation and a CalendarView.

        ews::distinguished_folder_id calendar_folder =
            ews::standard_folder::calendar;

        // The response includes calendar items that started at
        // 2017-03-01T00:00:00-07:00 or after and ended before
        // 2017-03-31T23:59:59-07:00
        std::string start_date("2017-03-01T00:00:00-07:00");
        std::string end_date("2017-03-31T23:59:59-07:00");

        const auto found_items =
            service.find_item(ews::calendar_view(start_date, end_date),
                              calendar_folder, ews::base_shape::id_only);
        std::cout << "# calender items found: " << found_items.size()
                  << std::endl;

        if (!found_items.empty())
        {

            // Then we retrieve the entire calendar items in a subsequent
            // GetItem operation.

            std::vector<ews::item_id> ids;
            ids.reserve(found_items.size());
            std::transform(begin(found_items), end(found_items),
                           std::back_inserter(ids),
                           [](const ews::calendar_item& calitem) {
                               return calitem.get_item_id();
                           });

            const auto calendar_items = service.get_calendar_items(
                ids, ews::base_shape::all_properties);

            // Done. Now we print some basic properties of each item.

            for (const auto& cal_item : calendar_items)
            {
                std::cout << "\n";

                const auto location = cal_item.get_location();

                std::cout << "Subject: " << cal_item.get_subject() << "\n"
                          << "Start: " << cal_item.get_start().to_string()
                          << "\n"
                          << "End: " << cal_item.get_end().to_string() << "\n"
                          << "Where: " << (location.empty() ? "-" : location)
                          << "\n";

                const auto body = cal_item.get_body();
                if (body.type() == ews::body_type::html)
                {
                    std::cout << "Body: We got some HTML content here!\n";
                }
                else if (body.type() == ews::body_type::plain_text)
                {
                    std::cout << "Body: '" << body.content() << "'\n";
                }

                auto resources = cal_item.get_resources();
                for (const auto& resource : resources)
                {
                    auto mailbox = resource.get_mailbox();
                    std::cout << " R: " << mailbox.name() << "\n";
                }

                auto req_attendees = cal_item.get_required_attendees();
                for (const auto& att : req_attendees)
                {
                    auto mailbox = att.get_mailbox();
                    std::cout << "AR: " << mailbox.name() << "\n";
                }

                auto opt_attendees = cal_item.get_optional_attendees();
                for (const auto& att : opt_attendees)
                {
                    auto mailbox = att.get_mailbox();
                    std::cout << "AO: " << mailbox.name() << "\n";
                }

                auto organizer = cal_item.get_organizer();
                std::cout << " O: " << organizer.name() << std::endl;
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
