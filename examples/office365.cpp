
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
#include <exception>
#include <iostream>
#include <ostream>
#include <stdlib.h>
#include <string>

// This example shows how to connect to an Office 365 account via EWS and
// retrieve all calendar items in a given month.

using ews::internal::http_request;

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        const ews::basic_credentials credentials(
            "dduck@duckburg.onmicrosoft.com", "secret");

        // First, we use Autodiscover to get the EWS end-point URL that we use
        // to access Office 365. This should always be something like
        // `https://outlook.office365.com/EWS/Exchange.asmx`.

        ews::autodiscover_hints hints;
        hints.autodiscover_url =
            "https://outlook.office365.com/autodiscover/autodiscover.xml";
        const auto autodiscover_result =
            ews::get_exchange_web_services_url<http_request>(
                "dduck@duckburg.onmicrosoft.com", credentials, hints);
        std::cout << "External EWS URL: "
                  << autodiscover_result.external_ews_url << std::endl;

        // Next we create a new ews::service instance in order to connect to
        // this URL.  Note that we use HTML basic authentication rather than
        // NTLM to authenticate.

        auto service =
            ews::service(autodiscover_result.external_ews_url, credentials);

        // Actually, this is all it takes to connect to Office 365. The rest is
        // EWS as with every on-premises Exchange server.  Next we retrieve
        // some calendar items from the account.

        ews::distinguished_folder_id calendar_folder =
            ews::standard_folder::calendar;
        std::string start_date("2017-03-01T00:00:00-07:00");
        std::string end_date("2017-03-31T23:59:59-07:00");

        const auto found_items =
            service.find_item(ews::calendar_view(start_date, end_date),
                              calendar_folder, ews::base_shape::id_only);
        std::cout << "# calender items found: " << found_items.size()
                  << std::endl;

        if (!found_items.empty())
        {

            std::vector<ews::item_id> ids;
            ids.reserve(found_items.size());
            std::transform(begin(found_items), end(found_items),
                           std::back_inserter(ids),
                           [](const ews::calendar_item& calitem) {
                               return calitem.get_item_id();
                           });

            const auto calendar_items = service.get_calendar_items(
                ids, ews::item_shape(ews::base_shape::all_properties));

            for (const auto& cal_item : calendar_items)
            {
                std::cout << "\n";

                std::cout << "Subject: " << cal_item.get_subject() << "\n"
                          << "Start: " << cal_item.get_start().to_string()
                          << "\n"
                          << "End: " << cal_item.get_end().to_string()
                          << "\n\n";
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
