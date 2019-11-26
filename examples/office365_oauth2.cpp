
//   Copyright 2019 otris software AG
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
// using OAuth2.

using ews::internal::http_request;

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        // URL of the Outlook instance
        const std::string outlookUrl = "https://outlook.office365.com";

        // The name of your tenant
        const std::string tenant = "example.onmicrosoft.com";

        // The GUID of the client that was granted access to your Office 365 tenant.
        const std::string client_id = "44acad18-b8e0-4fa9-9d2a-53fdeb55b46e";

        // The client secret as provided by Office 365.
        const std::string client_secret = "<do-not-commit-a-client-secret>";

        // URL of the ressource access by the client.
        const std::string ressource = "https://outlook.office365.com";

        // URL of the scope granted to the client
        const std::string scope = "https://outlook.office365.com/.default";

        const ews::oauth2_client_credentials credentials(
            tenant, client_id, client_secret, ressource, scope);

        // Connecting via OAuth2 and client credentials requires impersonation
        // with in EWS. So we use a SMTP address of an user the client should
        // act for.
        const auto someone = ews::connecting_sid(
            ews::connecting_sid::type::smtp_address, "someone@mail.invalid");

        // Next we create a new ews::service instance in order to connect to
        // Office 365.

        auto service = ews::service(outlookUrl, credentials);

        // Actually, this is all it takes to connect to Office 365. The rest is
        // EWS as with every on-premises Exchange server.  Next we retrieve
        // some calendar items from the account.

        ews::distinguished_folder_id calendar_folder =
            ews::standard_folder::calendar;
        std::string start_date("2017-03-01T00:00:00-07:00");
        std::string end_date("2017-03-31T23:59:59-07:00");

        const auto found_items =
            service.impersonate(someone).find_item(ews::calendar_view(start_date, end_date),
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
