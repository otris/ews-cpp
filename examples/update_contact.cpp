
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

        ews::distinguished_folder_id contacts_folder =
            ews::standard_folder::contacts;
        auto restriction =
            ews::is_equal_to(ews::contact_property_path::email_address_1,
                             "superhero@ducktales.com");
        auto item_ids = service.find_item(contacts_folder, restriction);

        std::cout << "Found " << item_ids.size() << " item(s)\n";

        for (const auto& id : item_ids)
        {
            auto contact = service.get_contact(id);
            auto job_title_property = ews::property(
                ews::contact_property_path::job_title, "Superhero");
            auto new_id = service.update_item(id, job_title_property);
            (void)new_id;

// TODO: check this!
#if 0
            std::cout << "Subject: " << contact.get_subject() << std::endl;
            std::cout << "GivenName: " << contact.get_given_name() << std::endl;
            std::cout << "EmailAddress1: " << contact.get_email_address_1() << std::endl;
            std::cout << "Surname: " << contact.get_surname() << std::endl;
#endif
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
