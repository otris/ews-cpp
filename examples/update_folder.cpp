
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

        ews::distinguished_folder_id inbox = ews::standard_folder::inbox;

        auto folder_ids = service.find_folder(inbox);
        if (!folder_ids.empty())
        {
            // Get first folder
            auto folder = service.get_folder(folder_ids.front());
            auto name = folder.get_display_name();

				// Add suffix to folders display
				// or remove suffix if already exists
            auto pos = name.find("_updated");
            if (pos == std::string::npos)
            {
                name += "_updated";
            }
            else
            {
                name = name.substr(0, pos);
            }

            // Create property and update
            auto prop = ews::property(
                ews::folder_property_path::display_name, name);
            auto new_id = service.update_folder(folder.get_folder_id(), prop);
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
