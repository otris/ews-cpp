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

        std::string name = "person";
        auto response =
            service.resolve_names(name, ews::search_scope::active_directory);
        std::cout << response.total_items_in_view << std::endl;
        for (const auto& reso : response.resolutions)
        {
            std::cout << reso.mailbox.name() << std::endl;
            std::cout << reso.mailbox.value() << std::endl;
            std::cout << reso.directory_id.get_id() << std::endl;
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


