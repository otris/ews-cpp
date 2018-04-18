
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

using ews::internal::http_request;

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        const auto env = ews::test::environment();
        ews::basic_credentials credentials(env.autodiscover_smtp_address,
                                           env.autodiscover_password);
        ews::autodiscover_hints hints;
        hints.autodiscover_url =
            "https://exch.otris.de/autodiscover/autodiscover.xml";

        auto result = ews::get_exchange_web_services_url<http_request>(
            env.autodiscover_smtp_address, credentials, hints);

        std::cout << result.internal_ews_url << std::endl;
        std::cout << result.external_ews_url << std::endl;
    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}
