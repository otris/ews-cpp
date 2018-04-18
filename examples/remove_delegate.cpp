
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
        // Following code shows how to remove two existing delegates from
        // userA's mailbox. In this example, one delegate is removed by using
        // the delegate's primary SMTP address, and the other one is removed by
        // using the delegate's security identifier (SID).

        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri, env.domain, env.username,
                                    env.password);

        std::vector<ews::user_id> delegates;
        delegates.push_back(
            ews::user_id::from_primary_smtp_address("userB@example.com"));
        delegates.push_back(ews::user_id::from_sid(
            "S-1-5-21-1333220396-2200287332-232816053-1118"));
        service.remove_delegate(ews::mailbox("userA@example.com"), delegates);
    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}
