
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
        // This is an example of an <AddDelegate> request showing an attempt to
        // give userA delegate permissions on folders that are owned by userB.

        typedef ews::delegate_user::permission_level permission_level;

        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri, env.domain, env.username,
                                    env.password);

        ews::delegate_user::delegate_permissions permissions;
        permissions.calendar_folder = permission_level::author;
        permissions.contacts_folder = permission_level::reviewer;
        auto userA = ews::delegate_user(
            ews::user_id::from_primary_smtp_address("userA@example.com"),
            permissions, false, false);
        std::vector<ews::delegate_user> delegates;
        delegates.push_back(userA);
        service.add_delegate(ews::mailbox("userB@example.com"), delegates);
    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}
