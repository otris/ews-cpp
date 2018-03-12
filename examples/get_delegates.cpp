
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

namespace
{
inline std::ostream&
operator<<(std::ostream& os,
           const ews::delegate_user::delegate_permissions& perms)
{
    os << "Calendar: " << ews::internal::enum_to_str(perms.calendar_folder)
       << "\n";
    os << "Tasks:    " << ews::internal::enum_to_str(perms.tasks_folder)
       << "\n";
    os << "Inbox:    " << ews::internal::enum_to_str(perms.inbox_folder)
       << "\n";
    os << "Contacts: " << ews::internal::enum_to_str(perms.contacts_folder)
       << "\n";
    os << "Notes:    " << ews::internal::enum_to_str(perms.notes_folder);
    return os;
}
} // namespace

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {

        // This example shows how to retrieve the delegates of a mailbox along
        // with the delegate's permissions using EWS <GetDelegate> operation.

        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri, env.domain, env.username,
                                    env.password);
        const auto users =
            service.get_delegate(ews::mailbox("test@example.com"), true);

        std::cout << "# of delegate users: " << users.size() << "\n";
        for (const auto& user : users)
        {
            std::cout << "→ " << user.get_user_id().get_display_name() << " <"
                      << user.get_user_id().get_primary_smtp_address() << ">\n";
            std::cout << "Permissions:\n";
            std::cout << user.get_permissions() << "\n";
            std::cout << "Delegate can see private items: "
                      << (user.get_view_private_items() ? "yes" : "no") << "\n";
            std::cout
                << "Delegate receives copies of meeting-related messages: "
                << (user.get_receive_copies_of_meeting_messages() ? "yes"
                                                                  : "no")
                << std::endl;
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
