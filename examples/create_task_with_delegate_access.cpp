
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

        // This example shows how to create a task inside another user's
        // mailbox.

        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri, env.domain, env.username,
                                    env.password);

        auto task = ews::task();
        task.set_subject("Get EWS delegation working");
        task.set_body(
            ews::body("Support adding, removing, retrieving delegates as well "
                      "as explicit and implicit access."));
        task.set_start_date(ews::date_time("2017-07-30T18:00:00Z"));
        task.set_due_date(ews::date_time("2017-07-30T19:00:00Z"));
        task.set_reminder_enabled(true);
        task.set_reminder_due_by(ews::date_time("2017-07-30T17:00:00Z"));

        auto assignee = ews::mailbox("assignee@example.com");

        auto id = service.create_item(
            task, ews::distinguished_folder_id(ews::standard_folder::tasks,
                                               assignee));

        // Now we can use the returned 'id' to access this item in the store
        // without explicitly referring to the assignee's mailbox anymore. This
        // is called implicit access in EWS documentation.

        // Now we gonna update the task because we already finished it
        // obviously.

        id = service.update_item(
            id, ews::property(ews::task_property_path::percent_complete, 100));

        // And better move it to the trash right away.

        service.delete_item(id, ews::delete_type::move_to_deleted_items);
    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}
