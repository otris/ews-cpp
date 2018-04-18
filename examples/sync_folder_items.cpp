
//   Copyright 2018 otris software AG
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

// EWS provides two methods that allow you to discover changes that have
// occurred in your mailbox from some starting point (indicated by a
// sync_state) until the time the request is made:
//
// - sync_folder_hierarchy is used to watch for changes made to your mailbox's
//   folders
// - sync_folder_items is used to determine changes to the contents of a single
//   folder
//
// This example is about the latter.

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri, env.domain, env.username,
                                    env.password);

        ews::distinguished_folder_id folder = ews::standard_folder::contacts;

        // Initial synchronization call
        auto result = service.sync_folder_items(folder);

        // Store sync_state for further synchronization calls
        const auto sync_state = result.get_sync_state();

        // Output all existing items
        for (const auto& item : result.get_created_items())
        {
            std::cout << item.id() << "\n";
        }

        // Create test contact
        auto contact = ews::contact();
        contact.set_given_name("Darkwing");
        contact.set_surname("Duck");
        contact.set_email_address(
            ews::email_address(ews::email_address::key::email_address_1,
                               "superhero@ducktales.com"));
        contact.set_job_title("Average Citizen");
        auto item_id = service.create_item(contact);

        // Follow-up synchronization call
        result = service.sync_folder_items(folder, sync_state);

        // Output all newly created items
        for (const auto& item : result.get_created_items())
        {
            std::cout << item.id() << "\n";
        }

        // TODO: Show that updated items are synced
        // But service.update_item(...) creates a new item

        // Delete test item
        service.delete_item(item_id);

        // Follow-up synchronization call
        result = service.sync_folder_items(folder, sync_state);

        // Output all deleted items
        for (const auto& item : result.get_deleted_items())
        {
            std::cout << item.id() << "\n";
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
