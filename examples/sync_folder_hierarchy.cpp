
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

        ews::distinguished_folder_id folder = ews::standard_folder::inbox;

        // Initial synchronization call
        auto result = service.sync_folder_hierarchy(folder);

        // Store sync_state for further synchronization calls
        const auto sync_state = result.get_sync_state();

        // Output all existing folders
        for (const auto& folder : result.get_created_folders())
        {
            std::cout << folder.get_display_name() << "\n";
        }

        // Create test folder
        ews::folder new_folder;
        new_folder.set_display_name("Duck Cave");
        auto new_folder_id = service.create_folder(new_folder, folder);

        // Follow-up synchronization call
        result = service.sync_folder_hierarchy(folder, sync_state);

        // Output all newly created folders
        for (const auto& folder : result.get_created_folders())
        {
            std::cout << folder.get_display_name() << "\n";
        }

        // Delete test item
        service.delete_folder(new_folder_id);

        // Follow-up synchronization call
        result = service.sync_folder_hierarchy(folder, sync_state);

        // Output all deleted folders
        for (const auto& folder : result.get_deleted_folder_ids())
        {
            std::cout << folder.id() << "\n";
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
