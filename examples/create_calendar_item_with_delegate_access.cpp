
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

        // This example shows how to access another user's mailbox using EWS
        // delegate access. In this particular example, we are going to create
        // a single appointment on behalf of the mailbox's owner.
        //
        // Two users are involved: owner@example.com is the owner of the
        // calendar folder. He has granted user delegate@example.com author
        // level permissions on his calendar folder.
        //
        // To find out how you can add a delegate to a mailbox, see
        // add_delegate.cpp example.

        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri, env.domain, env.username,
                                    env.password);

        auto owner = ews::mailbox("owner@example.com");

        auto appointment = ews::calendar_item();
        appointment.set_subject("otris Kunden- und Interessenten-Forum 2017 ");
        appointment.set_body(
            ews::body("Wir sehen uns hier: "
                      "otris.de/veranstaltungen/verantwortung-treffen/"));
        appointment.set_start(ews::date_time("2017-09-14T10:00:00+02:00"));
        appointment.set_end(ews::date_time("2017-09-14T16:30:00+02:00"));

        auto id = service.create_item(
            appointment, ews::send_meeting_invitations::send_to_none,
            ews::distinguished_folder_id(ews::standard_folder::calendar,
                                         owner));

        // Now we can use the returned 'id' to access this item in the store
        // without explicitly referring to the owner's mailbox anymore. This is
        // because the fact that this item is part of a different mailbox is
        // encoded in the 'id' itself. This is called implicit access
        // in EWS documentation.
    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}
