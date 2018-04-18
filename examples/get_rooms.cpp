
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

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri, env.domain, env.username,
                                    env.password);

        auto room_lists = service.get_room_lists();
        if (room_lists.empty())
        {
            std::cout << "There are no room lists configured\n";
        }
        else
        {
            for (const auto& room_list : room_lists)
            {
                std::cout << "The room list " << room_list.name()
                          << " contains the following rooms:\n";
                auto rooms = service.get_rooms(room_list);
                if (rooms.empty())
                {
                    std::cout << "This room list does not contain any rooms\n";
                }
                else
                {
                    for (const auto& room : rooms)
                    {
                        std::cout << room.name() << "\n";
                    }
                }
            }
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
