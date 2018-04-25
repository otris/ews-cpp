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

#include <chrono>
#include <cstdlib>
#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <exception>
#include <iostream>
#include <ostream>
#include <string>
#include <thread>

// Note: this example still does not compile on Visual Studio 2017
namespace
{
template <class... Ts> struct overloaded : Ts...
{
    using Ts::operator()...;
};
template <class... Ts> overloaded(Ts...)->overloaded<Ts...>;
} // namespace

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri, env.domain, env.username,
                                    env.password);

        // Subscribe to all <CreatedEvents> in inbox
        auto sub_info = service.subscribe({ews::standard_folder::inbox},
                                          {ews::event_type::created_event}, 10);

        // Create and send a message to trigger an event
        auto message = ews::message();
        message.set_to_recipients({ews::mailbox("test2@otris.de")});
        service.create_item(message,
                            ews::message_disposition::send_and_save_copy);

        // Wait for the event to be created
        std::this_thread::sleep_for(std::chrono::seconds(5));

        // Get and check the created_event.
        // This way, we can get all the needed information from the events for
        // further handling.
        auto notification = service.get_events(sub_info.get_subscription_id(),
                                               sub_info.get_watermark());
        std::cout << "SubscriptionId: " << notification.subscription_id
                  << std::endl;
        std::cout << "MoreEvents: " << notification.more_events << std::endl;
        for (auto& e : notification.events)
        {
            std::visit(
                overloaded{[](auto&&) { /* do nothing */ },
                           [](ews::created_event arg) {
                               std::cout
                                   << "EventType: "
                                   << ews::internal::enum_to_str(arg.get_type())
                                   << std::endl;
                               std::cout << "Watermark: " << arg.get_watermark()
                                         << std::endl;
                               std::cout << "Timestamp: " << arg.get_timestamp()
                                         << std::endl;
                           }},
                e);
        }

        service.unsubscribe(sub_info.get_subscription_id());
    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}
