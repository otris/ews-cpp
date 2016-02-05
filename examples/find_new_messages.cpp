#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>

#include <string>
#include <iostream>
#include <ostream>
#include <exception>
#include <cstdlib>

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        const auto env = ews::test::get_from_environment();
        auto service = ews::service(env.server_uri,
                                    env.domain,
                                    env.username,
                                    env.password);

        // Get all unread messages from inbox
        ews::distinguished_folder_id inbox = ews::standard_folder::inbox;
        auto search_expression = ews::is_equal_to(
                                        ews::message_property_path::is_read,
                                        false);
        auto item_ids = service.find_item(inbox, search_expression);

        if (item_ids.empty())
        {
            std::cout << "No unread messages found!\n";
        }
        else
        {
            for (const auto& id : item_ids)
            {
                auto msg = service.get_message(id);
                std::cout << "Marking " << msg.get_subject() << " as read\n";
                auto prop = ews::property(ews::message_property_path::is_read,
                                          true);
                service.update_item(msg.get_item_id(), prop);
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

// vim:et ts=4 sw=4 noic cc=80
