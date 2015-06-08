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

        ews::distinguished_folder_id contacts_folder =
            ews::standard_folder::contacts;
        ews::contact_property_path contact_property;
        auto restriction =
            ews::is_equal_to(contact_property.email_address_1, "john@doe.com");
        auto item_ids = service.find_item(contacts_folder, restriction);

        for (const auto& id : item_ids)
        {
            std::cout << id.to_xml() << std::endl;
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
