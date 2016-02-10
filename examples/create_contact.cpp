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
        const auto env = ews::test::environment();
        auto service = ews::service(env.server_uri,
                                    env.domain,
                                    env.username,
                                    env.password);

        auto contact = ews::contact();
        contact.set_given_name("Darkwing");
        contact.set_surname("Duck");
        contact.set_email_address_1(
                ews::mailbox("superhero@ducktales.com"));
        contact.set_job_title("Average Citizen");

        auto item_id = service.create_item(contact);
        std::cout << item_id.to_xml() << std::endl;
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
