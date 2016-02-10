#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>

#include <string>
#include <iostream>
#include <ostream>
#include <exception>
#include <cstdlib>

using ews::internal::http_request;

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        const auto env = ews::test::get_from_environment();
        ews::basic_credentials credentials(env.autodiscover_smtp_address,
                                           env.autodiscover_password);

        auto ews_url =
            ews::get_exchange_web_services_url<http_request>(
                                    env.autodiscover_smtp_address,
                                    ews::autodiscover_protocol::internal,
                                    credentials);

        std::cout << ews_url << std::endl;
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
