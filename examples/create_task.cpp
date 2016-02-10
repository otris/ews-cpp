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

        auto start_time = ews::date_time("2015-01-16T12:00:00Z");
        auto end_time   = ews::date_time("2015-01-16T12:30:00Z");

        auto task = ews::task();
        task.set_subject("Something important to do");
        task.set_body(ews::body("Some descriptive body text"));
        task.set_start_date(start_time);
        task.set_due_date(end_time);
        task.set_reminder_enabled(true);
        task.set_reminder_due_by(start_time);

        auto item_id = service.create_item(task);
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
