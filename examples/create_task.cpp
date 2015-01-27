#include <ews/ews.hpp>

#include <string>
#include <iostream>
#include <ostream>
#include <exception>
#include <cstdlib>

namespace
{
    const std::string server_uri = "https://192.168.56.2/ews/Exchange.asmx";
    const std::string domain = "TEST";
    const std::string password = "12345aA!";
    const std::string username = "mini";
}

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        auto service = ews::service(server_uri, domain, username, password);

        auto start_time = ews::date_time("2015-01-16T12:00:00Z");
        auto end_time   = ews::date_time("2015-01-16T12:30:00Z");

        auto task = ews::task();
        task.set_subject("Something important to do");
        task.set_body(ews::body("Some descriptive body text"));
        task.set_owner("Mini Mouse");
        task.set_start_date(start_time);
        task.set_due_date(end_time);
        task.set_reminder_enabled(true);
        task.set_reminder_due_by(start_time);

        auto item_id = service.create_item(task);
        std::cout << item_id.xml() << std::endl;
    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}
