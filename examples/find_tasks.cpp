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

        ews::distinguished_folder_id tasks_folder = ews::standard_folder::tasks;
        auto item_ids = service.find_item(tasks_folder,
                ews::is_equal_to(ews::task_property_path::is_complete, false));

        if (item_ids.empty())
        {
            std::cout << "Nothing to do. Get some beer with your friends!\n";
        }
        else
        {
            for (const auto& id : item_ids)
            {
                auto task = service.get_task(id);
                std::cout << task.get_subject() << std::endl;
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
