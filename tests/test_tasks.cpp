#include "fixtures.hpp"

#include <string>
#include <vector>
#include <algorithm>
#include <iterator>
#include <utility>

namespace tests
{
#pragma warning(suppress: 6262)
    TEST(OfflineTaskTest, FromXMLElement)
    {
        const auto task = make_fake_task();
        EXPECT_STREQ("abcde", task.get_item_id().id().c_str());
        EXPECT_STREQ("edcba", task.get_item_id().change_key().c_str());
        EXPECT_STREQ("Write poem", task.get_subject().c_str());
    }

    TEST_F(TaskTest, GetTaskWithInvalidIdThrows)
    {
        auto invalid_id = ews::item_id();
        EXPECT_THROW(service().get_task(invalid_id), ews::exchange_error);
    }

    TEST_F(TaskTest, GetTaskWithInvalidIdExceptionResponse)
    {
        auto invalid_id = ews::item_id();
        try
        {
            service().get_task(invalid_id);
            FAIL() << "Expected an exception";
        }
        catch (ews::exchange_error& exc)
        {
            EXPECT_EQ(ews::response_code::error_invalid_id_empty, exc.code());
            EXPECT_STREQ("ErrorInvalidIdEmpty", exc.what());
        }
    }

    TEST_F(TaskTest, CreateAndDelete)
    {
        auto start_time = ews::date_time("2015-01-17T12:00:00Z");
        auto end_time   = ews::date_time("2015-01-17T12:30:00Z");
        auto task = ews::task();
        task.set_subject("Something really important to do");
        task.set_body(ews::body("Some descriptive body text"));
        task.set_start_date(start_time);
        task.set_due_date(end_time);
        task.set_reminder_enabled(true);
        task.set_reminder_due_by(start_time);
        const auto item_id = service().create_item(task);

        auto created_task = service().get_task(item_id);
        // Check properties
        EXPECT_STREQ("Something really important to do",
                created_task.get_subject().c_str());
        EXPECT_EQ(start_time, created_task.get_start_date());
        EXPECT_EQ(end_time, created_task.get_due_date());
        EXPECT_TRUE(created_task.is_reminder_enabled());
        EXPECT_EQ(start_time, created_task.get_reminder_due_by());

        ASSERT_NO_THROW(
        {
            service().delete_task(
                    std::move(created_task), // Sink argument
                    ews::delete_type::hard_delete,
                    ews::affected_task_occurrences::all_occurrences);
        });
        EXPECT_STREQ("", created_task.get_subject().c_str());

        // Check if it is still in store
        EXPECT_THROW(service().get_task(item_id), ews::exchange_error);
    }

    TEST_F(TaskTest, FindTasks)
    {
        ews::distinguished_folder_id folder = ews::standard_folder::tasks;
        const auto initial_count = service().find_item(folder,
            ews::is_equal_to(ews::task_property_path::is_complete,
                             false)).size();

        auto start_time = ews::date_time("2015-05-29T17:00:00Z");
        auto end_time   = ews::date_time("2015-05-29T17:30:00Z");
        auto t = ews::task();
        t.set_subject("Feed the cat");
        t.set_body(ews::body("And don't forget to buy some Whiskas"));
        t.set_start_date(start_time);
        t.set_due_date(end_time);
        const auto item_id = service().create_item(t);
        t = service().get_task(item_id);
        ews::internal::on_scope_exit delete_item([&]
        {
            service().delete_task(std::move(t));
        });
        auto ids = service().find_item(
                folder,
                ews::is_equal_to(ews::task_property_path::is_complete, false));
        EXPECT_EQ(initial_count + 1, ids.size());
    }

    TEST_F(TaskTest, UpdateIsCompleteProperty)
    {
        auto get_milk = test_task();
        ASSERT_FALSE(get_milk.is_complete());
        auto prop = ews::property(ews::task_property_path::percent_complete,
                                  100);
        auto new_id = service().update_item(get_milk.get_item_id(), prop);
        get_milk = service().get_task(new_id);
        EXPECT_TRUE(get_milk.is_complete());
    }
}
