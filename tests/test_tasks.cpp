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

        auto complete_date = get_milk.get_complete_date();
        ASSERT_FALSE(get_milk.is_complete());
        auto prop = ews::property(ews::task_property_path::percent_complete,
                                  100);
        auto new_id = service().update_item(get_milk.get_item_id(), prop);
        get_milk = service().get_task(new_id);

        EXPECT_TRUE(get_milk.is_complete());
        get_milk.get_complete_date();
    }

    TEST_F(TaskTest, InitialActualWorkProperty)
    {
        auto task = ews::task();
        EXPECT_EQ(0, task.get_actual_work());
    }

    TEST_F(TaskTest, SetActualWorkProperty)
    {
        auto task = ews::task();
        task.set_actual_work(42);
        EXPECT_EQ(42, task.get_actual_work());
    }

    TEST_F(TaskTest, SetActualWorkPropertyServer)
    {
        auto task = test_task();
        auto prop = ews::property(ews::task_property_path::actual_work, 42);
        auto new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        EXPECT_EQ(42, task.get_actual_work());

        prop = ews::property(ews::task_property_path::actual_work, 1729);
        new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        EXPECT_EQ(1729, task.get_actual_work());
    }

    //TODO: assigned_time

    TEST_F(TaskTest, GetBillingInformationIsEmptyInitially)
    {
        auto task = test_task();
        auto billing_information = task.get_billing_information();
        EXPECT_TRUE(billing_information.empty());
    }

    TEST_F(TaskTest, SetBillingInformationProperty)
    {
        auto task = ews::task();
        task.set_billing_information("Billing Information Test");
        EXPECT_STREQ("Billing Information Test", task.get_billing_information().c_str());
    }

    TEST_F(TaskTest, SetBillingInformationPropertyServer)
    {
        auto task = test_task();
        auto prop = ews::property(ews::task_property_path::billing_information,
                                  "Billing Information Test 1");
        auto new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        EXPECT_STREQ("Billing Information Test 1",
                     task.get_billing_information().c_str());

        prop = ews::property(ews::task_property_path::billing_information,
                             "Billing Information Test 2");
        new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        EXPECT_STREQ("Billing Information Test 2",
                     task.get_billing_information().c_str());
    }

    // TODO: ChangeCount - Test with Delegation

    // TODO
    TEST_F(TaskTest, GetCompaniesIsEmptyInitially)
    {
        auto task = test_task();
        auto companies = task.get_companies();
        EXPECT_TRUE(companies.empty());
    }

    TEST_F(TaskTest, SetCompaniesProperty)
    {
        const auto comps = std::vector<std::string>{
            "Tic Tric Tac Inc.",
        };
        auto task = ews::task();
        task.set_companies(comps);
        ASSERT_EQ(1U, task.get_companies().size());
        EXPECT_STREQ("Tic Tric Tac Inc.", task.get_companies()[0].c_str());
    }

    TEST_F(TaskTest, SetCompaniesPropertyServer)
    {
        const auto comps = std::vector<std::string>{
            "Tic Tric Tac Inc.",
        };
        auto task = test_task();
        auto prop = ews::property(ews::task_property_path::companies, comps);
        auto new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        ASSERT_EQ(1U, task.get_companies().size());
        EXPECT_STREQ("Tic Tric Tac Inc.", task.get_companies()[0].c_str());
    }

    TEST_F(TaskTest, GetContactsIsEmptyInitially)
    {
        auto task = test_task();
        auto contacts = task.get_contacts();
        EXPECT_TRUE(contacts.empty());
    }

    // TODO: Contacts

    // TODO: DelegationState
    //TEST_F(TaskTest, GetDelegationState)
    //{
    //    auto task = test_task();
    //    EXPECT_EQ(ews::delegation_state::no_match, task.get_delegation_state());
    //}

    // TODO: Delegator

    // TODO: IsAssignmentEditable

    // TODO: IsRecurring

    // TODO: IsTeamTask

    TEST_F(TaskTest, GetMileageIsEmptyInitially)
    {
        auto task = test_task();
        auto mileage = task.get_mileage();
        EXPECT_TRUE(mileage.empty());
    }

    TEST_F(TaskTest, SetMileageProperty)
    {
        auto task = ews::task();
        task.set_mileage("Mileage Test");
        EXPECT_STREQ("Mileage Test", task.get_mileage().c_str());
    }

    TEST_F(TaskTest, SetMileagePropertyServer)
    {
        auto task = test_task();
        auto prop = ews::property(ews::task_property_path::mileage, "Mileage Test");
        auto new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        EXPECT_STREQ("Mileage Test", task.get_mileage().c_str());

        prop = ews::property(ews::task_property_path::mileage, "Mileage Test 2");
        new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        EXPECT_STREQ("Mileage Test 2", task.get_mileage().c_str());
    }

    TEST_F(TaskTest, InitialPercentCompleteProperty)
    {
        auto task = ews::task();
        EXPECT_EQ(0, task.get_percent_complete());
    }

    TEST_F(TaskTest, SetPercentCompleteProperty)
    {
        auto task = ews::task();
        task.set_percent_complete(55);
        EXPECT_EQ(55, task.get_percent_complete());
    }

    TEST_F(TaskTest, SetPercentCompletePropertyServer)
    {
        auto task = test_task();
        auto prop = ews::property(ews::task_property_path::percent_complete, 55);
        auto new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        EXPECT_EQ(55, task.get_percent_complete());

        prop = ews::property(ews::task_property_path::percent_complete, 100);
        new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        EXPECT_EQ(100, task.get_percent_complete());
    }

    //TODO: Recurrence

    //TODO: Status

    //TODO: StatusDescription

    TEST_F(TaskTest, InitialTotalWorkProperty)
    {
        auto task = ews::task();
        EXPECT_EQ(0, task.get_total_work());
    }

    TEST_F(TaskTest, SetTotalWorkProperty)
    {
        auto task = ews::task();
        task.set_total_work(3000);
        EXPECT_EQ(3000, task.get_total_work());
    }

    TEST_F(TaskTest, SetTotalWorkPropertyServer)
    {
        auto task = test_task();
        auto prop = ews::property(ews::task_property_path::total_work, 3000);
        auto new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        EXPECT_EQ(3000, task.get_total_work());

        prop = ews::property(ews::task_property_path::total_work, 6000);
        new_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(new_id);
        EXPECT_EQ(6000, task.get_total_work());
    }
}
