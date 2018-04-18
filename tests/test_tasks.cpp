
//   Copyright 2016 otris software AG
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
//
//   This project is hosted at https://github.com/otris

#include "fixtures.hpp"

#include <algorithm>
#include <iterator>
#include <string>
#include <utility>
#include <vector>

namespace tests
{
#ifdef _MSC_VER
#    pragma warning(suppress : 6262)
#endif
TEST(OfflineTaskTest, FromXMLElement)
{
    const auto task = make_fake_task();
    EXPECT_STREQ("abcde", task.get_item_id().id().c_str());
    EXPECT_STREQ("edcba", task.get_item_id().change_key().c_str());
    EXPECT_STREQ("Write poem", task.get_subject().c_str());
}

TEST(DateTimeTest, IsSet)
{
    const auto a = ews::date_time("2015-12-07T14:18:18.000Z");
    EXPECT_TRUE(a.is_set());
    const auto b = ews::date_time("");
    EXPECT_FALSE(b.is_set());
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
    }
}

TEST_F(TaskTest, CreateAndDelete)
{
    auto start_time = ews::date_time("2015-01-17T12:00:00Z");
    auto end_time = ews::date_time("2015-01-17T12:30:00Z");
    auto task = ews::task();
    task.set_subject("Something really important to do");
    task.set_body(ews::body("Some descriptive body text"));
    task.set_start_date(start_time);
    task.set_due_date(end_time);
    task.set_reminder_enabled(true);
    task.set_reminder_due_by(start_time);
    const auto item_id = service().create_item(task);

    auto created_task =
        service().get_task(item_id, ews::base_shape::all_properties);
    // Check properties
    EXPECT_STREQ("Something really important to do",
                 created_task.get_subject().c_str());
    EXPECT_EQ(start_time, created_task.get_start_date());
    EXPECT_EQ(end_time, created_task.get_due_date());
    EXPECT_TRUE(created_task.is_reminder_enabled());
    EXPECT_EQ(start_time, created_task.get_reminder_due_by());

    ASSERT_NO_THROW({
        service().delete_task(std::move(created_task), // Sink argument
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
    const auto initial_count =
        service()
            .find_item(folder, ews::is_equal_to(
                                   ews::task_property_path::is_complete, false))
            .size();

    auto start_time = ews::date_time("2015-05-29T17:00:00Z");
    auto end_time = ews::date_time("2015-05-29T17:30:00Z");
    auto t = ews::task();
    t.set_subject("Feed the cat");
    t.set_body(ews::body("And don't forget to buy some Whiskas"));
    t.set_start_date(start_time);
    t.set_due_date(end_time);
    const auto item_id = service().create_item(t);
    t = service().get_task(item_id);
    ews::internal::on_scope_exit delete_item(
        [&] { service().delete_task(std::move(t)); });
    auto ids = service().find_item(
        folder, ews::is_equal_to(ews::task_property_path::is_complete, false));
    EXPECT_EQ(initial_count + 1, ids.size());
}

TEST_F(TaskTest, ActualWorkPropertyInitialValue)
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

TEST_F(TaskTest, UpdateActualWorkProperty)
{
    auto task = test_task();
    auto prop = ews::property(ews::task_property_path::actual_work, 42);
    auto new_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(new_id, ews::base_shape::all_properties);
    EXPECT_EQ(42, task.get_actual_work());

    prop = ews::property(ews::task_property_path::actual_work, 1729);
    new_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(new_id, ews::base_shape::all_properties);
    EXPECT_EQ(1729, task.get_actual_work());
}

TEST_F(TaskTest, AssignedTimePropertyInitialValue)
{
    auto task = ews::task();
    EXPECT_FALSE(task.get_assigned_time().is_set());
}

TEST_F(TaskTest, BillingInformationPropertyInitialValue)
{
    auto task = ews::task();
    auto billing_information = task.get_billing_information();
    EXPECT_TRUE(billing_information.empty());
}

TEST_F(TaskTest, SetBillingInformationProperty)
{
    auto task = ews::task();
    task.set_billing_information("Bank transfer to Nigeria National Bank");
    EXPECT_STREQ("Bank transfer to Nigeria National Bank",
                 task.get_billing_information().c_str());
}

TEST_F(TaskTest, UpdateBillingInformationProperty)
{
    auto task = test_task();
    auto prop = ews::property(ews::task_property_path::billing_information,
                              "Billing Information Test 1");
    auto new_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Billing Information Test 1",
                 task.get_billing_information().c_str());

    prop = ews::property(ews::task_property_path::billing_information,
                         "Billing Information Test 2");
    new_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Billing Information Test 2",
                 task.get_billing_information().c_str());
}

TEST_F(TaskTest, ChangeCountPropertyInitialValue)
{
    auto task = ews::task();
    EXPECT_EQ(0, task.get_change_count());
}

// TODO: ChangeCount - more tests as soon as we support delegation

TEST_F(TaskTest, CompaniesPropertyInitialValue)
{
    auto task = ews::task();
    auto companies = task.get_companies();
    EXPECT_TRUE(companies.empty());
}

TEST_F(TaskTest, SetCompaniesProperty)
{
    std::vector<std::string> companies;
    companies.push_back("Tic Tric Tac Inc.");
    auto task = ews::task();
    task.set_companies(companies);
    ASSERT_EQ(1U, task.get_companies().size());
    EXPECT_STREQ("Tic Tric Tac Inc.", task.get_companies()[0].c_str());
}

TEST_F(TaskTest, UpdateCompaniesProperty)
{
    std::vector<std::string> companies;
    companies.push_back("Tic Tric Tac Inc.");
    auto task = test_task();
    auto prop = ews::property(ews::task_property_path::companies, companies);
    auto new_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(new_id, ews::base_shape::all_properties);
    ASSERT_EQ(1U, task.get_companies().size());
    EXPECT_STREQ("Tic Tric Tac Inc.", task.get_companies()[0].c_str());
}

TEST_F(TaskTest, CompleteDatePropertyInitialValue)
{
    auto task = ews::task();
    auto complete_date = task.get_complete_date();
    EXPECT_TRUE(complete_date.to_string().empty());
}

TEST_F(TaskTest, ContactsPropertyInitialValue)
{
    auto task = ews::task();
    auto associated_contacts = task.get_contacts();
    EXPECT_TRUE(associated_contacts.empty());
}

TEST_F(TaskTest, SetContactsProperty)
{
    auto contacts = std::vector<std::string>();
    contacts.push_back("Edgar Allan Poe");
    contacts.push_back("Ernest Hemingway");
    contacts.push_back("W. Somerset Maugham");

    auto task = ews::task();
    task.set_contacts(contacts);
    contacts = task.get_contacts();
    ASSERT_EQ(3U, contacts.size());
    EXPECT_STREQ("Edgar Allan Poe", contacts[0].c_str());
    EXPECT_STREQ("Ernest Hemingway", contacts[1].c_str());
    EXPECT_STREQ("W. Somerset Maugham", contacts[2].c_str());
}

TEST_F(TaskTest, UpdateContactsProperty)
{
    auto contacts = std::vector<std::string>();
    contacts.push_back("T. E. Lawrence");
    contacts.push_back("Dick Yates");

    auto task = test_task();
    auto prop = ews::property(ews::task_property_path::contacts, contacts);
    auto new_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(new_id, ews::base_shape::all_properties);
    contacts = task.get_contacts();
    ASSERT_EQ(2U, contacts.size());
    EXPECT_STREQ("T. E. Lawrence", contacts[0].c_str());
    EXPECT_STREQ("Dick Yates", contacts[1].c_str());
}

TEST_F(TaskTest, DelegationStatePropertyInitialValue)
{
    auto task = ews::task();
    // TODO: check if this is reasonable
    EXPECT_EQ(ews::delegation_state::no_match, task.get_delegation_state());
}

TEST_F(TaskTest, DelegatorPropertyInitialValue)
{
    auto task = ews::task();
    EXPECT_TRUE(task.get_delegator().empty());
}

// TODO: SetDelegatorProperty
// TODO: UpdateDelegatorProperty

// TODO: DueDatePropertyInitialValue
// TODO: SetDueDateProperty
// TODO: UpdateDueDateProperty

// TODO: IsAssignmentEditablePropertyInitialValue

TEST_F(TaskTest, UpdateIsCompleteProperty)
{
    auto get_milk = test_task();

    auto complete_date = get_milk.get_complete_date();
    ASSERT_FALSE(get_milk.is_complete());
    auto prop = ews::property(ews::task_property_path::percent_complete, 100);
    auto new_id = service().update_item(get_milk.get_item_id(), prop);
    get_milk = service().get_task(new_id, ews::base_shape::all_properties);

    EXPECT_TRUE(get_milk.is_complete());
    get_milk.get_complete_date();
}

// TODO: IsRecurringPropertyInitialValue
// TODO: more tests when we have task recurrences support

// TODO: IsTeamTaskPropertyInitialValue

TEST_F(TaskTest, MileagePropertyInitialValue)
{
    auto task = ews::task();
    auto str = task.get_mileage();
    EXPECT_TRUE(str.empty());
}

TEST_F(TaskTest, SetMileageProperty)
{
    auto task = ews::task();
    task.set_mileage("Thousands and thousands of parsecs");
    EXPECT_STREQ("Thousands and thousands of parsecs",
                 task.get_mileage().c_str());
}

TEST_F(TaskTest, UpdateMileageProperty)
{
    auto task = test_task();
    auto prop = ews::property(ews::task_property_path::mileage,
                              "Thousands and thousands of parsecs");
    auto new_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Thousands and thousands of parsecs",
                 task.get_mileage().c_str());

    prop = ews::property(ews::task_property_path::mileage, "A few steps");
    new_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("A few steps", task.get_mileage().c_str());
}

TEST_F(TaskTest, PercentCompletePropertyInitialValue)
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

TEST_F(TaskTest, UpdatePercentCompleteProperty)
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

// TODO: StartDatePropertyInitialValue
// TODO: SetStartDateProperty
// TODO: UpdateStartDateProperty

// TODO: StatusPropertyInitialValue
// TODO: SetStatusProperty
// TODO: UpdateProperty

// TODO: StatusDescriptionPropertyInitialValue

TEST_F(TaskTest, TotalWorkPropertyInitialValue)
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

TEST_F(TaskTest, UpdateTotalWorkProperty)
{
    auto task = test_task();
    auto prop = ews::property(ews::task_property_path::total_work, 3000);
    auto new_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(new_id, ews::base_shape::all_properties);
    EXPECT_EQ(3000, task.get_total_work());

    prop = ews::property(ews::task_property_path::total_work, 6000);
    new_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(new_id, ews::base_shape::all_properties);
    EXPECT_EQ(6000, task.get_total_work());
}
} // namespace tests
