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
        using xml_document = rapidxml::xml_document<>;

        // slang: 2013 SP1, not all properties included
        const auto xml = std::string(
            "<t:Task\n"
            "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n"
            "    <t:ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>\n"
            "    <t:ParentFolderId Id=\"qwertz\" ChangeKey=\"ztrewq\"/>\n"
            "    <t:ItemClass>IPM.Task</t:ItemClass>\n"
            "    <t:Subject>Write poem</t:Subject>\n"
            "    <t:Sensitivity>Normal</t:Sensitivity>\n"
            "    <t:Body BodyType=\"Text\" IsTruncated=\"false\"/>\n"
            "    <t:DateTimeReceived>2015-02-09T13:00:11Z</t:DateTimeReceived>\n"
            "    <t:Size>962</t:Size>\n"
            "    <t:Importance>Normal</t:Importance>\n"
            "    <t:IsSubmitted>false</t:IsSubmitted>\n"
            "    <t:IsDraft>false</t:IsDraft>\n"
            "    <t:IsFromMe>false</t:IsFromMe>\n"
            "    <t:IsResend>false</t:IsResend>\n"
            "    <t:IsUnmodified>false</t:IsUnmodified>\n"
            "    <t:DateTimeSent>2015-02-09T13:00:11Z</t:DateTimeSent>\n"
            "    <t:DateTimeCreated>2015-02-09T13:00:11Z</t:DateTimeCreated>\n"
            "    <t:DisplayCc/>\n"
            "    <t:DisplayTo/>\n"
            "    <t:HasAttachments>false</t:HasAttachments>\n"
            "    <t:Culture>en-US</t:Culture>\n"
            "    <t:EffectiveRights>\n"
            "            <t:CreateAssociated>false</t:CreateAssociated>\n"
            "            <t:CreateContents>false</t:CreateContents>\n"
            "            <t:CreateHierarchy>false</t:CreateHierarchy>\n"
            "            <t:Delete>true</t:Delete>\n"
            "            <t:Modify>true</t:Modify>\n"
            "            <t:Read>true</t:Read>\n"
            "            <t:ViewPrivateItems>true</t:ViewPrivateItems>\n"
            "    </t:EffectiveRights>\n"
            "    <t:LastModifiedName>Kwaltz</t:LastModifiedName>\n"
            "    <t:LastModifiedTime>2015-02-09T13:00:11Z</t:LastModifiedTime>\n"
            "    <t:IsAssociated>false</t:IsAssociated>\n"
            "    <t:Flag>\n"
            "            <t:FlagStatus>NotFlagged</t:FlagStatus>\n"
            "    </t:Flag>\n"
            "    <t:InstanceKey>AQAAAAAAARMBAAAAG4AqWQAAAAA=</t:InstanceKey>\n"
            "    <t:EntityExtractionResult/>\n"
            "    <t:ChangeCount>1</t:ChangeCount>\n"
            "    <t:IsComplete>false</t:IsComplete>\n"
            "    <t:IsRecurring>false</t:IsRecurring>\n"
            "    <t:PercentComplete>0</t:PercentComplete>\n"
            "    <t:Status>NotStarted</t:Status>\n"
            "    <t:StatusDescription>Not Started</t:StatusDescription>\n"
            "</t:Task>");
        std::vector<char> buf;
        std::copy(begin(xml), end(xml), std::back_inserter(buf));
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = doc.first_node();
        auto t = ews::task::from_xml_element(*node);

        EXPECT_STREQ("abcde", t.get_item_id().id().c_str());
        EXPECT_STREQ("edcba", t.get_item_id().change_key().c_str());
        EXPECT_STREQ("Write poem", t.get_subject().c_str());
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
