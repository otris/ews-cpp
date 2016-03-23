#include "fixtures.hpp"
#include <ews/ews.hpp>
#include <vector>
#include <cstring>

using ews::item_id;
typedef rapidxml::xml_document<> xml_document;

namespace tests
{
    TEST(ItemIdTest, ConstructWithIdOnly)
    {
        auto a = item_id("abcde");
        EXPECT_STREQ("abcde", a.id().c_str());
        EXPECT_STREQ("", a.change_key().c_str());
    }

    TEST(ItemIdTest, ConstructWithIdAndChangeKey)
    {
        auto a = item_id("abcde", "edcba");
        EXPECT_STREQ("abcde", a.id().c_str());
        EXPECT_STREQ("edcba", a.change_key().c_str());
    }

    TEST(ItemIdTest, DefaultConstruction)
    {
        auto a = item_id();
        EXPECT_FALSE(a.valid());
        EXPECT_STREQ("", a.id().c_str());
        EXPECT_STREQ("", a.change_key().c_str());
    }

#pragma warning(suppress: 6262)
    TEST(ItemIdTest, FromXMLNode)
    {
        char buf[] = "<ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>";
        xml_document doc;
        doc.parse<0>(buf);
        auto node = doc.first_node();
        auto a = item_id::from_xml_element(*node);
        EXPECT_STREQ("abcde", a.id().c_str());
        EXPECT_STREQ("edcba", a.change_key().c_str());
    }

    TEST(ItemIdTest, ToXMLWithNamespace)
    {
        const char* expected = "<t:ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>";
        const auto a = item_id("abcde", "edcba");
        EXPECT_STREQ(expected, a.to_xml().c_str());
    }

#pragma warning(suppress: 6262)
    TEST(ItemIdTest, FromAndToXMLRoundTrip)
    {
        const char* xml = "<t:ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>";
        std::vector<char> buf(xml, xml + std::strlen(xml));
        buf.push_back('\0');
        xml_document doc;
        doc.parse<rapidxml::parse_no_namespace>(&buf[0]);
        auto node = doc.first_node();
        auto obj = item_id::from_xml_element(*node);
        EXPECT_STREQ(xml, obj.to_xml().c_str());
    }

    TEST(BodyTest, DefaultConstruction)
    {
        auto b = ews::body();
        EXPECT_EQ(ews::body_type::plain_text, b.type());
        EXPECT_FALSE(b.is_truncated());
    }

    TEST(BodyTest, PlainTextToXML)
    {
        auto b = ews::body("Here is some plain text",
                           ews::body_type::plain_text);
        EXPECT_EQ(ews::body_type::plain_text, b.type());
        EXPECT_FALSE(b.is_truncated());
        const char* const expected =
            "<t:Body BodyType=\"Text\">Here is some plain text</t:Body>";
        EXPECT_STREQ(expected, b.to_xml().c_str());
    }

    TEST(BodyTest, MakeSureHTMLIsWrappedWithCDATA)
    {
        auto b = ews::body("<b>Here is some HTML</b>", ews::body_type::html);
        EXPECT_EQ(ews::body_type::html, b.type());
        EXPECT_FALSE(b.is_truncated());
        const char* const expected =
            "<t:Body BodyType=\"HTML\"><![CDATA[<b>Here is some HTML</b>]]></t:Body>";
        EXPECT_STREQ(expected, b.to_xml().c_str());
    }

    TEST(PropertyPathTest, ConstructFromURI)
    {
        ews::property_path path = "";

        path = "folder:FolderId";
        EXPECT_EQ("folder:FolderId", path);
        EXPECT_STREQ("FolderId", path.property_name().c_str());
        EXPECT_STREQ("Folder", path.class_name().c_str());

        path = "item:DisplayCc";
        EXPECT_EQ("item:DisplayCc", path);
        EXPECT_STREQ("DisplayCc", path.property_name().c_str());
        EXPECT_STREQ("Item", path.class_name().c_str());

        path = "message:ToRecipients";
        EXPECT_EQ("message:ToRecipients", path);
        EXPECT_STREQ("ToRecipients", path.property_name().c_str());
        EXPECT_STREQ("Message", path.class_name().c_str());

        path = "meeting:IsOutOfDate";
        EXPECT_EQ("meeting:IsOutOfDate", path);
        EXPECT_STREQ("IsOutOfDate", path.property_name().c_str());
        EXPECT_STREQ("Meeting", path.class_name().c_str());

        path = "meetingRequest:MeetingRequestType";
        EXPECT_EQ("meetingRequest:MeetingRequestType", path);
        EXPECT_STREQ("MeetingRequestType", path.property_name().c_str());
        EXPECT_STREQ("MeetingRequest", path.class_name().c_str());

        path = "calendar:Start";
        EXPECT_EQ("calendar:Start", path);
        EXPECT_STREQ("Start", path.property_name().c_str());
        EXPECT_STREQ("CalendarItem", path.class_name().c_str());

        path = "task:AssignedTime";
        EXPECT_EQ("task:AssignedTime", path);
        EXPECT_STREQ("AssignedTime", path.property_name().c_str());
        EXPECT_STREQ("Task", path.class_name().c_str());

        path = "contacts:Children";
        EXPECT_EQ("contacts:Children", path);
        EXPECT_STREQ("Children", path.property_name().c_str());
        EXPECT_STREQ("Contact", path.class_name().c_str());

        path = "distributionlist:Members";
        EXPECT_EQ("distributionlist:Members", path);
        EXPECT_STREQ("Members", path.property_name().c_str());
        EXPECT_STREQ("DistributionList", path.class_name().c_str());

        path = "postitem:PostedTime";
        EXPECT_EQ("postitem:PostedTime", path);
        EXPECT_STREQ("PostedTime", path.property_name().c_str());
        EXPECT_STREQ("PostItem", path.class_name().c_str());

        path = "conversation:ConversationId";
        EXPECT_EQ("conversation:ConversationId", path);
        EXPECT_STREQ("ConversationId", path.property_name().c_str());
        EXPECT_STREQ("Conversation", path.class_name().c_str());
    }

    TEST(PropertyPathTest, ClassNameThrowsOnInvalidURI)
    {
        ews::property_path path = "random:string";
        EXPECT_THROW(path.class_name(), ews::exception);
    }

    TEST(PropertyPathTest, ClassNameThrowsOnInvalidURIWhat)
    {
        ews::property_path path = "some:string";
        try
        {
            path.class_name();
            FAIL() << "Expected exception to be raised";
        }
        catch (ews::exception& exc)
        {
            EXPECT_STREQ("Unknown property path", exc.what());
        }
    }

    TEST(OfflineItemTest, DefaultConstruction)
    {
        auto i = ews::item();
        EXPECT_TRUE(i.get_mime_content().none());
        EXPECT_STREQ("", i.get_subject().c_str());
        EXPECT_FALSE(i.get_item_id().valid());
    }

    TEST_F(ItemTest, NoMimeContentIfNotRequested)
    {
        auto contact = ews::contact();
        auto& s = service();
        const auto item_id = s.create_item(contact);
        ews::internal::on_scope_exit remove_contact([&]
        {
            s.delete_contact(std::move(contact));
        });
        contact = s.get_contact(item_id);
        EXPECT_TRUE(contact.get_mime_content().none());
    }

    TEST_F(ItemTest, GetMimeContentProperty)
    {
        auto contact = ews::contact();
        auto& s = service();
        const auto item_id = s.create_item(contact);
        ews::internal::on_scope_exit remove_contact([&]
        {
            s.delete_contact(std::move(contact));
        });
        auto additional_properties = std::vector<ews::property_path>();
        additional_properties.push_back(ews::item_property_path::mime_content);
        contact = s.get_contact(item_id, additional_properties);
        EXPECT_FALSE(contact.get_mime_content().none());
    }

    TEST(OfflineItemTest, GetParentFolderIdProperty)
    {
        const auto task = make_fake_task();
        auto parent_folder_id = task.get_parent_folder_id();
        EXPECT_TRUE(parent_folder_id.valid());
        EXPECT_STREQ("qwertz", parent_folder_id.id().c_str());
        EXPECT_STREQ("ztrewq", parent_folder_id.change_key().c_str());
    }

    TEST(OfflineItemTest, GetItemClassProperty)
    {
        const auto task = make_fake_task();
        const auto item_class = task.get_item_class();
        EXPECT_STREQ("IPM.Task", item_class.c_str());
    }

    TEST(OfflineItemTest, GetSensitivityProperty)
    {
        const auto task = make_fake_task();
        EXPECT_EQ(ews::sensitivity::confidential, task.get_sensitivity());
    }

    TEST(OfflineItemTest, GetSensitivityPropertyDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_EQ(ews::sensitivity::normal, task.get_sensitivity());
    }

    TEST(OfflineItemTest, SetSensitivity)
    {
        auto task = ews::task();
        task.set_sensitivity(ews::sensitivity::personal);
        EXPECT_EQ(ews::sensitivity::personal, task.get_sensitivity());
    }

    TEST_F(ItemTest, UpdateSensitivityProperty)
    {
        auto task = ews::task();
        task.set_sensitivity(ews::sensitivity::personal);
        auto item_id = service().create_item(task);
        ews::internal::on_scope_exit remove_task([&]
        {
            service().delete_task(std::move(task));
        });
        task = service().get_task(item_id);
        EXPECT_EQ(ews::sensitivity::personal, task.get_sensitivity());

        auto prop = ews::property(ews::item_property_path::sensitivity,
                                  ews::sensitivity::confidential);
        item_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(item_id);
        EXPECT_EQ(ews::sensitivity::confidential, task.get_sensitivity());
    }

    TEST(OfflineItemTest, GetAndSetBodyProperty)
    {
        auto item = ews::item();

        // TODO: better: what to do!? EXPECT_FALSE(item.get_body().none());
        // boost::optional desperately missing

        auto original =
            ews::body("<p>Some of the finest Vogon poetry</p>",
                      ews::body_type::html);
        item.set_body(original);

        auto actual = item.get_body();
        EXPECT_EQ(original.type(), actual.type());
        EXPECT_EQ(original.is_truncated(), actual.is_truncated());
        EXPECT_STREQ(original.content().c_str(),
                     actual.content().c_str());
    }

    // TODO: Attachments-Test

    TEST(OfflineItemTest, GetDateTimeReceivedProperty)
    {
        const auto task = make_fake_task();
        EXPECT_EQ(ews::date_time("2015-02-09T13:00:11Z"), task.get_date_time_received());
    }

    TEST(OfflineItemTest, GetDateTimeReceivedPropertyDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_EQ(ews::date_time(), task.get_date_time_received());
    }

    TEST_F(ItemTest, GetDateTimeReceivedProperty)
    {
        auto task = ews::task();
        auto item_id = service().create_item(task);
        ews::internal::on_scope_exit remove_task([&]
        {
            service().delete_task(std::move(task));
        });
        task = service().get_task(item_id);
        EXPECT_TRUE(task.get_date_time_received().is_set());
    }

    TEST(OfflineItemTest, GetSizeProperty)
    {
        const auto task = make_fake_task();
        EXPECT_EQ(962U, task.get_size());
    }

    TEST(OfflineItemTest, GetSizePropertyDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_EQ(0U, task.get_size());
    }

    TEST(OfflineItemTest, SetCategoriesPropertyDefaultConstructed)
    {
      auto task = ews::task();
      std::vector<std::string> categories;
      categories.push_back("ham");
      categories.push_back("spam");
      task.set_categories(categories);

      ASSERT_EQ(2, task.get_categories().size());
      EXPECT_EQ("ham", task.get_categories()[0]);
      EXPECT_EQ("spam", task.get_categories()[1]);
    }

    TEST(OfflineItemTest, GetCategoriesProperty)
    {
        const auto task = make_fake_task();
        EXPECT_EQ(0U, task.get_categories().size());
    }

    TEST(OfflineItemTest, GetCategoriesPropertyDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_EQ(0U, task.get_categories().size());
    }

    TEST_F(ItemTest, GetCategoriesProperty)
    {
        auto task = ews::task();
        std::vector<std::string> categories;
        categories.push_back("ham");
        categories.push_back("spam");
        task.set_categories(categories);
        auto item_id = service().create_item(task);
        ews::internal::on_scope_exit remove_task([&]
        {
            service().delete_task(std::move(task));
        });
        task = service().get_task(item_id);
        ASSERT_EQ(2, task.get_categories().size());
        EXPECT_EQ("ham", task.get_categories()[0]);
        EXPECT_EQ("spam", task.get_categories()[1]);

        // update
        std::vector<std::string> prop_categories;
        prop_categories.push_back("note");
        prop_categories.push_back("info");
        auto prop = ews::property(ews::item_property_path::categories,
                                  prop_categories);
        item_id = service().update_item(task.get_item_id(), prop);
        task = service().get_task(item_id);
        ASSERT_EQ(2, task.get_categories().size());
        EXPECT_EQ("note", task.get_categories()[0]);
        EXPECT_EQ("info", task.get_categories()[1]);
    }

    TEST(OfflineItemTest, GetImportanceProperty)
    {
        const auto task = make_fake_task();
        EXPECT_EQ(ews::importance::normal, task.get_importance());
    }

    TEST(OfflineItemTest, GetImportancePropertyDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_EQ(ews::importance::normal, task.get_importance());
    }

    TEST(OfflineItemTest, IsSubmittedProperty)
    {
        const auto task = make_fake_task();
        EXPECT_FALSE(task.is_submitted());
    }

    TEST(OfflineItemTest, IsSubmittedDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_FALSE(task.is_submitted());
    }

    TEST(OfflineItemTest, IsDraftProperty)
    {
        const auto task = make_fake_task();
        EXPECT_FALSE(task.is_draft());
    }

    TEST(OfflineItemTest, IsDraftPropertyDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_FALSE(task.is_draft());
    }

    TEST(OfflineItemTest, IsFromMeProperty)
    {
        const auto task = make_fake_task();
        EXPECT_FALSE(task.is_from_me());
    }

    TEST(OfflineItemTest, IsFromMeDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_FALSE(task.is_from_me());
    }

    TEST(OfflineItemTest, IsResendProperty)
    {
        const auto task = make_fake_task();
        EXPECT_FALSE(task.is_resend());
    }

    TEST(OfflineItemTest, IsResendDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_FALSE(task.is_resend());
    }

    TEST(OfflineItemTest, IsUnmodifiedProperty)
    {
        const auto task = make_fake_task();
        EXPECT_FALSE(task.is_unmodified());
    }

    TEST(OfflineItemTest, IsUnmodifiedDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_FALSE(task.is_unmodified());
    }

    TEST(OfflineItemTest, GetDateTimeSentProperty)
    {
        const auto task = make_fake_task();
        EXPECT_EQ(ews::date_time("2015-02-09T13:00:11Z"), task.get_date_time_sent());
    }

    TEST(OfflineItemTest, GetDateTimeSentPropertyDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_EQ(ews::date_time(), task.get_date_time_sent());
    }

    TEST_F(ItemTest, GetDateTimeSentProperty)
    {
        auto task = ews::task();
        auto item_id = service().create_item(task);
        ews::internal::on_scope_exit remove_task([&]
        {
            service().delete_task(std::move(task));
        });
        task = service().get_task(item_id);
        EXPECT_TRUE(task.get_date_time_sent().is_set());
    }

    TEST(OfflineItemTest, GetDateTimeCreatedProperty)
    {
        const auto task = make_fake_task();
        EXPECT_EQ(ews::date_time("2015-02-09T13:00:11Z"), task.get_date_time_created());
    }

    TEST(OfflineItemTest, GetDateTimeCreatedPropertyDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_EQ(ews::date_time(), task.get_date_time_sent());
    }

    TEST_F(ItemTest, GetDateTimeCreatedProperty)
    {
        auto task = ews::task();
        auto item_id = service().create_item(task);
        ews::internal::on_scope_exit remove_task([&]
        {
            service().delete_task(std::move(task));
        });
        task = service().get_task(item_id);
        EXPECT_TRUE(task.get_date_time_created().is_set());
    }

    TEST(OfflineItemTest, ReminderDueByPropertyDefaultConstructed)
    {
        auto task = ews::task();
        EXPECT_EQ(ews::date_time(), task.get_reminder_due_by());
        // set
        task.set_reminder_due_by(ews::date_time("2012-09-11T10:00:11Z"));
        EXPECT_EQ(ews::date_time("2012-09-11T10:00:11Z"), task.get_reminder_due_by());
        // update
        task.set_reminder_due_by(ews::date_time("2001-09-11T12:00:11Z"));
        EXPECT_EQ(ews::date_time("2001-09-11T12:00:11Z"), task.get_reminder_due_by());
    }

    TEST_F(ItemTest, ReminderDueByProperty)
    {
        auto task = ews::task();
        task.set_reminder_due_by(ews::date_time("2001-09-11T12:00:11Z"));
        auto item_id = service().create_item(task);
        ews::internal::on_scope_exit remove_task([&]
        {
            service().delete_task(std::move(task));
        });
        task = service().get_task(item_id);
        EXPECT_EQ(ews::date_time("2001-09-11T12:00:11Z"), task.get_reminder_due_by());
    }

    TEST(OfflineItemTest, ReminderMinutesBeforeStartPropertyDefaultConstructed)
    {
        auto task = ews::task();
        // empty without set
        ASSERT_EQ(0U, task.get_reminder_minutes_before_start());
        // set
        task.set_reminder_minutes_before_start(999);
        EXPECT_EQ(999, task.get_reminder_minutes_before_start());
        // update
        task.set_reminder_minutes_before_start(100);
        EXPECT_EQ(100, task.get_reminder_minutes_before_start());
    }

    TEST_F(ItemTest, ReminderMinutesBeforeStartProperty)
    {
        auto task = ews::task();
        // empty
        ASSERT_EQ(0, task.get_reminder_minutes_before_start());
        task.set_reminder_minutes_before_start(999);
        auto item_id = service().create_item(task);
        ews::internal::on_scope_exit remove_task([&]
        {
            service().delete_task(std::move(task));
        });
        task = service().get_task(item_id);
        EXPECT_EQ(999, task.get_reminder_minutes_before_start());
    }

    TEST(OfflineItemTest, DisplayCcPropertyDefaultConstructed)
    {
        auto task = ews::task();
        EXPECT_EQ("", task.get_display_cc());
    }

    TEST_F(ItemTest, DisplayCcProperty)
    {
        auto task = ews::task();
        auto item_id = service().create_item(task);
        ews::internal::on_scope_exit remove_task([&]
        {
            service().delete_task(std::move(task));
        });
        task = service().get_task(item_id);
        EXPECT_EQ("", task.get_display_cc());
        // TODO: more tests
    }

    TEST(OfflineItemTest, DisplayToPropertyDefaultConstructed)
    {
        auto task = ews::task();
        EXPECT_EQ("", task.get_display_to());
    }

    TEST_F(ItemTest, DisplayToProperty)
    {
        auto task = ews::task();
        auto item_id = service().create_item(task);
        ews::internal::on_scope_exit remove_task([&]
        {
            service().delete_task(std::move(task));
        });
        task = service().get_task(item_id);
        EXPECT_EQ("", task.get_display_to());
        // TODO: more tests
    }
}

// vim:et ts=4 sw=4 noic cc=80
