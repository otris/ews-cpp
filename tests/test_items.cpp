
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

#include <ews/ews.hpp>
#include <string.h>
#include <vector>

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

#ifdef _MSC_VER
#    pragma warning(suppress : 6262)
#endif
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

#ifdef _MSC_VER
#    pragma warning(suppress : 6262)
#endif
TEST(ItemIdTest, FromAndToXMLRoundTrip)
{
    const char* xml = "<t:ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>";
    std::vector<char> buf(xml, xml + strlen(xml));
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
    auto b = ews::body("Here is some plain text", ews::body_type::plain_text);
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
    const char* const expected = "<t:Body "
                                 "BodyType=\"HTML\"><![CDATA[<b>Here is "
                                 "some HTML</b>]]></t:Body>";
    EXPECT_STREQ(expected, b.to_xml().c_str());
}

TEST(PropertyPathTest, ConstructFromURI)
{
    ews::property_path path = ews::folder_property_path::folder_id;
    EXPECT_STREQ("<t:FieldURI FieldURI=\"folder:FolderId\"/>",
                 path.to_xml().c_str());

    path = "item:DisplayCc";
    EXPECT_STREQ("<t:FieldURI FieldURI=\"item:DisplayCc\"/>",
                 path.to_xml().c_str());

    path = "message:ToRecipients";
    EXPECT_STREQ("<t:FieldURI FieldURI=\"message:ToRecipients\"/>",
                 path.to_xml().c_str());

    path = "meeting:IsOutOfDate";
    EXPECT_STREQ("<t:FieldURI FieldURI=\"meeting:IsOutOfDate\"/>",
                 path.to_xml().c_str());

    path = "meetingRequest:MeetingRequestType";
    EXPECT_STREQ("<t:FieldURI FieldURI=\"meetingRequest:MeetingRequestType\"/>",
                 path.to_xml().c_str());

    path = "calendar:Start";
    EXPECT_STREQ("<t:FieldURI FieldURI=\"calendar:Start\"/>",
                 path.to_xml().c_str());

    path = "task:AssignedTime";
    EXPECT_STREQ("<t:FieldURI FieldURI=\"task:AssignedTime\"/>",
                 path.to_xml().c_str());

    path = "contacts:Children";
    EXPECT_STREQ("<t:FieldURI FieldURI=\"contacts:Children\"/>",
                 path.to_xml().c_str());

    path = "distributionlist:Members";
    EXPECT_STREQ("<t:FieldURI FieldURI=\"distributionlist:Members\"/>",
                 path.to_xml().c_str());

    path = "postitem:PostedTime";
    EXPECT_STREQ("<t:FieldURI FieldURI=\"postitem:PostedTime\"/>",
                 path.to_xml().c_str());

    path = "conversation:ConversationId";
    EXPECT_STREQ("<t:FieldURI FieldURI=\"conversation:ConversationId\"/>",
                 path.to_xml().c_str());
}

TEST(PropertyPathTest, ClassNameThrowsOnInvalidURIWhat)
{
    try
    {
        ews::property_path path = "some:string";
        FAIL() << "Expected exception to be raised";
    }
    catch (ews::exception& exc)
    {
        EXPECT_STREQ("Unknown property path", exc.what());
    }
}

TEST(IndexedPropertyPath, ToXML)
{
    ews::indexed_property_path path("contacts:PhoneNumber", "BusinessPhone");
    EXPECT_STREQ("<t:IndexedFieldURI FieldURI=\"contacts:PhoneNumber\" "
                 "FieldIndex=\"BusinessPhone\"/>",
                 path.to_xml().c_str());
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
    ews::internal::on_scope_exit remove_contact(
        [&] { s.delete_contact(std::move(contact)); });
    contact = s.get_contact(item_id);
    EXPECT_TRUE(contact.get_mime_content().none());
}

TEST_F(ItemTest, GetMimeContentProperty)
{
    auto contact = ews::contact();
    auto& s = service();
    const auto item_id = s.create_item(contact);
    ews::internal::on_scope_exit remove_contact(
        [&] { s.delete_contact(std::move(contact)); });
    auto additional_properties = std::vector<ews::property_path>();
    additional_properties.push_back(ews::item_property_path::mime_content);
    ews::item_shape shape(std::move(additional_properties));
    contact = s.get_contact(item_id, shape);
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
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
    task = service().get_task(item_id, ews::base_shape::all_properties);
    EXPECT_EQ(ews::sensitivity::personal, task.get_sensitivity());

    auto prop = ews::property(ews::item_property_path::sensitivity,
                              ews::sensitivity::confidential);
    item_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(item_id, ews::base_shape::all_properties);
    EXPECT_EQ(ews::sensitivity::confidential, task.get_sensitivity());
}

TEST(OfflineItemTest, GetAndSetBodyProperty)
{
    auto item = ews::item();

    auto original = ews::body("<p>Some of the finest Vogon poetry</p>",
                              ews::body_type::html);
    item.set_body(original);

    auto actual = item.get_body();
    EXPECT_EQ(original.type(), actual.type());
    EXPECT_EQ(original.is_truncated(), actual.is_truncated());
    EXPECT_STREQ(original.content().c_str(), actual.content().c_str());
}

TEST_F(ItemTest, BodyPropertyIsProperlyEscaped)
{
    auto task = ews::task();
    task.set_body(
        ews::body("some special character &", ews::body_type::plain_text));
    auto id = service().create_item(task);
    ews::internal::on_scope_exit cleanup([&] { service().delete_item(id); });

    ews::property prop(
        ews::item_property_path::body,
        ews::body("this should work too &", ews::body_type::plain_text));
    ews::update update(prop, ews::update::operation::set_item_field);
    id = service().update_item(id, update);
}

// TODO: Attachments-Test

TEST(OfflineItemTest, GetDateTimeReceivedProperty)
{
    const auto task = make_fake_task();
    EXPECT_EQ(ews::date_time("2015-02-09T13:00:11Z"),
              task.get_date_time_received());
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
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
    task = service().get_task(item_id, ews::base_shape::all_properties);
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

    ASSERT_EQ(2U, task.get_categories().size());
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
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
    task = service().get_task(item_id, ews::base_shape::all_properties);
    ASSERT_EQ(2U, task.get_categories().size());
    EXPECT_EQ("ham", task.get_categories()[0]);
    EXPECT_EQ("spam", task.get_categories()[1]);

    // update
    std::vector<std::string> prop_categories;
    prop_categories.push_back("note");
    prop_categories.push_back("info");
    auto prop =
        ews::property(ews::item_property_path::categories, prop_categories);
    item_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(item_id, ews::base_shape::all_properties);
    ASSERT_EQ(2U, task.get_categories().size());
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

TEST(OfflineItemTest, GetInReplyToPropertyDefaultConstructed)
{
    const auto task = ews::task();
    EXPECT_EQ("", task.get_in_reply_to());
}

TEST_F(ItemTest, GetInReplyToProperty)
{
    auto task = ews::task();
    ASSERT_EQ("", task.get_in_reply_to());

    auto item_id = service().create_item(task);
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
    task = service().get_task(item_id);

    // set
    auto prop = ews::property(ews::item_property_path::in_reply_to,
                              "nobody@noreply.com");
    item_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(item_id, ews::base_shape::all_properties);
    ASSERT_EQ("nobody@noreply.com", task.get_in_reply_to());

    // update
    prop = ews::property(ews::item_property_path::in_reply_to,
                         "somebody@noreply.com");
    item_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(item_id, ews::base_shape::all_properties);
    EXPECT_EQ("somebody@noreply.com", task.get_in_reply_to());
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

#ifdef EWS_HAS_FILESYSTEM_HEADER
TEST(OfflineItemTest, GetInternetMessageHeaders)
{
    auto message = make_fake_message();
    auto headers = message.get_internet_message_headers();
    ASSERT_FALSE(headers.empty());

    std::vector<ews::internet_message_header> expected_headers;
    expected_headers.push_back(ews::internet_message_header(
        "Received", "from duckburg2013.otris.de (192.168.4.234) "
                    "by duckburg2013.otris.de (192.168.4.234) "
                    "with Microsoft SMTP Server (TLS) id 15.0.847.32 "
                    "via Mailbox Transport; Sun, 7 Feb 2016 12:12:49 +0100"));
    expected_headers.push_back(
        ews::internet_message_header("MIME-Version", "1.0"));
    expected_headers.push_back(
        ews::internet_message_header("Date", "Sun, 7 Feb 2016 12:12:31 +0100"));
    expected_headers.push_back(
        ews::internet_message_header("Content-Type", "multipart/report"));
    expected_headers.push_back(
        ews::internet_message_header("X-MS-Exchange-Organization-SCL", "-1"));
    expected_headers.push_back(
        ews::internet_message_header("Content-Language", "en-US"));
    expected_headers.push_back(ews::internet_message_header(
        "Message-ID",
        "<28b94593-526c-42d8-b49b-257f04f15083@duckburg2013.otris.de>"));
    expected_headers.push_back(ews::internet_message_header(
        "In-Reply-To",
        "<c829d7b23a1b4c138c0b58d80b97b595@duckburg2013.otris.de>"));
    expected_headers.push_back(ews::internet_message_header(
        "References",
        "<c829d7b23a1b4c138c0b58d80b97b595@duckburg2013.otris.de>"));
    expected_headers.push_back(
        ews::internet_message_header("Thread-Topic", "Test mail"));
    expected_headers.push_back(ews::internet_message_header(
        "Thread-Index", "AQHRYAVXv4yEkT9GTECGcXS3Z6t3OJ8gcNbs"));
    expected_headers.push_back(
        ews::internet_message_header("Subject", "Undeliverable: Test mail"));
    expected_headers.push_back(
        ews::internet_message_header("Auto-Submitted", "auto-replied"));
    expected_headers.push_back(ews::internet_message_header(
        "X-MS-Exchange-Organization-AuthSource", "duckburg2013.otris.de"));
    expected_headers.push_back(ews::internet_message_header(
        "X-MS-Exchange-Organization-AuthAs", "Internal"));
    expected_headers.push_back(ews::internet_message_header(
        "X-MS-Exchange-Organization-AuthMechanism", "05"));
    expected_headers.push_back(ews::internet_message_header(
        "X-MS-Exchange-Organization-Network-Message-Id",
        "6b449cb6-88e2-4a17-acde-08d32faf931b"));
    expected_headers.push_back(
        ews::internet_message_header("Return-Path", "<>"));

    auto expected = expected_headers.cbegin();
    for (const auto& header : headers)
    {
        ASSERT_FALSE(header.get_name().empty());
        ASSERT_FALSE(header.get_value().empty());

        EXPECT_EQ(expected->get_name(), header.get_name());
        EXPECT_EQ(expected->get_value(), header.get_value());
        expected++;
    }
}
#endif // EWS_HAS_FILESYSTEM_HEADER

TEST(OfflineItemTest, GetDateTimeSentProperty)
{
    const auto task = make_fake_task();
    EXPECT_EQ(ews::date_time("2015-02-09T13:00:11Z"),
              task.get_date_time_sent());
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
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
    task = service().get_task(item_id, ews::base_shape::all_properties);
    EXPECT_TRUE(task.get_date_time_sent().is_set());
}

TEST(OfflineItemTest, GetDateTimeCreatedProperty)
{
    const auto task = make_fake_task();
    EXPECT_EQ(ews::date_time("2015-02-09T13:00:11Z"),
              task.get_date_time_created());
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
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
    task = service().get_task(item_id, ews::base_shape::all_properties);
    EXPECT_TRUE(task.get_date_time_created().is_set());
}

TEST(OfflineItemTest, ReminderDueByPropertyDefaultConstructed)
{
    auto task = ews::task();
    EXPECT_EQ(ews::date_time(), task.get_reminder_due_by());

    task.set_reminder_due_by(ews::date_time("2012-09-11T10:00:11Z"));
    EXPECT_EQ(ews::date_time("2012-09-11T10:00:11Z"),
              task.get_reminder_due_by());

    task.set_reminder_due_by(ews::date_time("2001-09-11T12:00:11Z"));
    EXPECT_EQ(ews::date_time("2001-09-11T12:00:11Z"),
              task.get_reminder_due_by());
}

TEST_F(ItemTest, ReminderDueByProperty)
{
    auto task = ews::task();
    task.set_reminder_due_by(ews::date_time("2001-09-11T12:00:11Z"));
    auto item_id = service().create_item(task);
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
    task = service().get_task(item_id, ews::base_shape::all_properties);
    EXPECT_EQ(ews::date_time("2001-09-11T12:00:11Z"),
              task.get_reminder_due_by());
}

TEST(OfflineItemTest, ReminderMinutesBeforeStartPropertyDefaultConstructed)
{
    auto task = ews::task();
    // empty without set
    ASSERT_EQ(0U, task.get_reminder_minutes_before_start());
    // set
    task.set_reminder_minutes_before_start(999);
    EXPECT_EQ(999U, task.get_reminder_minutes_before_start());
    // update
    task.set_reminder_minutes_before_start(100);
    EXPECT_EQ(100U, task.get_reminder_minutes_before_start());
}

TEST_F(ItemTest, ReminderMinutesBeforeStartProperty)
{
    auto task = ews::task();
    // empty
    ASSERT_EQ(0U, task.get_reminder_minutes_before_start());
    task.set_reminder_minutes_before_start(999);
    auto item_id = service().create_item(task);
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
    task = service().get_task(item_id, ews::base_shape::all_properties);
    EXPECT_EQ(999U, task.get_reminder_minutes_before_start());
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
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
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
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
    task = service().get_task(item_id);
    EXPECT_EQ("", task.get_display_to());
}

TEST(OfflineExtendedFieldUriTest, DistPropertySetIdNameRoundTrip)
{
    // 1. based on distinguished_property_set_id and property_name
    const char* xml = "<t:ExtendedFieldURI "
                      "DistinguishedPropertySetId=\"PublicStrings\" "
                      "PropertyName=\"ShoeSize\" "
                      "PropertyType=\"Float\"/>";
    std::vector<char> buf(xml, xml + strlen(xml));
    buf.push_back('\0');

    xml_document doc;
    doc.parse<rapidxml::parse_no_namespace>(&buf[0]);
    auto node = doc.first_node();
    auto obj = ews::extended_field_uri::from_xml_element(*node);
    EXPECT_STREQ("PublicStrings",
                 obj.get_distinguished_property_set_id().c_str());
    EXPECT_STREQ("", obj.get_property_set_id().c_str());
    EXPECT_STREQ("", obj.get_property_tag().c_str());
    EXPECT_STREQ("ShoeSize", obj.get_property_name().c_str());
    EXPECT_STREQ("", obj.get_property_id().c_str());
    EXPECT_STREQ("Float", obj.get_property_type().c_str());
}

TEST(OfflineExtendedFieldUriTest, DistPropertySetIdIdRoundTrip)
{
    // 2. based on distinguished_property_set_id and property_id
    const char* xml = "<t:ExtendedFieldURI "
                      "DistinguishedPropertySetId=\"PublicStrings\" "
                      "PropertyId=\"42\" "
                      "PropertyType=\"Boolean\"/>";
    std::vector<char> buf(xml, xml + strlen(xml));
    buf.push_back('\0');

    xml_document doc;
    doc.parse<rapidxml::parse_no_namespace>(&buf[0]);
    auto node = doc.first_node();
    auto obj = ews::extended_field_uri::from_xml_element(*node);
    EXPECT_STREQ("PublicStrings",
                 obj.get_distinguished_property_set_id().c_str());
    EXPECT_STREQ("", obj.get_property_set_id().c_str());
    EXPECT_STREQ("", obj.get_property_tag().c_str());
    EXPECT_STREQ("", obj.get_property_name().c_str());
    EXPECT_STREQ("42", obj.get_property_id().c_str());
    EXPECT_STREQ("Boolean", obj.get_property_type().c_str());
}

TEST(OfflineExtendedFieldUriTest, PropertySetIdIdRoundTrip)
{
    // 3. based on property_set_id and property_id
    const char* xml = "<t:ExtendedFieldURI "
                      "PropertySetId=\"24040483-cda4-4521-bb5f-a83fac4d19a4\" "
                      "PropertyId=\"2\" "
                      "PropertyType=\"IntegerArray\"/>";
    std::vector<char> buf(xml, xml + strlen(xml));
    buf.push_back('\0');

    xml_document doc;
    doc.parse<rapidxml::parse_no_namespace>(&buf[0]);
    auto node = doc.first_node();
    auto obj = ews::extended_field_uri::from_xml_element(*node);
    EXPECT_STREQ("", obj.get_distinguished_property_set_id().c_str());
    EXPECT_STREQ("24040483-cda4-4521-bb5f-a83fac4d19a4",
                 obj.get_property_set_id().c_str());
    EXPECT_STREQ("", obj.get_property_tag().c_str());
    EXPECT_STREQ("", obj.get_property_name().c_str());
    EXPECT_STREQ("2", obj.get_property_id().c_str());
    EXPECT_STREQ("IntegerArray", obj.get_property_type().c_str());
}

TEST(OfflineExtendedFieldUriTest, PropertySetIdNameRoundTrip)
{
    // 4. based on property_set_id and property_name
    const char* xml = "<t:ExtendedFieldURI "
                      "PropertySetId=\"24040483-cda4-4521-bb5f-a83fac4d19a4\" "
                      "PropertyName=\"Rumpelstiltskin\" "
                      "PropertyType=\"Integer\"/>";
    std::vector<char> buf(xml, xml + strlen(xml));
    buf.push_back('\0');

    xml_document doc;
    doc.parse<rapidxml::parse_no_namespace>(&buf[0]);
    auto node = doc.first_node();
    auto obj = ews::extended_field_uri::from_xml_element(*node);
    EXPECT_STREQ("", obj.get_distinguished_property_set_id().c_str());
    EXPECT_STREQ("24040483-cda4-4521-bb5f-a83fac4d19a4",
                 obj.get_property_set_id().c_str());
    EXPECT_STREQ("", obj.get_property_tag().c_str());
    EXPECT_STREQ("Rumpelstiltskin", obj.get_property_name().c_str());
    EXPECT_STREQ("", obj.get_property_id().c_str());
    EXPECT_STREQ("Integer", obj.get_property_type().c_str());
}

TEST(OfflineExtendedFieldUriTest, PropertyTagRoundTrip)
{
    // 5. based on property_tag
    const char* xml = "<t:ExtendedFieldURI "
                      "PropertyTag=\"0x0036\" "
                      "PropertyType=\"Binary\"/>";
    std::vector<char> buf(xml, xml + strlen(xml));
    buf.push_back('\0');

    xml_document doc;
    doc.parse<rapidxml::parse_no_namespace>(&buf[0]);
    auto node = doc.first_node();
    auto obj = ews::extended_field_uri::from_xml_element(*node);
    EXPECT_STREQ("", obj.get_distinguished_property_set_id().c_str());
    EXPECT_STREQ("", obj.get_property_set_id().c_str());
    EXPECT_STREQ("0x0036", obj.get_property_tag().c_str());
    EXPECT_STREQ("", obj.get_property_name().c_str());
    EXPECT_STREQ("", obj.get_property_id().c_str());
    EXPECT_STREQ("Binary", obj.get_property_type().c_str());
}

TEST(OfflineExtendedPropertyTest, ExtendedProperty)
{
    auto msg = ews::message();

    std::vector<std::string> values;
    values.push_back("a lonesome violine string");

    ews::extended_field_uri field_uri(
        ews::extended_field_uri::property_set_id(
            "24040483-cda4-4521-bb5f-a83fac4d19a4"),
        ews::extended_field_uri::property_id("2"),
        ews::extended_field_uri::property_type("String"));

    ews::extended_property prop(field_uri, values);
    msg.set_extended_property(prop); // properties are set

    auto ep_actual = msg.get_extended_properties();
    ASSERT_FALSE(ep_actual.empty());
    auto efu_actual = ep_actual[0].get_extended_field_uri();

    EXPECT_STREQ("", efu_actual.get_distinguished_property_set_id().c_str());
    EXPECT_STREQ("24040483-cda4-4521-bb5f-a83fac4d19a4",
                 efu_actual.get_property_set_id().c_str());
    EXPECT_STREQ("", efu_actual.get_property_tag().c_str());
    EXPECT_STREQ("", efu_actual.get_property_name().c_str());
    EXPECT_STREQ("2", efu_actual.get_property_id().c_str());
    EXPECT_STREQ("String", efu_actual.get_property_type().c_str());

    msg = ews::message(); // ???
    field_uri = ews::extended_field_uri(
        ews::extended_field_uri::property_tag("0x0036"),
        ews::extended_field_uri::property_type("Integer"));

    prop = ews::extended_property(field_uri, values);
    msg.set_extended_property(prop);
    ep_actual = msg.get_extended_properties();
    ASSERT_FALSE(ep_actual.empty());
    efu_actual = ep_actual[0].get_extended_field_uri();

    EXPECT_STREQ("", efu_actual.get_distinguished_property_set_id().c_str());
    EXPECT_STREQ("", efu_actual.get_property_set_id().c_str());
    EXPECT_STREQ("0x0036", efu_actual.get_property_tag().c_str());
    EXPECT_STREQ("", efu_actual.get_property_name().c_str());
    EXPECT_STREQ("", efu_actual.get_property_id().c_str());
    EXPECT_STREQ("Integer", efu_actual.get_property_type().c_str());
}

TEST_F(ItemTest, ExtendedProperty)
{
    auto msg = ews::message(); // see book "Exchange Server 2007 - EWS" p.538

    // Set some constructors, send and get the properties back from Server
    // 1. based on property_set_id and property_id
    std::vector<ews::extended_field_uri> all_field_uri;
    ews::extended_field_uri field_uri1(
        ews::extended_field_uri::property_set_id(
            "24040483-cda4-4521-bb5f-a83fac4d19a4"),
        ews::extended_field_uri::property_id("2"),
        ews::extended_field_uri::property_type("StringArray"));
    std::vector<std::string> values;
    values.push_back("first string");
    values.push_back("second string");
    values.push_back("third string");
    ews::extended_property prop(field_uri1, values);
    msg.set_extended_property(prop);
    // 2. based on property_tag
    values.clear();
    values.push_back("12345");
    ews::extended_field_uri field_uri2(
        ews::extended_field_uri::property_tag("0x0036"),
        ews::extended_field_uri::property_type("Integer"));
    prop = ews::extended_property(field_uri2, values);
    msg.set_extended_property(prop);
    // 3. based on distinguished_property_set_id and property_name
    values.clear();
    values.push_back("12");
    ews::extended_field_uri field_uri3(
        ews::extended_field_uri::distinguished_property_set_id("PublicStrings"),
        ews::extended_field_uri::property_name("ShoeSize"),
        ews::extended_field_uri::property_type("Float"));
    prop = ews::extended_property(field_uri3, values);
    msg.set_extended_property(prop);

    auto item_id =
        service().create_item(msg, // message with all properties
                              ews::message_disposition::save_only); // created

    ews::internal::on_scope_exit remove_msg(
        [&] // make sure to remove msg
        { service().delete_message(std::move(msg)); });
    // set all field_uris we want to receive
    all_field_uri.push_back(field_uri1);
    all_field_uri.push_back(field_uri2);
    all_field_uri.push_back(field_uri3);
    ews::item_shape shape(std::move(all_field_uri));
    msg = service().get_message(item_id, shape);

    auto ep_actual = msg.get_extended_properties();
    ASSERT_TRUE(ep_actual.size() == 3);

    auto efu_actual = ep_actual[0].get_extended_field_uri();
    EXPECT_STREQ("first string", ep_actual[0].get_values()[0].c_str());
    EXPECT_STREQ("second string", ep_actual[0].get_values()[1].c_str());
    EXPECT_STREQ("third string", ep_actual[0].get_values()[2].c_str());
    EXPECT_STREQ("", efu_actual.get_distinguished_property_set_id().c_str());
    EXPECT_STREQ("24040483-cda4-4521-bb5f-a83fac4d19a4",
                 efu_actual.get_property_set_id().c_str());
    EXPECT_STREQ("", efu_actual.get_property_tag().c_str());
    EXPECT_STREQ("", efu_actual.get_property_name().c_str());
    EXPECT_STREQ("2", efu_actual.get_property_id().c_str());
    EXPECT_STREQ("StringArray", efu_actual.get_property_type().c_str());

    efu_actual = ep_actual[1].get_extended_field_uri();
    EXPECT_STREQ("12345", ep_actual[1].get_values()[0].c_str());
    EXPECT_STREQ("", efu_actual.get_distinguished_property_set_id().c_str());
    EXPECT_STREQ("", efu_actual.get_property_set_id().c_str());
    EXPECT_STREQ("0x36", efu_actual.get_property_tag().c_str()); // Exchange
                                                                 // removes
                                                                 // leading
                                                                 // zeroes
    EXPECT_STREQ("", efu_actual.get_property_name().c_str());
    EXPECT_STREQ("", efu_actual.get_property_id().c_str());
    EXPECT_STREQ("Integer", efu_actual.get_property_type().c_str());

    efu_actual = ep_actual[2].get_extended_field_uri();
    EXPECT_STREQ("12", ep_actual[2].get_values()[0].c_str());
    EXPECT_STREQ("PublicStrings",
                 efu_actual.get_distinguished_property_set_id().c_str());
    EXPECT_STREQ("", efu_actual.get_property_set_id().c_str());
    EXPECT_STREQ("", efu_actual.get_property_tag().c_str());
    EXPECT_STREQ("ShoeSize", efu_actual.get_property_name().c_str());
    EXPECT_STREQ("", efu_actual.get_property_id().c_str());
    EXPECT_STREQ("Float", efu_actual.get_property_type().c_str());
}

TEST(OfflineItemTest, CulturePropertyDefaultConstructed)
{
    auto task = ews::task();
    ASSERT_EQ("", task.get_culture());
    task.set_culture("zu-ZA");
    EXPECT_EQ("zu-ZA", task.get_culture());
}

TEST(OfflineItemTest, CultureProperty)
{
    auto task = make_fake_task();
    ASSERT_EQ("en-US", task.get_culture());
    task.set_culture("zu-ZA");
    EXPECT_EQ("zu-ZA", task.get_culture());
}

TEST_F(ItemTest, CultureProperty)
{
    auto task = ews::task();
    ASSERT_EQ("", task.get_culture());
    task.set_culture("mn-Mong-CN");
    ASSERT_EQ("mn-Mong-CN", task.get_culture());

    auto item_id = service().create_item(task);
    ews::internal::on_scope_exit remove_task(
        [&] { service().delete_task(std::move(task)); });
    task = service().get_task(item_id);

    auto prop = ews::property(ews::item_property_path::culture, "zu-ZA");
    item_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(item_id, ews::base_shape::all_properties);
    ASSERT_EQ("zu-ZA", task.get_culture());

    prop = ews::property(ews::item_property_path::culture, "yo-NG");
    item_id = service().update_item(task.get_item_id(), prop);
    task = service().get_task(item_id, ews::base_shape::all_properties);
    EXPECT_EQ("yo-NG", task.get_culture());
}

TEST_F(ItemTest, CreateItemsWithVector)
{
    std::vector<ews::contact> contacts;
    for (int i = 0; i < 10; i++)
    {
        contacts.push_back(ews::contact());
    }

    auto contact_ids = service().create_item(contacts);
    ews::internal::on_scope_exit remove_contacts([&] {
        for (auto& id : contact_ids)
        {
            service().delete_item(id);
        }
    });
    EXPECT_EQ(10u, contact_ids.size());
}

TEST_F(ItemTest, CreateItemsWithEmptyVector)
{
    std::vector<ews::contact> contacts;

    try
    {
        auto contact_ids = service().create_item(contacts);
    }
    catch (ews::exception&)
    {
        return;
    }
    FAIL();
}

TEST_F(ItemTest, SendItemsWithVector)
{
    std::vector<ews::message> messages;
    auto message = ews::message();
    std::vector<ews::mailbox> recipients;
    for (int i = 0; i < 3; i++)
    {
        message.set_subject("?");
        recipients.push_back(ews::mailbox("whoishere@home.com"));
        message.set_to_recipients(recipients);
        messages.push_back(message);
    }

    auto message_ids =
        service().create_item(messages, ews::message_disposition::save_only);
    ews::internal::on_scope_exit remove_messages([&] {
        for (auto& id : message_ids)
        {
            try
            {
                service().delete_item(id);
            }
            catch (std::exception&)
            {
                // Swallow
            }
        }
    });

    EXPECT_NO_THROW(service().send_item(message_ids));
}

TEST_F(ItemTest, SendItemsWithVectorAndFolder)
{
    std::vector<ews::message> messages;
    auto message = ews::message();
    std::vector<ews::mailbox> recipients;
    const ews::distinguished_folder_id drafts = ews::standard_folder::drafts;
    const auto initial_count = service().find_item(drafts).size();
    for (int i = 0; i < 3; i++)
    {
        message.set_subject("?");
        recipients.push_back(ews::mailbox("whoishere@home.com"));
        message.set_to_recipients(recipients);
        messages.push_back(message);
    }

    auto message_ids =
        service().create_item(messages, ews::message_disposition::save_only);
    ews::internal::on_scope_exit remove_messages([&] {
        for (auto& id : message_ids)
        {
            try
            {
                service().delete_item(id);
            }
            catch (std::exception&)
            {
                // Swallow
            }
        }
    });

    EXPECT_NO_THROW(service().send_item(message_ids, drafts));
    EXPECT_EQ(initial_count + 3, service().find_item(drafts).size());
}

TEST_F(ItemTest, SendItemsWithEmptyVector)
{
    std::vector<ews::item_id> ids;
    EXPECT_THROW(service().send_item(ids), ews::exception);
}
} // namespace tests
