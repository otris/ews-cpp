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
        EXPECT_STREQ(expected, a.to_xml("t").c_str());
    }

#pragma warning(suppress: 6262)
    TEST(ItemIdTest, FromAndToXMLRoundTrip)
    {
        const char* xml = "<ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>";
        std::vector<char> buf(xml, xml + std::strlen(xml));
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
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
            "<Body BodyType=\"Text\">Here is some plain text</Body>";
        EXPECT_STREQ(expected, b.to_xml().c_str());
    }

    TEST(BodyTest, MakeSureHTMLIsWrappedWithCDATA)
    {
        auto b = ews::body("<b>Here is some HTML</b>", ews::body_type::html);
        EXPECT_EQ(ews::body_type::html, b.type());
        EXPECT_FALSE(b.is_truncated());
        const char* const expected =
"<x:Body BodyType=\"HTML\"><![CDATA[<b>Here is some HTML</b>]]></x:Body>";
        EXPECT_STREQ(expected, b.to_xml("x").c_str());
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
        EXPECT_STREQ("Calendar", path.class_name().c_str());

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

    TEST(ItemTest, DefaultConstruction)
    {
        auto i = ews::item();
        EXPECT_TRUE(i.get_mime_content().none());
        EXPECT_STREQ("", i.get_subject().c_str());
        EXPECT_FALSE(i.get_item_id().valid());
    }

    TEST(ItemTest, GetAndSetBodyProperty)
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

    TEST(ItemTest, GetParentFolderIdProperty)
    {
        const auto task = make_fake_task();
        auto parent_folder_id = task.get_parent_folder_id();
        EXPECT_TRUE(parent_folder_id.valid());
        EXPECT_STREQ("qwertz", parent_folder_id.id().c_str());
        EXPECT_STREQ("ztrewq", parent_folder_id.change_key().c_str());
    }

    TEST(ItemTest, GetItemClassProperty)
    {
        const auto task = make_fake_task();
        const auto item_class = task.get_item_class();
        EXPECT_STREQ("IPM.Task", item_class.c_str());
    }

    TEST(ItemTest, GetSensitivityProperty)
    {
        const auto task = make_fake_task();
        EXPECT_EQ(ews::sensitivity::confidential, task.get_sensitivity());
    }

    TEST(ItemTest, GetSensitivityPropertyDefaultConstructed)
    {
        const auto task = ews::task();
        EXPECT_EQ(ews::sensitivity::normal, task.get_sensitivity());
    }

    TEST(ItemTest, SetSensitivity)
    {
        auto task = ews::task();
        task.set_sensitivity(ews::sensitivity::personal);
        EXPECT_EQ(ews::sensitivity::personal, task.get_sensitivity());
    }

    class ItemTestAgainstExchangeServer : public ServiceFixture {};

    TEST_F(ItemTestAgainstExchangeServer, GetMimeContentProperty)
    {
        // A contact is an item
        auto contact = ews::contact();
        auto s = service();
        const auto item_id = s.create_item(contact);
        ews::internal::on_scope_exit remove_contact([&]
        {
            s.delete_contact(std::move(contact));
        });
        contact = s.get_contact(item_id);
        EXPECT_FALSE(contact.get_mime_content().none());
    }

    TEST_F(ItemTestAgainstExchangeServer, UpdateSensitivityProperty)
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
}
