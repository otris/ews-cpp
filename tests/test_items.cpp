#include "fixtures.hpp"

#include <ews/ews.hpp>

namespace tests
{
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
        const auto task = make_task(
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

        auto parent_folder_id = task.get_parent_folder_id();
        EXPECT_TRUE(parent_folder_id.valid());
        EXPECT_STREQ("qwertz", parent_folder_id.id().c_str());
        EXPECT_STREQ("ztrewq", parent_folder_id.change_key().c_str());
    }

    class ItemTestAgainstExchangeServer : public ServiceFixture {};

    TEST_F(ItemTestAgainstExchangeServer, GetMimeContentProperty)
    {
        // A contact is an item
        auto contact = ews::contact();
        auto s = service();
        const auto item_id = s.create_item(contact);
        ews::internal::on_scope_exit delete_task([&]
        {
            s.delete_contact(std::move(contact));
        });
        contact = s.get_contact(item_id);
        EXPECT_FALSE(contact.get_mime_content().none());
    }
}
