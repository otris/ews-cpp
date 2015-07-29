#include "fixtures.hpp"

#include <vector>
#include <cstring>

using xml_document = rapidxml::xml_document<>;

namespace tests
{
    TEST(AttachmentIdTest, DefaultConstruction)
    {
        auto obj = ews::attachment_id();
        EXPECT_FALSE(obj.valid());
        EXPECT_STREQ("", obj.id().c_str());
        EXPECT_FALSE(obj.root_item_id().valid());
    }

    TEST(AttachmentIdTest, ConstructFromId)
    {
        // As in GetAttachment and DeleteAttachment operations
        auto obj = ews::attachment_id("abcde");
        EXPECT_TRUE(obj.valid());
        EXPECT_STREQ("abcde", obj.id().c_str());
        EXPECT_FALSE(obj.root_item_id().valid());
    }

    TEST(AttachmentIdTest, ConstructFromIdAndRootItemId)
    {
        // As in CreateAttachment operation
        auto obj = ews::attachment_id("abcde",
                                      ews::item_id("edcba", "qwertz"));
        EXPECT_TRUE(obj.valid());
        EXPECT_STREQ("abcde", obj.id().c_str());
        EXPECT_TRUE(obj.root_item_id().valid());
        EXPECT_STREQ("edcba", obj.root_item_id().id().c_str());
        EXPECT_STREQ("qwertz", obj.root_item_id().change_key().c_str());
    }

    TEST(AttachmentIdTest, FromXmlNodeWithIdAttributeOnly)
    {
        const char* xml = R"(<AttachmentId Id="abcde"/>)";
        std::vector<char> buf(xml, xml + std::strlen(xml));
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = doc.first_node();
        auto obj = ews::attachment_id::from_xml_element(*node);
        EXPECT_TRUE(obj.valid());
        EXPECT_STREQ("abcde", obj.id().c_str());
        EXPECT_FALSE(obj.root_item_id().valid());
    }

    TEST(AttachmentIdTest, FromXmlNodeWithIdAndRootIdAttributes)
    {
        const char* xml =
R"(<AttachmentId Id="abcde" RootItemId="qwertz" RootItemChangeKey="edcba"/>)";
        std::vector<char> buf(xml, xml + std::strlen(xml));
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = doc.first_node();
        auto obj = ews::attachment_id::from_xml_element(*node);
        EXPECT_TRUE(obj.valid());
        EXPECT_STREQ("abcde", obj.id().c_str());
        EXPECT_TRUE(obj.root_item_id().valid());
        EXPECT_STREQ("qwertz", obj.root_item_id().id().c_str());
        EXPECT_STREQ("edcba", obj.root_item_id().change_key().c_str());
    }

    TEST(AttachmentIdTest, ToXMLWithNamespace)
    {
        const char* expected =
R"(<ns:AttachmentId Id="abcde" RootItemId="qwertz" RootItemChangeKey="edcba"/>)";
        const auto obj =
            ews::attachment_id("abcde", ews::item_id("qwertz", "edcba"));
        EXPECT_STREQ(expected, obj.to_xml("ns").c_str());
    }

    TEST(AttachmentIdTest, ToXML)
    {
        const char* expected =
R"(<AttachmentId Id="abcde" RootItemId="qwertz" RootItemChangeKey="edcba"/>)";
        const auto obj =
            ews::attachment_id("abcde", ews::item_id("qwertz", "edcba"));
        EXPECT_STREQ(expected, obj.to_xml().c_str());
    }

    TEST(AttachmentIdTest, FromAndToXMLRoundTrip)
    {
        const char* xml =
R"(<AttachmentId Id="abcde" RootItemId="qwertz" RootItemChangeKey="edcba"/>)";
        std::vector<char> buf(xml, xml + std::strlen(xml));
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = doc.first_node();
        auto obj = ews::attachment_id::from_xml_element(*node);
        EXPECT_STREQ(xml, obj.to_xml().c_str());
    }

    TEST_F(ItemAttachmentTest, DefaultConstructorCreatesItemAttachment)
    {
        auto attachment = ews::attachment();
        EXPECT_EQ(ews::attachment::type::item, attachment.get_type());
    }

    // TEST_F(ItemAttachmentTest, CreateFromItem)
    // {
    //     auto item_attachment = ews::attachment::from_item();
    // }

#ifdef EWS_USE_BOOST_LIBRARY
    TEST_F(FileAttachmentTest, CreateFromFile)
    {
        const auto path = assets_dir()/"ballmer_peak.png";
        auto file_attachment = ews::attachment::from_file(path.string(),
                                                          "image/png",
                                                          "Ballmer Peak");
        EXPECT_EQ(ews::attachment::type::file, file_attachment.get_type());
        EXPECT_FALSE(file_attachment.id().valid());
        EXPECT_STREQ("Ballmer Peak", file_attachment.name().c_str());
        EXPECT_STREQ("image/png", file_attachment.content_type().c_str());
        EXPECT_NE(nullptr, file_attachment.content());
        EXPECT_EQ(93525U, file_attachment.content_size());
    }

    TEST_F(FileAttachmentTest, CreateFromFileThrowsIfFileDoesNotExists)
    {
        const auto path = assets_dir()/"unlikely_to_exist.txt";
        EXPECT_THROW(
        {
            ews::attachment::from_file(path.string(), "image/png", "");

        }, ews::exception);
    }

    TEST_F(FileAttachmentTest, CreateFromFileExceptionMessage)
    {
        const auto path = assets_dir()/"unlikely_to_exist.txt";
        try
        {
            ews::attachment::from_file(path.string(), "image/png", "");
            FAIL() << "Expected exception to be raised";
        }
        catch (ews::exception& exc)
        {
            auto expected = "Could not open file for reading: " + path.string();
            EXPECT_STREQ(expected.c_str(), exc.what());
        }
    }
#endif // EWS_USE_BOOST_LIBRARY

    // TEST_F(AttachmentTest, CreateAndDeleteAttachment)
    // {
    //     auto& msg = test_message();
    //     auto file_attachment = ews::attachment::from_file();
    //     auto attachment_id = service().create_attachment(msg, file_attachment);
    //     ASSERT_TRUE(attachment_id.valid());
    // }
}
