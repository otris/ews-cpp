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

    TEST_F(AttachmentTest, CreateAndDeleteAttachment)
    {
        FAIL() << "TODO";
    }
}
