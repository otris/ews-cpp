#include <ews/ews.hpp>
#include <gtest/gtest.h>
#include <vector>
#include <cstring>

using ews::item_id;
using xml_document = rapidxml::xml_document<>;

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
}
