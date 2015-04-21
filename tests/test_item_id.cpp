#include <ews/ews.hpp>
#include <gtest/gtest.h>

using ews::item_id;
using xml_document = rapidxml::xml_document<>;

namespace tests
{
    TEST(ItemIdTest, ConstructWithIdOnly)
    {
        auto a = item_id("abcde");
        EXPECT_STREQ(a.id().c_str(), "abcde");
        EXPECT_STREQ(a.change_key().c_str(), "");
    }

    TEST(ItemIdTest, ConstructWithIdAndChangeKey)
    {
        auto a = item_id("abcde", "edcba");
        EXPECT_STREQ(a.id().c_str(), "abcde");
        EXPECT_STREQ(a.change_key().c_str(), "edcba");
    }

    TEST(ItemIdTest, DefaultConstruction)
    {
        auto a = item_id();
        EXPECT_FALSE(a.valid());
        EXPECT_STREQ(a.id().c_str(), "");
        EXPECT_STREQ(a.change_key().c_str(), "");
    }

#pragma warning(suppress: 6262)
    TEST(ItemIdTest, FromXmlNode)
    {
        char buf[] = "<ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>";
        xml_document doc;
        doc.parse<0>(buf);
        auto node = doc.first_node();
        auto a = item_id::from_xml_element(*node);
        EXPECT_STREQ(a.id().c_str(), "abcde");
        EXPECT_STREQ(a.change_key().c_str(), "edcba");
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
#define TESTS_ITEM_ID_XML "<ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>"

        char buf[] = TESTS_ITEM_ID_XML;
        xml_document doc;
        doc.parse<0>(buf);
        auto node = doc.first_node();
        auto a = item_id::from_xml_element(*node);
        EXPECT_STREQ(TESTS_ITEM_ID_XML, a.to_xml().c_str());
    }
}
