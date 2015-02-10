#include <ews/ews.hpp>
#include <gtest/gtest.h>

using ews::item_id;
using xml_document = rapidxml::xml_document<>;

namespace tests
{
    TEST(ItemIdTest, ConstructWithIdOnly)
    {
        auto a = item_id("abcde");
        ASSERT_STREQ(a.id().c_str(), "abcde");
        ASSERT_STREQ(a.change_key().c_str(), "");
    }

    TEST(ItemIdTest, ConstructWithIdAndChangeKey)
    {
        auto a = item_id("abcde", "edcba");
        ASSERT_STREQ(a.id().c_str(), "abcde");
        ASSERT_STREQ(a.change_key().c_str(), "edcba");
    }

    TEST(ItemIdTest, FromXmlNode)
    {
        char buf[] = "<ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>";
        xml_document doc;
        doc.parse<0>(buf);
        auto node = doc.first_node();
        auto a = item_id::from_xml_element(*node);
        ASSERT_STREQ(a.id().c_str(), "abcde");
        ASSERT_STREQ(a.change_key().c_str(), "edcba");
    }

    TEST(ItemIdTest, ToXMLWithNamespace)
    {
        const char* expected = "<t:ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>";
        const auto a = item_id("abcde", "edcba");
        ASSERT_STREQ(expected, a.to_xml("t").c_str());
    }

    TEST(ItemIdTest, FromAndToXMLRoundTrip)
    {
#define TESTS_ITEM_ID_XML "<ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>"

        char buf[] = TESTS_ITEM_ID_XML;
        xml_document doc;
        doc.parse<0>(buf);
        auto node = doc.first_node();
        auto a = item_id::from_xml_element(*node);
        ASSERT_STREQ(TESTS_ITEM_ID_XML, a.to_xml().c_str());
    }
}
