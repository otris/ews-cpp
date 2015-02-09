#include <ews/ews.hpp>
#include <gtest/gtest.h>

using ews::item_id;

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
        using rapidxml::xml_document;
        char buf[] = "<ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>";
        xml_document<> doc;
        doc.parse<0>(buf);
        auto node = doc.first_node();
        auto a = item_id::from_xml_element(*node);
        ASSERT_STREQ(a.id().c_str(), "abcde");
        ASSERT_STREQ(a.change_key().c_str(), "edcba");
    }
}
