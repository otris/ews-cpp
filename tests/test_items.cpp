#include <ews/ews.hpp>
#include <gtest/gtest.h>

using ews::item;

namespace tests
{
    TEST(ItemsTest, DefaultConstruction)
    {
        auto i = item();
        EXPECT_STREQ(i.get_subject().c_str(), "");
        EXPECT_FALSE(i.get_item_id().valid());
    }
}
