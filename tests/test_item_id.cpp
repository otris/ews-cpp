#include <ews/ews.hpp>
#include <gtest/gtest.h>

TEST(ItemIdTest, Construction)
{
    using ews::item_id;

    const std::string id("AAMkAGM3OTdhOGY0LThhYzgtNDRlMy05NTRiLTE5MDFmODVmNjVmOAAuAAAAAAC4y7EaeL6ESIX5BrvsazIoAQDf1GRGe+uRQ4zRVnseDbqzAAAAAAEBAAA=");
    const std::string change_key("AQAAABYAAADf1GRGe+uRQ4zRVnseDbqzAAACf02r");

    auto a = item_id(id);
    ASSERT_STREQ(a.id().c_str(), id.c_str());
    ASSERT_STREQ(a.change_key().c_str(), "");

    auto b = item_id(id, change_key);
    ASSERT_STREQ(b.id().c_str(), id.c_str());
    ASSERT_STREQ(b.change_key().c_str(), change_key.c_str());
}
