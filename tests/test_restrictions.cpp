#include "fixture.hpp"

using namespace ews;

namespace tests
{
    TEST(RestrictionTest, IsEqualToBooleanConstantRendersCorrectly)
    {
        const char* expected =
R"(<IsEqualTo><FieldURI FieldURI="task:IsComplete"/><FieldURIOrConstant><Constant Value="true"/></FieldURIOrConstant></IsEqualTo>)";
        task_property_path task;
        auto restr = is_equal_to(task.is_complete, true);
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, IsEqualToBooleanConstantRendersCorrectlyWithNamespace)
    {
        const char* expected =
R"(<s:IsEqualTo><s:FieldURI FieldURI="task:IsComplete"/><s:FieldURIOrConstant><s:Constant Value="false"/></s:FieldURIOrConstant></s:IsEqualTo>)";
        task_property_path task;
        auto restr = is_equal_to(task.is_complete, false);
        EXPECT_STREQ(expected, restr.to_xml("s").c_str());
    }

    TEST(RestrictionTest, IsEqualToStringConstantRendersCorrectly)
    {
        // TODO: check with valgrind memtool; see capture clause in c'tor
        // overload
        const char* expected =
R"(<IsEqualTo><FieldURI FieldURI="folder:DisplayName"/><FieldURIOrConstant><Constant Value="Inbox"/></FieldURIOrConstant></IsEqualTo>)";
        folder_property_path folder;
        auto restr = is_equal_to(folder.display_name, "Inbox");
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, IsEqualToDateTimeConstantRendersCorrectly)
    {
        const char* expected =
R"(<IsEqualTo><FieldURI FieldURI="item:DateTimeCreated"/><FieldURIOrConstant><Constant Value="2015-05-28T17:39:11Z"/></FieldURIOrConstant></IsEqualTo>)";
        const auto yesterday = date_time("2015-05-28T17:39:11Z");
        item_property_path item;
        auto restr = is_equal_to(item.date_time_created, yesterday);
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }
}
