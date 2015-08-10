#include "fixtures.hpp"

using namespace ews;

namespace tests
{
    TEST(RestrictionTest, IsEqualToBooleanConstantRendersCorrectly)
    {
        const char* expected =
"<IsEqualTo><FieldURI FieldURI=\"task:IsComplete\"/><FieldURIOrConstant><Constant Value=\"true\"/></FieldURIOrConstant></IsEqualTo>";
        auto restr = is_equal_to(task_property_path::is_complete, true);
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, IsEqualToBooleanConstantRendersCorrectlyWithNamespace)
    {
        const char* expected =
"<s:IsEqualTo><s:FieldURI FieldURI=\"task:IsComplete\"/><s:FieldURIOrConstant><s:Constant Value=\"false\"/></s:FieldURIOrConstant></s:IsEqualTo>";
        auto restr = is_equal_to(task_property_path::is_complete, false);
        EXPECT_STREQ(expected, restr.to_xml("s").c_str());
    }

    TEST(RestrictionTest, IsEqualToStringConstantRendersCorrectly)
    {
        const char* expected =
"<IsEqualTo><FieldURI FieldURI=\"folder:DisplayName\"/><FieldURIOrConstant><Constant Value=\"Inbox\"/></FieldURIOrConstant></IsEqualTo>";
        auto restr = is_equal_to(folder_property_path::display_name, "Inbox");
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, IsEqualToDateTimeConstantRendersCorrectly)
    {
        const char* expected =
"<IsEqualTo><FieldURI FieldURI=\"item:DateTimeCreated\"/><FieldURIOrConstant><Constant Value=\"2015-05-28T17:39:11Z\"/></FieldURIOrConstant></IsEqualTo>";
        const auto yesterday = date_time("2015-05-28T17:39:11Z");
        auto restr = is_equal_to(item_property_path::date_time_created,
                                 yesterday);
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, IndexedFieldURIIsEqualToStringConstantRendersCorrectly)
    {
        const char* expected =
"<IsEqualTo><IndexedFieldURI FieldURI=\"contacts:EmailAddress\" FieldIndex=\"EmailAddress3\"/><FieldURIOrConstant><Constant Value=\"jane.dow@contoso.com\"/></FieldURIOrConstant></IsEqualTo>";
        auto restr = is_equal_to(contact_property_path::email_address_3,
                                 "jane.dow@contoso.com");
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, IndexedFieldURIIsEqualToStringConstantRendersCorrectlyWithNamespace)
    {
        const char* expected =
"<t:IsEqualTo><t:IndexedFieldURI FieldURI=\"contacts:EmailAddress\" FieldIndex=\"EmailAddress1\"/><t:FieldURIOrConstant><t:Constant Value=\"bruce@willis.com\"/></t:FieldURIOrConstant></t:IsEqualTo>";
        auto restr = is_equal_to(contact_property_path::email_address_1,
                                 "bruce@willis.com");
        EXPECT_STREQ(expected, restr.to_xml("t").c_str());
    }
}
