#include "fixtures.hpp"

namespace tests
{
    TEST(RestrictionTest, ContainsRendersCorrectly)
    {
        // Mode is `Substring` and comparison `Loose` by default

        const char* expected = "<Contains "
                               "ContainmentMode=\"Substring\" "
                               "ContainmentComparison=\"Loose\">"
                               "<FieldURI FieldURI=\"item:Subject\"/>"
                               "<Constant Value=\"Vogon poetry\"/>"
                               "</Contains>";

        auto restr =
            ews::contains(ews::item_property_path::subject, "Vogon poetry");
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, AndExpressionRendersCorrectly)
    {
        using ews::and_;
        using ews::contains;
        using ews::is_equal_to;

        const char* expected = "<And>"
                               "<IsEqualTo>"
                               "<FieldURI FieldURI=\"task:IsComplete\"/>"
                               "<FieldURIOrConstant>"
                               "<Constant Value=\"true\"/>"
                               "</FieldURIOrConstant>"
                               "</IsEqualTo>"
                               "<Contains "
                               "ContainmentMode=\"Substring\" "
                               "ContainmentComparison=\"Loose\">"
                               "<FieldURI FieldURI=\"item:Subject\"/>"
                               "<Constant Value=\"Baseball\"/>"
                               "</Contains>"
                               "</And>";

        auto restr =
            and_(is_equal_to(ews::task_property_path::is_complete, true),
                 contains(ews::item_property_path::subject, "Baseball"));

        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, OrExpressionRendersCorrectly)
    {
        using ews::or_;
        using ews::contains;
        using ews::is_equal_to;

        const char* expected = "<Or>"
                               "<IsEqualTo>"
                               "<FieldURI FieldURI=\"task:IsComplete\"/>"
                               "<FieldURIOrConstant>"
                               "<Constant Value=\"false\"/>"
                               "</FieldURIOrConstant>"
                               "</IsEqualTo>"
                               "<Contains "
                               "ContainmentMode=\"Substring\" "
                               "ContainmentComparison=\"Loose\">"
                               "<FieldURI FieldURI=\"item:Subject\"/>"
                               "<Constant Value=\"Soccer\"/>"
                               "</Contains>"
                               "</Or>";

        auto restr =
            or_(is_equal_to(ews::task_property_path::is_complete, false),
                contains(ews::item_property_path::subject, "Soccer"));

        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, NotExpressionRendersCorrectly)
    {
        using ews::not_;
        using ews::is_equal_to;

        const char* expected = "<Not>"
                               "<IsEqualTo>"
                               "<FieldURI FieldURI=\"task:IsComplete\"/>"
                               "<FieldURIOrConstant>"
                               "<Constant Value=\"true\"/>"
                               "</FieldURIOrConstant>"
                               "</IsEqualTo>"
                               "</Not>";

        auto restr =
            not_(is_equal_to(ews::task_property_path::is_complete, true));
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, IsEqualToBooleanConstantRendersCorrectly)
    {
        const char* expected = "<IsEqualTo>"
                               "<FieldURI FieldURI=\"task:IsComplete\"/>"
                               "<FieldURIOrConstant>"
                               "<Constant Value=\"true\"/>"
                               "</FieldURIOrConstant>"
                               "</IsEqualTo>";

        auto restr =
            ews::is_equal_to(ews::task_property_path::is_complete, true);
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, IsEqualToBooleanConstantRendersCorrectlyWithNamespace)
    {
        const char* expected = "<s:IsEqualTo>"
                               "<s:FieldURI FieldURI=\"task:IsComplete\"/>"
                               "<s:FieldURIOrConstant>"
                               "<s:Constant Value=\"false\"/>"
                               "</s:FieldURIOrConstant>"
                               "</s:IsEqualTo>";

        auto restr =
            ews::is_equal_to(ews::task_property_path::is_complete, false);
        EXPECT_STREQ(expected, restr.to_xml("s").c_str());
    }

    TEST(RestrictionTest, IsEqualToStringConstantRendersCorrectly)
    {
        const char* expected = "<IsEqualTo>"
                               "<FieldURI FieldURI=\"folder:DisplayName\"/>"
                               "<FieldURIOrConstant>"
                               "<Constant Value=\"Inbox\"/>"
                               "</FieldURIOrConstant>"
                               "</IsEqualTo>";

        auto restr =
            ews::is_equal_to(ews::folder_property_path::display_name, "Inbox");
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, IsEqualToDateTimeConstantRendersCorrectly)
    {
        const char* expected = "<IsEqualTo>"
                               "<FieldURI FieldURI=\"item:DateTimeCreated\"/>"
                               "<FieldURIOrConstant>"
                               "<Constant Value=\"2015-05-28T17:39:11Z\"/>"
                               "</FieldURIOrConstant>"
                               "</IsEqualTo>";

        const auto yesterday = ews::date_time("2015-05-28T17:39:11Z");
        auto restr = ews::is_equal_to(
            ews::item_property_path::date_time_created, yesterday);
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest,
         IndexedFieldURIIsEqualToStringConstantRendersCorrectly)
    {
        const char* expected = "<IsEqualTo>"
                               "<IndexedFieldURI "
                               "FieldURI=\"contacts:EmailAddress\" "
                               "FieldIndex=\"EmailAddress3\"/>"
                               "<FieldURIOrConstant>"
                               "<Constant Value=\"jane.dow@contoso.com\"/>"
                               "</FieldURIOrConstant>"
                               "</IsEqualTo>";

        auto restr =
            ews::is_equal_to(ews::contact_property_path::email_address_3,
                             "jane.dow@contoso.com");
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest,
         IndexedFieldURIIsEqualToStringConstantRendersCorrectlyWithNamespace)
    {
        const char* expected = "<t:IsEqualTo>"
                               "<t:IndexedFieldURI "
                               "FieldURI=\"contacts:EmailAddress\" "
                               "FieldIndex=\"EmailAddress1\"/>"
                               "<t:FieldURIOrConstant>"
                               "<t:Constant Value=\"bruce@willis.com\"/>"
                               "</t:FieldURIOrConstant>"
                               "</t:IsEqualTo>";

        auto restr = ews::is_equal_to(
            ews::contact_property_path::email_address_1, "bruce@willis.com");
        EXPECT_STREQ(expected, restr.to_xml("t").c_str());
    }

    TEST(RestrictionTest, IsNotEqualToBooleanConstantRendersCorrectly)
    {
        const char* expected = "<t:IsNotEqualTo>"
                               "<t:FieldURI FieldURI=\"task:IsComplete\"/>"
                               "<t:FieldURIOrConstant>"
                               "<t:Constant Value=\"true\"/>"
                               "</t:FieldURIOrConstant>"
                               "</t:IsNotEqualTo>";

        auto restr =
            ews::is_not_equal_to(ews::task_property_path::is_complete, true);
        EXPECT_STREQ(expected, restr.to_xml("t").c_str());
    }

    TEST(RestrictionTest, IsGreaterThanIntConstantRendersCorrectly)
    {
        const char* expected = "<IsGreaterThan>"
                               "<FieldURI FieldURI=\"task:PercentComplete\"/>"
                               "<FieldURIOrConstant>"
                               "<Constant Value=\"75\"/>"
                               "</FieldURIOrConstant>"
                               "</IsGreaterThan>";

        auto restr =
            ews::is_greater_than(ews::task_property_path::percent_complete, 75);
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }

    TEST(RestrictionTest, IsGreaterThanOrEqualToIntConstantRendersCorrectly)
    {
        const char* expected = "<IsGreaterThanOrEqualTo>"
                               "<FieldURI FieldURI=\"task:PercentComplete\"/>"
                               "<FieldURIOrConstant>"
                               "<Constant Value=\"80\"/>"
                               "</FieldURIOrConstant>"
                               "</IsGreaterThanOrEqualTo>";

        auto restr = ews::is_greater_than_or_equal_to(
            ews::task_property_path::percent_complete, 80);
        EXPECT_STREQ(expected, restr.to_xml().c_str());
    }
}

// vim:et ts=4 sw=4 noic cc=80
