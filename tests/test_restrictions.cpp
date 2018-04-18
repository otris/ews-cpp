
//   Copyright 2016 otris software AG
//
//   Licensed under the Apache License, Version 2.0 (the "License");
//   you may not use this file except in compliance with the License.
//   You may obtain a copy of the License at
//
//       http://www.apache.org/licenses/LICENSE-2.0
//
//   Unless required by applicable law or agreed to in writing, software
//   distributed under the License is distributed on an "AS IS" BASIS,
//   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//   See the License for the specific language governing permissions and
//   limitations under the License.
//
//   This project is hosted at https://github.com/otris

#include "fixtures.hpp"

namespace tests
{
TEST(RestrictionTest, ContainsRendersCorrectly)
{
    // Mode is `Substring` and comparison `Loose` by default

    const char* expected = "<t:Contains "
                           "ContainmentMode=\"Substring\" "
                           "ContainmentComparison=\"Loose\">"
                           "<t:FieldURI FieldURI=\"item:Subject\"/>"
                           "<t:Constant Value=\"Vogon poetry\"/>"
                           "</t:Contains>";

    auto restr =
        ews::contains(ews::item_property_path::subject, "Vogon poetry");
    EXPECT_STREQ(expected, restr.to_xml().c_str());
}

TEST(RestrictionTest, AndExpressionRendersCorrectly)
{
    using ews::and_;
    using ews::contains;
    using ews::is_equal_to;

    const char* expected = "<t:And>"
                           "<t:IsEqualTo>"
                           "<t:FieldURI FieldURI=\"task:IsComplete\"/>"
                           "<t:FieldURIOrConstant>"
                           "<t:Constant Value=\"true\"/>"
                           "</t:FieldURIOrConstant>"
                           "</t:IsEqualTo>"
                           "<t:Contains "
                           "ContainmentMode=\"Substring\" "
                           "ContainmentComparison=\"Loose\">"
                           "<t:FieldURI FieldURI=\"item:Subject\"/>"
                           "<t:Constant Value=\"Baseball\"/>"
                           "</t:Contains>"
                           "</t:And>";

    auto restr = and_(is_equal_to(ews::task_property_path::is_complete, true),
                      contains(ews::item_property_path::subject, "Baseball"));

    EXPECT_STREQ(expected, restr.to_xml().c_str());
}

TEST(RestrictionTest, OrExpressionRendersCorrectly)
{
    using ews::contains;
    using ews::is_equal_to;
    using ews::or_;

    const char* expected = "<t:Or>"
                           "<t:IsEqualTo>"
                           "<t:FieldURI FieldURI=\"task:IsComplete\"/>"
                           "<t:FieldURIOrConstant>"
                           "<t:Constant Value=\"false\"/>"
                           "</t:FieldURIOrConstant>"
                           "</t:IsEqualTo>"
                           "<t:Contains "
                           "ContainmentMode=\"Substring\" "
                           "ContainmentComparison=\"Loose\">"
                           "<t:FieldURI FieldURI=\"item:Subject\"/>"
                           "<t:Constant Value=\"Soccer\"/>"
                           "</t:Contains>"
                           "</t:Or>";

    auto restr = or_(is_equal_to(ews::task_property_path::is_complete, false),
                     contains(ews::item_property_path::subject, "Soccer"));

    EXPECT_STREQ(expected, restr.to_xml().c_str());
}

TEST(RestrictionTest, NotExpressionRendersCorrectly)
{
    using ews::is_equal_to;
    using ews::not_;

    const char* expected = "<t:Not>"
                           "<t:IsEqualTo>"
                           "<t:FieldURI FieldURI=\"task:IsComplete\"/>"
                           "<t:FieldURIOrConstant>"
                           "<t:Constant Value=\"true\"/>"
                           "</t:FieldURIOrConstant>"
                           "</t:IsEqualTo>"
                           "</t:Not>";

    auto restr = not_(is_equal_to(ews::task_property_path::is_complete, true));
    EXPECT_STREQ(expected, restr.to_xml().c_str());
}

TEST(RestrictionTest, IsEqualToBooleanConstantRendersCorrectly)
{
    const char* expected = "<t:IsEqualTo>"
                           "<t:FieldURI FieldURI=\"task:IsComplete\"/>"
                           "<t:FieldURIOrConstant>"
                           "<t:Constant Value=\"true\"/>"
                           "</t:FieldURIOrConstant>"
                           "</t:IsEqualTo>";

    auto restr = ews::is_equal_to(ews::task_property_path::is_complete, true);
    EXPECT_STREQ(expected, restr.to_xml().c_str());
}

TEST(RestrictionTest, IsEqualToBooleanConstantRendersCorrectlyWithNamespace)
{
    const char* expected = "<t:IsEqualTo>"
                           "<t:FieldURI FieldURI=\"task:IsComplete\"/>"
                           "<t:FieldURIOrConstant>"
                           "<t:Constant Value=\"false\"/>"
                           "</t:FieldURIOrConstant>"
                           "</t:IsEqualTo>";

    auto restr = ews::is_equal_to(ews::task_property_path::is_complete, false);
    EXPECT_STREQ(expected, restr.to_xml().c_str());
}

TEST(RestrictionTest, IsEqualToStringConstantRendersCorrectly)
{
    const char* expected = "<t:IsEqualTo>"
                           "<t:FieldURI FieldURI=\"folder:DisplayName\"/>"
                           "<t:FieldURIOrConstant>"
                           "<t:Constant Value=\"Inbox\"/>"
                           "</t:FieldURIOrConstant>"
                           "</t:IsEqualTo>";

    auto restr =
        ews::is_equal_to(ews::folder_property_path::display_name, "Inbox");
    EXPECT_STREQ(expected, restr.to_xml().c_str());
}

TEST(RestrictionTest, IsEqualToDateTimeConstantRendersCorrectly)
{
    const char* expected = "<t:IsEqualTo>"
                           "<t:FieldURI FieldURI=\"item:DateTimeCreated\"/>"
                           "<t:FieldURIOrConstant>"
                           "<t:Constant Value=\"2015-05-28T17:39:11Z\"/>"
                           "</t:FieldURIOrConstant>"
                           "</t:IsEqualTo>";

    const auto yesterday = ews::date_time("2015-05-28T17:39:11Z");
    auto restr =
        ews::is_equal_to(ews::item_property_path::date_time_created, yesterday);
    EXPECT_STREQ(expected, restr.to_xml().c_str());
}

TEST(RestrictionTest, IndexedFieldURIIsEqualToStringConstantRendersCorrectly)
{
    const char* expected = "<t:IsEqualTo>"
                           "<t:IndexedFieldURI "
                           "FieldURI=\"contacts:EmailAddress\" "
                           "FieldIndex=\"EmailAddress3\"/>"
                           "<t:FieldURIOrConstant>"
                           "<t:Constant Value=\"jane.dow@contoso.com\"/>"
                           "</t:FieldURIOrConstant>"
                           "</t:IsEqualTo>";

    auto restr = ews::is_equal_to(ews::contact_property_path::email_address_3,
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

    auto restr = ews::is_equal_to(ews::contact_property_path::email_address_1,
                                  "bruce@willis.com");
    EXPECT_STREQ(expected, restr.to_xml().c_str());
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
    EXPECT_STREQ(expected, restr.to_xml().c_str());
}

TEST(RestrictionTest, IsGreaterThanIntConstantRendersCorrectly)
{
    const char* expected = "<t:IsGreaterThan>"
                           "<t:FieldURI FieldURI=\"task:PercentComplete\"/>"
                           "<t:FieldURIOrConstant>"
                           "<t:Constant Value=\"75\"/>"
                           "</t:FieldURIOrConstant>"
                           "</t:IsGreaterThan>";

    auto restr =
        ews::is_greater_than(ews::task_property_path::percent_complete, 75);
    EXPECT_STREQ(expected, restr.to_xml().c_str());
}

TEST(RestrictionTest, IsGreaterThanOrEqualToIntConstantRendersCorrectly)
{
    const char* expected = "<t:IsGreaterThanOrEqualTo>"
                           "<t:FieldURI FieldURI=\"task:PercentComplete\"/>"
                           "<t:FieldURIOrConstant>"
                           "<t:Constant Value=\"80\"/>"
                           "</t:FieldURIOrConstant>"
                           "</t:IsGreaterThanOrEqualTo>";

    auto restr = ews::is_greater_than_or_equal_to(
        ews::task_property_path::percent_complete, 80);
    EXPECT_STREQ(expected, restr.to_xml().c_str());
}
} // namespace tests
