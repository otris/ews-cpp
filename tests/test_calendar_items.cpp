
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

#include <algorithm>
#include <ews/rapidxml/rapidxml.hpp>
#include <ews/rapidxml/rapidxml_print.hpp>
#include <iterator>
#include <memory>
#include <sstream>
#include <string.h>
#include <string>
#include <vector>

namespace tests
{
TEST(AttendeeTest, ToXML)
{
    auto a = ews::attendee(ews::mailbox("gaylord.focker@uchospitals.edu"),
                           ews::response_type::accept,
                           ews::date_time("2004-11-11T11:11:11Z"));

    EXPECT_STREQ(
        "<t:Attendee>"
        "<t:Mailbox>"
        "<t:EmailAddress>gaylord.focker@uchospitals.edu</t:EmailAddress>"
        "</t:Mailbox>"
        "<t:ResponseType>Accept</t:ResponseType>"
        "<t:LastResponseTime>2004-11-11T11:11:11Z</t:LastResponseTime>"
        "</t:Attendee>",
        a.to_xml().c_str());
}

TEST(AttendeeTest, FromXML)
{
    const char* xml =
        "<Attendee>"
        "<Mailbox>"
        "<EmailAddress>gaylord.focker@uchospitals.edu</EmailAddress>"
        "</Mailbox>"
        "<ResponseType>Accept</ResponseType>"
        "<LastResponseTime>2004-11-11T11:11:11Z</LastResponseTime>"
        "</Attendee>";
    std::vector<char> buf;
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc;
    doc.parse<0>(&buf[0]);
    auto node = doc.first_node();
    auto a = ews::attendee::from_xml_element(*node);

    EXPECT_STREQ("gaylord.focker@uchospitals.edu",
                 a.get_mailbox().value().c_str());
    EXPECT_EQ(ews::response_type::accept, a.get_response_type());
    EXPECT_EQ(ews::date_time("2004-11-11T11:11:11Z"),
              a.get_last_response_time());
}

TEST(CalendarViewTest, Construct)
{
    const auto start = ews::date_time("2016-01-12T10:00:00Z");
    const auto end = ews::date_time("2016-01-12T12:00:00Z");
    ews::calendar_view cv(start, end);
    EXPECT_EQ(start, cv.get_start_date());
    EXPECT_EQ(end, cv.get_end_date());
    EXPECT_EQ(0U, cv.get_max_entries_returned());
    EXPECT_STREQ("<m:CalendarView StartDate=\"2016-01-12T10:00:00Z\" "
                 "EndDate=\"2016-01-12T12:00:00Z\" />",
                 cv.to_xml().c_str());
}

TEST(CalendarViewTest, ConstructWithMaxEntriesReturnedAttribute)
{
    const auto start = ews::date_time("2016-01-12T10:00:00Z");
    const auto end = ews::date_time("2016-01-12T12:00:00Z");
    ews::calendar_view cv(start, end, 7);
    EXPECT_EQ(start, cv.get_start_date());
    EXPECT_EQ(end, cv.get_end_date());
    EXPECT_EQ(7U, cv.get_max_entries_returned());
    EXPECT_STREQ("<m:CalendarView MaxEntriesReturned=\"7\" "
                 "StartDate=\"2016-01-12T10:00:00Z\" "
                 "EndDate=\"2016-01-12T12:00:00Z\" />",
                 cv.to_xml().c_str());
}

TEST(OccurrenceInfoTest, ConstructFromXML)
{
    const char* xml = "<Occurrence>"
                      "<ItemId Id=\"xyz\" ChangeKey=\"xyz\" />"
                      "<Start>2011-11-11T11:11:11Z</Start>"
                      "<End>2011-11-11T11:11:11Z</End>"
                      "<OriginalStart>2011-11-11T11:11:11Z</OriginalStart>"
                      "</Occurrence>";

    std::vector<char> buf;
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc;
    doc.parse<0>(&buf[0]);
    auto node = doc.first_node();

    auto a = ews::occurrence_info::from_xml_element(*node);
    EXPECT_EQ(ews::date_time("2011-11-11T11:11:11Z"), a.get_start());

    EXPECT_EQ(ews::date_time("2011-11-11T11:11:11Z"), a.get_end());

    EXPECT_EQ(ews::date_time("2011-11-11T11:11:11Z"), a.get_original_start());
}

TEST(OccurrenceInfoTest, DefaultConstruct)
{
    ews::occurrence_info a;
    EXPECT_TRUE(a.none());
}

TEST(RecurrenceRangeTest, NoEnd)
{
    const auto start_date = ews::date("1994-10-10");
    ews::no_end_recurrence_range r(start_date);
    EXPECT_EQ(start_date, r.get_start_date());

    const char* xml = "<Recurrence "
                      "xmlns:t=\"http://schemas.microsoft.com/exchange/"
                      "services/2006/types\"></Recurrence>";
    std::vector<char> buf;
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc;
    doc.parse<0>(&buf[0]);
    auto parent = doc.first_node();

    std::string str;
    rapidxml::print(std::back_inserter(str), r.to_xml_element(*parent),
                    rapidxml::print_no_indenting);

    EXPECT_STREQ("<t:NoEndRecurrence>"
                 "<t:StartDate>1994-10-10</t:StartDate>"
                 "</t:NoEndRecurrence>",
                 str.c_str());

    auto result = ews::recurrence_range::from_xml_element(*parent);
    ASSERT_TRUE(result);
    auto no_end_recurrence =
        dynamic_cast<ews::no_end_recurrence_range*>(result.get());
    ASSERT_TRUE(no_end_recurrence);
    EXPECT_EQ(start_date, no_end_recurrence->get_start_date());
}

TEST(RecurrenceRangeTest, EndDate)
{
    const auto start_date = ews::date("1961-08-13");
    const auto end_date = ews::date("1989-11-09");
    ews::end_date_recurrence_range r(start_date, end_date);
    EXPECT_EQ(start_date, r.get_start_date());
    EXPECT_EQ(end_date, r.get_end_date());

    const char* xml = "<Recurrence "
                      "xmlns:t=\"http://schemas.microsoft.com/exchange/"
                      "services/2006/types\"></Recurrence>";
    std::vector<char> buf;
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc;
    doc.parse<0>(&buf[0]);
    auto parent = doc.first_node();

    std::string str;
    rapidxml::print(std::back_inserter(str), r.to_xml_element(*parent),
                    rapidxml::print_no_indenting);

    EXPECT_STREQ("<t:EndDateRecurrence>"
                 "<t:StartDate>1961-08-13</t:StartDate>"
                 "<t:EndDate>1989-11-09</t:EndDate>"
                 "</t:EndDateRecurrence>",
                 str.c_str());

    auto result = ews::recurrence_range::from_xml_element(*parent);
    ASSERT_TRUE(result);
    auto end_date_recurrence_range =
        dynamic_cast<ews::end_date_recurrence_range*>(result.get());
    ASSERT_TRUE(end_date_recurrence_range);
    EXPECT_EQ(start_date, end_date_recurrence_range->get_start_date());
    EXPECT_EQ(end_date, end_date_recurrence_range->get_end_date());
}

TEST(RecurrenceRangeTest, Numbered)
{
    const auto start_date = ews::date("1989-01-01");
    ews::numbered_recurrence_range r(start_date, 18);
    EXPECT_EQ(start_date, r.get_start_date());
    EXPECT_EQ(18U, r.get_number_of_occurrences());

    const char* xml = "<Recurrence "
                      "xmlns:t=\"http://schemas.microsoft.com/exchange/"
                      "services/2006/types\"></Recurrence>";
    std::vector<char> buf;
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc;
    doc.parse<0>(&buf[0]);
    auto parent = doc.first_node();

    std::string str;
    rapidxml::print(std::back_inserter(str), r.to_xml_element(*parent),
                    rapidxml::print_no_indenting);

    EXPECT_STREQ("<t:NumberedRecurrence>"
                 "<t:StartDate>1989-01-01</t:StartDate>"
                 "<t:NumberOfOccurrences>18</t:NumberOfOccurrences>"
                 "</t:NumberedRecurrence>",
                 str.c_str());

    auto result = ews::recurrence_range::from_xml_element(*parent);
    ASSERT_TRUE(result);
    auto numbered_recurrence_range =
        dynamic_cast<ews::numbered_recurrence_range*>(result.get());
    ASSERT_TRUE(numbered_recurrence_range);
    EXPECT_EQ(start_date, numbered_recurrence_range->get_start_date());
    EXPECT_EQ(18U, numbered_recurrence_range->get_number_of_occurrences());
}

TEST(RecurrencePatternTest, AbsoluteYearly)
{
    ews::absolute_yearly_recurrence r(10, ews::month::oct);
    EXPECT_EQ(10U, r.get_day_of_month());
    EXPECT_EQ(ews::month::oct, r.get_month());

    const char* xml = "<Recurrence "
                      "xmlns:t=\"http://schemas.microsoft.com/exchange/"
                      "services/2006/types\"></Recurrence>";
    std::vector<char> buf;
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc;
    doc.parse<0>(&buf[0]);
    auto parent = doc.first_node();

    std::string str;
    rapidxml::print(std::back_inserter(str), r.to_xml_element(*parent),
                    rapidxml::print_no_indenting);
    EXPECT_STREQ("<t:AbsoluteYearlyRecurrence>"
                 "<t:DayOfMonth>10</t:DayOfMonth>"
                 "<t:Month>October</t:Month>"
                 "</t:AbsoluteYearlyRecurrence>",
                 str.c_str());
}

TEST(RecurrencePatternTest, RelativeYearly)
{
    ews::relative_yearly_recurrence r(
        ews::day_of_week::mon, ews::day_of_week_index::third, ews::month::apr);
    EXPECT_EQ(ews::day_of_week::mon, r.get_days_of_week());
    EXPECT_EQ(ews::day_of_week_index::third, r.get_day_of_week_index());
    EXPECT_EQ(ews::month::apr, r.get_month());

    EXPECT_STREQ("<t:RelativeYearlyRecurrence>"
                 "<t:DaysOfWeek>Monday</t:DaysOfWeek>"
                 "<t:DayOfWeekIndex>Third</t:DayOfWeekIndex>"
                 "<t:Month>April</t:Month>"
                 "</t:RelativeYearlyRecurrence>",
                 r.to_xml().c_str());

    const char* xml = "<Recurrence "
                      "xmlns:t=\"http://schemas.microsoft.com/exchange/"
                      "services/2006/types\"></Recurrence>";
    std::vector<char> buf;
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc;
    doc.parse<0>(&buf[0]);
    auto parent = doc.first_node();

    std::string str;
    rapidxml::print(std::back_inserter(str), r.to_xml_element(*parent),
                    rapidxml::print_no_indenting);

    EXPECT_STREQ("<t:RelativeYearlyRecurrence>"
                 "<t:DaysOfWeek>Monday</t:DaysOfWeek>"
                 "<t:DayOfWeekIndex>Third</t:DayOfWeekIndex>"
                 "<t:Month>April</t:Month>"
                 "</t:RelativeYearlyRecurrence>",
                 str.c_str());
}

TEST(RecurrencePatternTest, AbsoluteMonthly)
{
    ews::absolute_monthly_recurrence r(1, 5);
    EXPECT_EQ(1U, r.get_interval());
    EXPECT_EQ(5U, r.get_days_of_month());

    EXPECT_STREQ("<t:AbsoluteMonthlyRecurrence>"
                 "<t:Interval>1</t:Interval>"
                 "<t:DayOfMonth>5</t:DayOfMonth>"
                 "</t:AbsoluteMonthlyRecurrence>",
                 r.to_xml().c_str());

    const char* xml = "<Recurrence "
                      "xmlns:t=\"http://schemas.microsoft.com/exchange/"
                      "services/2006/types\"></Recurrence>";
    std::vector<char> buf;
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc;
    doc.parse<0>(&buf[0]);
    auto parent = doc.first_node();

    std::string str;
    rapidxml::print(std::back_inserter(str), r.to_xml_element(*parent),
                    rapidxml::print_no_indenting);

    EXPECT_STREQ("<t:AbsoluteMonthlyRecurrence>"
                 "<t:Interval>1</t:Interval>"
                 "<t:DayOfMonth>5</t:DayOfMonth>"
                 "</t:AbsoluteMonthlyRecurrence>",
                 str.c_str());
}

TEST(RecurrencePatternTest, RelativeMonthly)
{
    ews::relative_monthly_recurrence r(1, ews::day_of_week::thu,
                                       ews::day_of_week_index::third);
    EXPECT_EQ(1U, r.get_interval());
    EXPECT_EQ(ews::day_of_week::thu, r.get_days_of_week());
    EXPECT_EQ(ews::day_of_week_index::third, r.get_day_of_week_index());

    EXPECT_STREQ("<t:RelativeMonthlyRecurrence>"
                 "<t:Interval>1</t:Interval>"
                 "<t:DaysOfWeek>Thursday</t:DaysOfWeek>"
                 "<t:DayOfWeekIndex>Third</t:DayOfWeekIndex>"
                 "</t:RelativeMonthlyRecurrence>",
                 r.to_xml().c_str());

    const char* xml = "<Recurrence "
                      "xmlns:t=\"http://schemas.microsoft.com/exchange/"
                      "services/2006/types\"></Recurrence>";
    std::vector<char> buf;
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc;
    doc.parse<0>(&buf[0]);
    auto parent = doc.first_node();

    std::string str;
    rapidxml::print(std::back_inserter(str), r.to_xml_element(*parent),
                    rapidxml::print_no_indenting);

    EXPECT_STREQ("<t:RelativeMonthlyRecurrence>"
                 "<t:Interval>1</t:Interval>"
                 "<t:DaysOfWeek>Thursday</t:DaysOfWeek>"
                 "<t:DayOfWeekIndex>Third</t:DayOfWeekIndex>"
                 "</t:RelativeMonthlyRecurrence>",
                 str.c_str());
}

TEST(RecurrencePatternTest, Weekly)
{
    ews::weekly_recurrence r1(1, ews::day_of_week::mon);
    EXPECT_EQ(1U, r1.get_interval());
    ASSERT_EQ(1U, r1.get_days_of_week().size());
    EXPECT_EQ(ews::day_of_week::mon, r1.get_days_of_week().front());
    EXPECT_EQ(ews::day_of_week::mon, r1.get_first_day_of_week());

    EXPECT_STREQ("<t:WeeklyRecurrence>"
                 "<t:Interval>1</t:Interval>"
                 "<t:DaysOfWeek>Monday</t:DaysOfWeek>"
                 "<t:FirstDayOfWeek>Monday</t:FirstDayOfWeek>"
                 "</t:WeeklyRecurrence>",
                 r1.to_xml().c_str());

    const char* xml1 = "<Recurrence "
                       "xmlns:t=\"http://schemas.microsoft.com/exchange/"
                       "services/2006/types\"></Recurrence>";
    std::vector<char> buf;
    std::copy(xml1, xml1 + strlen(xml1), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc1;
    doc1.parse<0>(&buf[0]);
    auto parent = doc1.first_node();

    std::string str;
    rapidxml::print(std::back_inserter(str), r1.to_xml_element(*parent),
                    rapidxml::print_no_indenting);

    EXPECT_STREQ("<t:WeeklyRecurrence>"
                 "<t:Interval>1</t:Interval>"
                 "<t:DaysOfWeek>Monday</t:DaysOfWeek>"
                 "<t:FirstDayOfWeek>Monday</t:FirstDayOfWeek>"
                 "</t:WeeklyRecurrence>",
                 str.c_str());

    // On multiple days
    auto days = std::vector<ews::day_of_week>();
    days.push_back(ews::day_of_week::thu);
    days.push_back(ews::day_of_week::fri);
    ews::weekly_recurrence r2(2, days, ews::day_of_week::sun);
    EXPECT_EQ(2U, r2.get_interval());
    ASSERT_EQ(2U, r2.get_days_of_week().size());
    EXPECT_EQ(ews::day_of_week::thu, r2.get_days_of_week()[0]);
    EXPECT_EQ(ews::day_of_week::fri, r2.get_days_of_week()[1]);
    EXPECT_EQ(ews::day_of_week::sun, r2.get_first_day_of_week());

    const char* xml2 = "<Recurrence "
                       "xmlns:t=\"http://schemas.microsoft.com/exchange/"
                       "services/2006/types\"></Recurrence>";
    buf.clear();
    std::copy(xml2, xml2 + strlen(xml2), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc2;
    doc2.parse<0>(&buf[0]);
    parent = doc2.first_node();

    str.clear();
    rapidxml::print(std::back_inserter(str), r2.to_xml_element(*parent),
                    rapidxml::print_no_indenting);

    EXPECT_STREQ("<t:WeeklyRecurrence>"
                 "<t:Interval>2</t:Interval>"
                 "<t:DaysOfWeek>Thursday Friday</t:DaysOfWeek>"
                 "<t:FirstDayOfWeek>Sunday</t:FirstDayOfWeek>"
                 "</t:WeeklyRecurrence>",
                 str.c_str());
}

TEST(RecurrencePatternTest, Daily)
{
    ews::daily_recurrence r(3);
    EXPECT_EQ(3U, r.get_interval());

    EXPECT_STREQ("<t:DailyRecurrence>"
                 "<t:Interval>3</t:Interval>"
                 "</t:DailyRecurrence>",
                 r.to_xml().c_str());

    const char* xml = "<Recurrence "
                      "xmlns:t=\"http://schemas.microsoft.com/exchange/"
                      "services/2006/types\"></Recurrence>";
    std::vector<char> buf;
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    rapidxml::xml_document<> doc;
    doc.parse<0>(&buf[0]);
    auto parent = doc.first_node();

    std::string str;
    rapidxml::print(std::back_inserter(str), r.to_xml_element(*parent),
                    rapidxml::print_no_indenting);

    EXPECT_STREQ("<t:DailyRecurrence>"
                 "<t:Interval>3</t:Interval>"
                 "</t:DailyRecurrence>",
                 str.c_str());
}

TEST_F(CalendarItemTest, GetCalendarItemWithInvalidIdThrows)
{
    auto invalid_id = ews::item_id();
    EXPECT_THROW(service().get_calendar_item(invalid_id), ews::exchange_error);
}

TEST_F(CalendarItemTest, GetCalendarItemWithInvalidIdExceptionResponse)
{
    auto invalid_id = ews::item_id();
    try
    {
        service().get_calendar_item(invalid_id);
        FAIL() << "Expected an exception";
    }
    catch (ews::exchange_error& exc)
    {
        EXPECT_EQ(ews::response_code::error_invalid_id_empty, exc.code());
    }
}

TEST_F(CalendarItemTest, CreateAndDeleteCalendarItem)
{
    const ews::distinguished_folder_id calendar_folder =
        ews::standard_folder::calendar;
    const auto initial_count = service().find_item(calendar_folder).size();

    auto calitem = ews::calendar_item();
    calitem.set_subject("Write chapter explaining Vogon poetry");
    ews::body calbody("What is six times seven?");
    calitem.set_body(calbody);

    {
        ews::internal::on_scope_exit remove_item([&] {
            service().delete_calendar_item(
                std::move(calitem), ews::delete_type::hard_delete,
                ews::send_meeting_cancellations::send_to_none);
        });

        auto item_id = service().create_item(calitem);

        calitem = service().get_calendar_item(item_id,
                                              ews::base_shape::all_properties);
        auto subject = calitem.get_subject();
        EXPECT_STREQ("Write chapter explaining Vogon poetry", subject.c_str());
        auto body = calitem.get_body();
        EXPECT_STREQ("What is six times seven?", body.content().c_str());
    }

    // Check sink argument
    EXPECT_TRUE(calitem.get_subject().empty());

    auto items = service().find_item(calendar_folder);
    EXPECT_EQ(initial_count, items.size());
}

// <Start/>
TEST(OfflineCalendarItemTest, StartPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.get_start().is_set());
}

TEST(OfflineCalendarItemTest, SetStartProperty)
{
    auto cal = ews::calendar_item();
    auto start = ews::date_time("2015-12-10T10:57:26.000Z");
    cal.set_start(start);
    EXPECT_EQ(start, cal.get_start());
}

TEST_F(CalendarItemTest, UpdateStartProperty)
{
    const auto new_start = ews::date_time("2004-12-25T11:00:00Z");
    auto cal = test_calendar_item();
    auto prop = ews::property(ews::calendar_property_path::start, new_start);
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id);
    EXPECT_EQ(new_start, cal.get_start());
}

// <End/>
TEST(OfflineCalendarItemTest, EndPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.get_end().is_set());
}

TEST(OfflineCalendarItemTest, SetEndProperty)
{
    auto cal = ews::calendar_item();
    auto end = ews::date_time("2015-12-10T10:57:26.000Z");
    cal.set_end(end);
    EXPECT_EQ(end, cal.get_end());
}

TEST_F(CalendarItemTest, UpdateEndProperty)
{
    const auto new_end = ews::date_time("2004-12-28T10:00:00Z");
    auto cal = test_calendar_item();
    auto prop = ews::property(ews::calendar_property_path::end, new_end);
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id);
    EXPECT_EQ(new_end, cal.get_end());
}

// <OriginalStart/>
TEST(OfflineCalendarItemTest, OriginalStartPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.get_original_start().is_set());
}

// <IsAllDayEvent/>
TEST(OfflineCalendarItemTest, IsAllDayEventPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.is_all_day_event());
}

TEST(OfflineCalendarItemTest, SetIsAllDayEventProperty)
{
    auto cal = ews::calendar_item();
    cal.set_all_day_event_enabled(true);
    EXPECT_TRUE(cal.is_all_day_event());
}

TEST_F(CalendarItemTest, UpdateIsAllDayEventProperty)
{
    auto cal = test_calendar_item();
    auto prop =
        ews::property(ews::calendar_property_path::is_all_day_event, true);
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id, ews::base_shape::all_properties);
    EXPECT_TRUE(cal.is_all_day_event());
}

// <LegacyFreeBusyStatus/>
TEST(OfflineCalendarItemTest, LegacyFreeBusyStatusPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_EQ(ews::free_busy_status::busy, cal.get_legacy_free_busy_status());
}

TEST(OfflineCalendarItemTest, SetLegacyFreeBusyStatusProperty)
{
    auto cal = ews::calendar_item();
    cal.set_legacy_free_busy_status(ews::free_busy_status::out_of_office);
    EXPECT_EQ(ews::free_busy_status::out_of_office,
              cal.get_legacy_free_busy_status());
}

TEST_F(CalendarItemTest, UpdateLegacyFreeBusyStatusProperty)
{
    auto cal = test_calendar_item();
    auto prop =
        ews::property(ews::calendar_property_path::legacy_free_busy_status,
                      ews::free_busy_status::out_of_office);
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id);
    EXPECT_EQ(ews::free_busy_status::out_of_office,
              cal.get_legacy_free_busy_status());
}

// <Location/>
TEST(OfflineCalendarItemTest, LocationPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_location().empty());
}

TEST(OfflineCalendarItemTest, SetLocationProperty)
{
    auto cal = ews::calendar_item();
    cal.set_location("Their place");
    EXPECT_STREQ("Their place", cal.get_location().c_str());
}

TEST_F(CalendarItemTest, UpdateLocationProperty)
{
    auto cal = test_calendar_item();
    auto prop =
        ews::property(ews::calendar_property_path::location, "Our place");
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id);
    EXPECT_STREQ("Our place", cal.get_location().c_str());
}

// <When/>
TEST(OfflineCalendarItemTest, WhenPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_when().empty());
}

TEST(OfflineCalendarItemTest, SetWhenProperty)
{
    auto cal = ews::calendar_item();
    cal.set_when("Before we get married");
    EXPECT_STREQ("Before we get married", cal.get_when().c_str());
}

TEST_F(CalendarItemTest, UpdateWhenProperty)
{
    auto cal = test_calendar_item();
    auto prop =
        ews::property(ews::calendar_property_path::when, "Next Christmas");
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Next Christmas", cal.get_when().c_str());
}

// <IsMeeting/>
TEST(OfflineCalendarItemTest, IsMeetingPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.is_meeting());
}

// <IsCancelled/>
TEST(OfflineCalendarItemTest, IsCancelledPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.is_cancelled());
}

// <IsRecurring/>
TEST(OfflineCalendarItemTest, IsRecurringPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.is_recurring());
}

// <MeetingRequestWasSent/>
TEST(OfflineCalendarItemTest, MeetingRequestWasSentPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.meeting_request_was_sent());
}

// <IsResponseRequested/>
TEST(OfflineCalendarItemTest, IsResponseRequestedPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.is_response_requested());
}

// <CalendarItemType/>
TEST(OfflineCalendarItemTest, CalendarItemTypePropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_EQ(ews::calendar_item_type::single, cal.get_calendar_item_type());
}

// <MyResponseType/>
TEST(OfflineCalendarItemTest, MyResponseTypePropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_EQ(ews::response_type::unknown, cal.get_my_response_type());
}

// <Organizer/>
TEST(OfflineCalendarItemTest, OrganizerPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_organizer().none());
}

// <RequiredAttendees/>
TEST(OfflineCalendarItemTest, RequiredAttendeesPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_required_attendees().empty());
}

TEST(OfflineCalendarItemTest, SetRequiredAttendeesProperty)
{
    auto cal = ews::calendar_item();
    const auto empty_vec = std::vector<ews::attendee>();
    cal.set_required_attendees(empty_vec);
    EXPECT_TRUE(cal.get_required_attendees().empty());

    std::vector<ews::attendee> vec;
    vec.push_back(ews::attendee(ews::mailbox("gaylord.focker@uchospitals.edu"),
                                ews::response_type::accept,
                                ews::date_time("2004-11-11T11:11:11Z")));
    vec.push_back(ews::attendee(ews::mailbox("pam@nursery.org"),
                                ews::response_type::no_response_received,
                                ews::date_time("2004-12-24T08:00:00Z")));
    cal.set_required_attendees(vec);
    auto result = cal.get_required_attendees();
    ASSERT_FALSE(result.empty());
    EXPECT_TRUE(contains_if(result, [&](const ews::attendee& a) {
        return a.get_mailbox().value() == "pam@nursery.org" &&
               a.get_response_type() ==
                   ews::response_type::no_response_received &&
               a.get_last_response_time() ==
                   ews::date_time("2004-12-24T08:00:00Z");
    }));

    // Finally, remove all
    cal.set_required_attendees(empty_vec);
    EXPECT_TRUE(cal.get_required_attendees().empty());
}

TEST_F(CalendarItemTest, UpdateRequiredAttendeesProperty)
{
    // Add one
    auto cal = test_calendar_item();
    auto vec = cal.get_required_attendees();
    const auto initial_count = vec.size();
    vec.push_back(ews::attendee(ews::mailbox("pam@nursery.org"),
                                ews::response_type::accept,
                                ews::date_time("2004-12-24T10:00:00Z")));
    auto prop =
        ews::property(ews::calendar_property_path::required_attendees, vec);
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id);
    EXPECT_EQ(initial_count + 1, cal.get_required_attendees().size());

    // Remove all again
    prop = ews::property(ews::calendar_property_path::required_attendees,
                         std::vector<ews::attendee>());
    auto update = ews::update(prop, ews::update::operation::delete_item_field);
    new_id = service().update_item(cal.get_item_id(), update);
    cal = service().get_calendar_item(new_id);
    EXPECT_TRUE(cal.get_required_attendees().empty());
}

// <OptionalAttendees/>
TEST(OfflineCalendarItemTest, OptionalAttendeesPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_optional_attendees().empty());
}

TEST(OfflineCalendarItemTest, SetOptionalAttendeesProperty)
{
    auto cal = ews::calendar_item();
    const auto empty_vec = std::vector<ews::attendee>();
    cal.set_optional_attendees(empty_vec);
    EXPECT_TRUE(cal.get_optional_attendees().empty());

    std::vector<ews::attendee> vec;
    vec.push_back(ews::attendee(ews::mailbox("gaylord.focker@uchospitals.edu"),
                                ews::response_type::accept,
                                ews::date_time("2004-11-11T11:11:11Z")));
    vec.push_back(ews::attendee(ews::mailbox("pam@nursery.org"),
                                ews::response_type::no_response_received,
                                ews::date_time("2004-12-24T08:00:00Z")));
    cal.set_optional_attendees(vec);
    auto result = cal.get_optional_attendees();
    ASSERT_FALSE(result.empty());
    EXPECT_TRUE(contains_if(result, [&](const ews::attendee& a) {
        return a.get_mailbox().value() == "pam@nursery.org" &&
               a.get_response_type() ==
                   ews::response_type::no_response_received &&
               a.get_last_response_time() ==
                   ews::date_time("2004-12-24T08:00:00Z");
    }));

    // Finally, remove all
    cal.set_optional_attendees(empty_vec);
    EXPECT_TRUE(cal.get_optional_attendees().empty());
}

TEST_F(CalendarItemTest, UpdateOptionalAttendeesProperty)
{
    // Add one
    auto cal = test_calendar_item();
    auto vec = cal.get_optional_attendees();
    const auto initial_count = vec.size();
    vec.push_back(ews::attendee(ews::mailbox("pam@nursery.org"),
                                ews::response_type::accept,
                                ews::date_time("2004-12-24T10:00:00Z")));
    auto prop =
        ews::property(ews::calendar_property_path::optional_attendees, vec);
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id);
    EXPECT_EQ(initial_count + 1, cal.get_optional_attendees().size());

    // Remove all again
    prop = ews::property(ews::calendar_property_path::optional_attendees,
                         std::vector<ews::attendee>());
    auto update = ews::update(prop, ews::update::operation::delete_item_field);
    new_id = service().update_item(cal.get_item_id(), update);
    cal = service().get_calendar_item(new_id);
    EXPECT_TRUE(cal.get_optional_attendees().empty());
}

// <Resources/>
TEST(OfflineCalendarItemTest, ResourcesPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_resources().empty());
}

TEST(OfflineCalendarItemTest, SetResourcesProperty)
{
    auto cal = ews::calendar_item();
    const auto empty_vec = std::vector<ews::attendee>();
    cal.set_resources(empty_vec);
    EXPECT_TRUE(cal.get_resources().empty());

    std::vector<ews::attendee> vec;
    vec.push_back(ews::attendee(ews::mailbox("gaylord.focker@uchospitals.edu"),
                                ews::response_type::accept,
                                ews::date_time("2004-11-11T11:11:11Z")));
    vec.push_back(ews::attendee(ews::mailbox("pam@nursery.org"),
                                ews::response_type::no_response_received,
                                ews::date_time("2004-12-24T08:00:00Z")));
    cal.set_resources(vec);
    auto result = cal.get_resources();
    ASSERT_FALSE(result.empty());
    EXPECT_TRUE(contains_if(result, [&](const ews::attendee& a) {
        return a.get_mailbox().value() == "pam@nursery.org" &&
               a.get_response_type() ==
                   ews::response_type::no_response_received &&
               a.get_last_response_time() ==
                   ews::date_time("2004-12-24T08:00:00Z");
    }));

    // Finally, remove all
    cal.set_resources(empty_vec);
    EXPECT_TRUE(cal.get_resources().empty());
}

TEST_F(CalendarItemTest, UpdateResourcesProperty)
{
    // Add one
    auto cal = test_calendar_item();
    auto vec = cal.get_resources();
    const auto initial_count = vec.size();
    vec.push_back(ews::attendee(ews::mailbox("pam@nursery.org"),
                                ews::response_type::accept,
                                ews::date_time("2004-12-24T10:00:00Z")));
    auto prop = ews::property(ews::calendar_property_path::resources, vec);
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id);
    EXPECT_EQ(initial_count + 1, cal.get_resources().size());

    // Remove all again
    prop = ews::property(ews::calendar_property_path::resources,
                         std::vector<ews::attendee>());
    auto update = ews::update(prop, ews::update::operation::delete_item_field);
    new_id = service().update_item(cal.get_item_id(), update);
    cal = service().get_calendar_item(new_id);
    EXPECT_TRUE(cal.get_resources().empty());
}

// <ConflictingMeetingCount/>
TEST(OfflineCalendarItemTest, ConflictingMeetingCountPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_EQ(0, cal.get_conflicting_meeting_count());
}

// <AdjacentMeetingCount/>
TEST(OfflineCalendarItemTest, AdjacentMeetingCountPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_EQ(0, cal.get_adjacent_meeting_count());
}

// <Duration/>
TEST(OfflineCalendarItemTest, DurationPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.get_duration().is_set());
}

TEST_F(CalendarItemTest, GetDurationProperty)
{
    auto cal = test_calendar_item();
    EXPECT_TRUE(cal.get_duration().is_set());
}

// <TimeZone/>
TEST(OfflineCalendarItemTest, TimeZonePropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_time_zone().empty());
}

// <AppointmentReplyTime/>
TEST(OfflineCalendarItemTest, AppointmentReplyTimePropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.get_appointment_reply_time().is_set());
}

// <AppointmentSequenceNumber/>
TEST(OfflineCalendarItemTest, AppointmentSequenceNumberPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_EQ(0, cal.get_appointment_sequence_number());
}

// <AppointmentState/>
TEST(OfflineCalendarItemTest, AppointmentStatePropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_EQ(0, cal.get_appointment_state());
}

// <Recurrence/>
TEST(OfflineCalendarItemTest, RecurrencePropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.get_recurrence().first);
    EXPECT_FALSE(cal.get_recurrence().second);

    // Set
    ews::absolute_yearly_recurrence birthday(10, ews::month::oct);
    auto start_date = ews::date_time("1994-10-10");
    ews::no_end_recurrence_range no_end(start_date);

    cal.set_recurrence(birthday, no_end);
    auto result = cal.get_recurrence();
    ASSERT_TRUE(result.first && result.second);
    auto pattern1 =
        dynamic_cast<ews::absolute_yearly_recurrence*>(result.first.get());
    ASSERT_TRUE(pattern1);
    EXPECT_EQ(10U, pattern1->get_day_of_month());
    EXPECT_EQ(ews::month::oct, pattern1->get_month());
    auto range1 =
        dynamic_cast<ews::no_end_recurrence_range*>(result.second.get());
    ASSERT_TRUE(range1);
    EXPECT_EQ(start_date, range1->get_start_date());

    // Replace
    ews::absolute_monthly_recurrence mortgage_payment(1, 5);
    start_date = ews::date_time("2016-01-01");
    ews::numbered_recurrence_range end(start_date, 48);

    cal.set_recurrence(mortgage_payment, end);
    result = cal.get_recurrence();
    ASSERT_TRUE(result.first && result.second);
    auto pattern2 =
        dynamic_cast<ews::absolute_monthly_recurrence*>(result.first.get());
    ASSERT_TRUE(pattern2);
    EXPECT_EQ(1U, pattern2->get_interval());
    EXPECT_EQ(5U, pattern2->get_days_of_month());
    auto range2 =
        dynamic_cast<ews::numbered_recurrence_range*>(result.second.get());
    ASSERT_TRUE(range2);
    EXPECT_EQ(start_date, range2->get_start_date());
    EXPECT_EQ(48U, range2->get_number_of_occurrences());
}

TEST_F(CalendarItemTest, GetRecurrenceProperty)
{
    // From item that is not part of a series
    auto cal = test_calendar_item();
    auto recurrence = cal.get_recurrence();
    EXPECT_FALSE(recurrence.first && recurrence.second);
}

TEST_F(CalendarItemTest, CreateRecurringSeries)
{
    auto master = ews::calendar_item();
    master.set_subject("Monthly Mortgage Payment is due");
    master.set_start(ews::date_time("2014-12-01T00:00:00Z"));
    master.set_end(ews::date_time("2014-12-01T00:05:00Z"));
    master.set_recurrence(
        ews::absolute_monthly_recurrence(1, 5),
        ews::end_date_recurrence_range(ews::date("2015-01-01Z"),
                                       ews::date("2037-01-01Z")));

    auto master_id = service().create_item(master);
    ews::internal::on_scope_exit remove_items(
        [&] { service().delete_item(master_id); });
    master =
        service().get_calendar_item(master_id, ews::base_shape::all_properties);

    auto recurrence = master.get_recurrence();
    ASSERT_TRUE(recurrence.first && recurrence.second);
    auto pattern =
        dynamic_cast<ews::absolute_monthly_recurrence*>(recurrence.first.get());
    ASSERT_TRUE(pattern);
    EXPECT_EQ(1U, pattern->get_interval());
    EXPECT_EQ(5U, pattern->get_days_of_month());
    auto range =
        dynamic_cast<ews::end_date_recurrence_range*>(recurrence.second.get());
    ASSERT_TRUE(range);
    EXPECT_EQ(ews::date("2015-01-05Z"), range->get_start_date());
    EXPECT_EQ(ews::date("2037-01-01Z"), range->get_end_date());
}

TEST_F(CalendarItemTest, DISABLED_UpdateRecurringSeries)
{
    auto master = ews::calendar_item();
    master.set_subject("Monthly Mortgage Payment is due");
    master.set_start(ews::date_time("2015-12-01T00:00:00Z"));
    master.set_end(ews::date_time("2015-12-01T00:05:00Z"));
    master.set_recurrence(
        ews::absolute_monthly_recurrence(1, 5),
        ews::end_date_recurrence_range(ews::date("2016-01-01Z"),
                                       ews::date("2037-01-01Z")));

    auto master_id = service().create_item(master);
    ews::internal::on_scope_exit remove_items(
        [&] { service().delete_item(master_id); });
    master = service().get_calendar_item(master_id);
    EXPECT_FALSE(master.is_recurring());
    EXPECT_EQ(ews::calendar_item_type::recurring_master,
              master.get_calendar_item_type());

    auto prop = ews::property(
        ews::calendar_property_path::recurrence,
        ews::absolute_monthly_recurrence(1, 2),
        ews::numbered_recurrence_range(ews::date("2016-01-01Z"), 4));

    auto new_id = service().update_item(master.get_item_id(), prop);
    master = service().get_calendar_item(new_id);
    auto recurrence = master.get_recurrence();
    auto pattern =
        dynamic_cast<ews::absolute_monthly_recurrence*>(recurrence.first.get());
    ASSERT_TRUE(pattern);
    EXPECT_EQ(2U, pattern->get_days_of_month());
}

// <FirstOccurrence/>
TEST(OfflineCalendarItemTest, FirstOccurrencePropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_first_occurrence().none());
}

// <LastOccurrence/>
TEST(OfflineCalendarItemTest, LastOccurrencePropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_last_occurrence().none());
}

// <ModifiedOccurrences/>
TEST(OfflineCalendarItemTest, ModifiedOccurrencesPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_modified_occurrences().empty());
}

// <DeletedOccurrences/>
TEST(OfflineCalendarItemTest, DeletedOccurrencesPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_deleted_occurrences().empty());
}

// <ConferenceType/>
TEST(OfflineCalendarItemTest, ConferenceTypePropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_EQ(0, cal.get_conference_type());
}

TEST(OfflineCalendarItemTest, SetConferenceTypeProperty)
{
    auto cal = ews::calendar_item();
    cal.set_conference_type(1);
    EXPECT_EQ(1, cal.get_conference_type());
}

TEST_F(CalendarItemTest, UpdateConferenceTypeProperty)
{
    auto cal = test_calendar_item();
    auto prop = ews::property(ews::calendar_property_path::conference_type, 2);
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id, ews::base_shape::all_properties);
    EXPECT_EQ(2, cal.get_conference_type());
}

// <AllowNewTimeProposal/>
TEST(OfflineCalendarItemTest, AllowNewTimeProposalPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.is_new_time_proposal_allowed());
}

TEST(OfflineCalendarItemTest, SetAllowNewTimeProposalProperty)
{
    auto cal = ews::calendar_item();
    cal.set_new_time_proposal_allowed(true);
    EXPECT_TRUE(cal.is_new_time_proposal_allowed());
    cal.set_new_time_proposal_allowed(false);
    EXPECT_FALSE(cal.is_new_time_proposal_allowed());
}

TEST_F(CalendarItemTest, UpdateAllowNewTimeProposalProperty)
{
    auto cal = test_calendar_item();
    auto prop = ews::property(
        ews::calendar_property_path::allow_new_time_proposal, true);
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id, ews::base_shape::all_properties);
    EXPECT_TRUE(cal.is_new_time_proposal_allowed());
}

// <IsOnlineMeeting/>
TEST(OfflineCalendarItemTest, IsOnlineMeetingPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_FALSE(cal.is_online_meeting());
}

TEST(OfflineCalendarItemTest, SetIsOnlineMeetingProperty)
{
    auto cal = ews::calendar_item();
    cal.set_online_meeting_enabled(true);
    EXPECT_TRUE(cal.is_online_meeting());
    cal.set_online_meeting_enabled(false);
    EXPECT_FALSE(cal.is_online_meeting());
}

TEST_F(CalendarItemTest, UpdateIsOnlineMeetingProperty)
{
    auto cal = test_calendar_item();
    auto prop =
        ews::property(ews::calendar_property_path::is_online_meeting, true);
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id, ews::base_shape::all_properties);
    EXPECT_TRUE(cal.is_online_meeting());

    prop = ews::property(ews::calendar_property_path::is_online_meeting, false);
    new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id);
    EXPECT_FALSE(cal.is_online_meeting());
}

// <MeetingWorkspaceUrl/>
TEST(OfflineCalendarItemTest, MeetingWorkspaceUrlPropertyInitialValue)
{
    auto cal = ews::calendar_item();
    EXPECT_TRUE(cal.get_meeting_workspace_url().empty());
}

TEST(OfflineCalendarItemTest, SetMeetingWorkspaceUrlProperty)
{
    auto cal = ews::calendar_item();
    cal.set_meeting_workspace_url("kitchen");
    EXPECT_STREQ("kitchen", cal.get_meeting_workspace_url().c_str());

    cal.set_meeting_workspace_url("13th floor");
    EXPECT_STREQ("13th floor", cal.get_meeting_workspace_url().c_str());
}

TEST_F(CalendarItemTest, UpdateMeetingWorkspaceUrlProperty)
{
    auto cal = test_calendar_item();
    auto prop = ews::property(
        ews::calendar_property_path::meeting_workspace_url, "kitchen");
    auto new_id = service().update_item(cal.get_item_id(), prop);
    cal = service().get_calendar_item(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("kitchen", cal.get_meeting_workspace_url().c_str());
}

TEST_F(CalendarItemTest, FindCalendarItemsWithCalendarViews)
{
    // 10 AM - 11 AM
    auto calitemA = ews::calendar_item();
    calitemA.set_subject("Appointment A");
    calitemA.set_start(ews::date_time("2016-01-12T10:00:00Z"));
    calitemA.set_end(ews::date_time("2016-01-12T11:00:00Z"));
    auto idA = service().create_item(calitemA);

    // 11 AM - 12 PM
    auto calitemB = ews::calendar_item();
    calitemB.set_subject("Appointment B");
    calitemB.set_start(ews::date_time("2016-01-12T11:00:00Z"));
    calitemB.set_end(ews::date_time("2016-01-12T12:00:00Z"));
    auto idB = service().create_item(calitemB);

    // 12 PM - 1 PM
    auto calitemC = ews::calendar_item();
    calitemC.set_subject("Appointment C");
    calitemC.set_start(ews::date_time("2016-01-12T12:00:00Z"));
    calitemC.set_end(ews::date_time("2016-01-12T13:00:00Z"));
    auto idC = service().create_item(calitemC);

    ews::internal::on_scope_exit remove_items([&] {
        service().delete_item(idA);
        service().delete_item(idB);
        service().delete_item(idC);
    });

    const ews::distinguished_folder_id calendar_folder =
        ews::standard_folder::calendar;
    // 11 AM - 12 PM -> A, B
    auto view1 = ews::calendar_view(ews::date_time("2016-01-12T11:00:00Z"),
                                    ews::date_time("2016-01-12T12:00:00Z"));
    auto result = service().find_item(view1, calendar_folder);
    ASSERT_EQ(2U, result.size());

    // 11:01 AM - 12 PM -> B
    auto view2 = ews::calendar_view(ews::date_time("2016-01-12T11:01:00Z"),
                                    ews::date_time("2016-01-12T12:00:00Z"));
    result = service().find_item(view2, calendar_folder);
    ASSERT_EQ(1U, result.size());
    EXPECT_STREQ("Appointment B", result[0].get_subject().c_str());

    // 11 AM - 12:01 PM -> A, B, C
    auto view3 = ews::calendar_view(ews::date_time("2016-01-12T11:00:00Z"),
                                    ews::date_time("2016-01-12T12:01:00Z"));
    result = service().find_item(view3, calendar_folder);
    ASSERT_EQ(3U, result.size());
}
} // namespace tests
