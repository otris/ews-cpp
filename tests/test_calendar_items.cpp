#include "fixtures.hpp"

namespace tests
{
    TEST_F(CalendarItemTest, GetCalendarItemWithInvalidIdThrows)
    {
        auto invalid_id = ews::item_id();
        EXPECT_THROW(service().get_calendar_item(invalid_id),
                     ews::exchange_error);
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
            EXPECT_STREQ("ErrorInvalidIdEmpty", exc.what());
        }
    }

    TEST_F(CalendarItemTest, CreateAndDeleteCalendarItem)
    {
        const ews::distinguished_folder_id calendar_folder =
            ews::standard_folder::calendar;
        const auto initial_count = service().find_item(calendar_folder).size();

        auto calitem = ews::calendar_item();
        calitem.set_subject("Write chapter explaining Vogon poetry");

        {
            ews::internal::on_scope_exit remove_item([&]
            {
                service().delete_calendar_item(
                               std::move(calitem),
                               ews::delete_type::hard_delete,
                               ews::send_meeting_cancellations::send_to_none);
            });

            auto item_id = service().create_item(calitem);

            calitem = service().get_calendar_item(item_id);
            auto subject = calitem.get_subject();
            EXPECT_STREQ("Write chapter explaining Vogon poetry",
                         subject.c_str());
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
        auto prop = ews::property(
                        ews::calendar_property_path::is_all_day_event,
                        true);
        auto new_id = service().update_item(cal.get_item_id(), prop);
        cal = service().get_calendar_item(new_id);
        EXPECT_TRUE(cal.is_all_day_event());
    }

    // <LegacyFreeBusyStatus/>
    TEST(OfflineCalendarItemTest, LegacyFreeBusyStatusPropertyInitialValue)
    {
        auto cal = ews::calendar_item();
        EXPECT_EQ(ews::free_busy_status::busy,
                  cal.get_legacy_free_busy_status());
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
        auto prop = ews::property(
                        ews::calendar_property_path::legacy_free_busy_status,
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
        auto prop = ews::property(
                        ews::calendar_property_path::location,
                        "Our place");
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
        auto prop = ews::property(
                                ews::calendar_property_path::when,
                                "Next Christmas");
        auto new_id = service().update_item(cal.get_item_id(), prop);
        cal = service().get_calendar_item(new_id);
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
        EXPECT_EQ(ews::calendar_item_type::single,
            cal.get_calendar_item_type());
    }

    // <MyResponseType/>
    TEST(OfflineCalendarItemTest, MyResponseTypePropertyInitialValue)
    {
        auto cal = ews::calendar_item();
        EXPECT_EQ(ews::response_type::unknown, cal.get_my_response_type());
    }
}

// vim:et ts=4 sw=4 noic cc=80
