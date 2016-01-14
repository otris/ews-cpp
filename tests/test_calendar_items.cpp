#include "fixtures.hpp"

#include <ews/rapidxml/rapidxml.hpp>

#include <vector>
#include <iterator>
#include <memory>
#include <algorithm>
#include <cstring>

namespace tests
{
    TEST(AttendeeTest, ToXML)
    {
        auto a = ews::attendee(
                    ews::mailbox("gaylord.focker@uchospitals.edu"),
                    ews::response_type::accept,
                    ews::date_time("2004-11-11T11:11:11Z"));

        EXPECT_STREQ("<y:Attendee>"
                       "<y:Mailbox>"
                         "<y:EmailAddress>gaylord.focker@uchospitals.edu</y:EmailAddress>"
                       "</y:Mailbox>"
                       "<y:ResponseType>Accept</y:ResponseType>"
                       "<y:LastResponseTime>2004-11-11T11:11:11Z</y:LastResponseTime>"
                     "</y:Attendee>",
                     a.to_xml("y").c_str());
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
        std::copy(xml, xml + std::strlen(xml), std::back_inserter(buf));
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
        EXPECT_EQ(0, cv.get_max_entries_returned());
        EXPECT_STREQ("<y:CalendarView StartDate=\"2016-01-12T10:00:00Z\" "
                     "EndDate=\"2016-01-12T12:00:00Z\" />",
                     cv.to_xml("y").c_str());
    }

    TEST(CalendarViewTest, ConstructWithMaxEntriesReturnedAttribute)
    {
        const auto start = ews::date_time("2016-01-12T10:00:00Z");
        const auto end = ews::date_time("2016-01-12T12:00:00Z");
        ews::calendar_view cv(start, end, 7);
        EXPECT_EQ(start, cv.get_start_date());
        EXPECT_EQ(end, cv.get_end_date());
        EXPECT_EQ(7, cv.get_max_entries_returned());
        EXPECT_STREQ("<m:CalendarView MaxEntriesReturned=\"7\" "
                     "StartDate=\"2016-01-12T10:00:00Z\" "
                     "EndDate=\"2016-01-12T12:00:00Z\" />",
                     cv.to_xml("m").c_str());
    }

    TEST(OccurrenceInfoTest, ConstructFromXML)
    {
        const char* xml =
            "<Occurrence>"
                "<ItemId Id=\"xyz\" ChangeKey=\"xyz\" />"
                "<Start>2011-11-11T11:11:11Z</Start>"
                "<End>2011-11-11T11:11:11Z</End>"
                "<OriginalStart>2011-11-11T11:11:11Z</OriginalStart>"
            "</Occurrence>";

        std::vector<char> buf;
        std::copy(xml, xml + std::strlen(xml), std::back_inserter(buf));
        buf.push_back('\0');
        rapidxml::xml_document<> doc;
        doc.parse<0>(&buf[0]);
        auto node = doc.first_node();

        auto a = ews::occurrence_info::from_xml_element(*node);
        EXPECT_EQ(ews::date_time("2011-11-11T11:11:11Z"), a.get_start());
        EXPECT_EQ(ews::date_time("2011-11-11T11:11:11Z"), a.get_end());
        EXPECT_EQ(ews::date_time("2011-11-11T11:11:11Z"), a.get_original_start());
    }

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
        vec.push_back(
            ews::attendee(
                ews::mailbox("gaylord.focker@uchospitals.edu"),
                ews::response_type::accept,
                ews::date_time("2004-11-11T11:11:11Z")));
        vec.push_back(
            ews::attendee(
                ews::mailbox("pam@nursery.org"),
                ews::response_type::no_response_received,
                ews::date_time("2004-12-24T08:00:00Z")));
        cal.set_required_attendees(vec);
        auto result = cal.get_required_attendees();
        ASSERT_FALSE(result.empty());
        EXPECT_TRUE(contains_if(result, [&](const ews::attendee& a)
        {
            return a.get_mailbox().value() == "pam@nursery.org"
                && a.get_response_type() == ews::response_type::no_response_received
                && a.get_last_response_time() == ews::date_time("2004-12-24T08:00:00Z");
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
        vec.push_back(
            ews::attendee(
                ews::mailbox("pam@nursery.org"),
                ews::response_type::accept,
                ews::date_time("2004-12-24T10:00:00Z")));
        auto prop = ews::property(
                ews::calendar_property_path::required_attendees,
                vec);
        auto new_id = service().update_item(cal.get_item_id(), prop);
        cal = service().get_calendar_item(new_id);
        EXPECT_EQ(initial_count + 1, cal.get_required_attendees().size());

        // Remove all again
        prop = ews::property(
            ews::calendar_property_path::required_attendees,
            std::vector<ews::attendee>());
        new_id = service().update_item(cal.get_item_id(), prop);
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
        vec.push_back(
            ews::attendee(
                ews::mailbox("gaylord.focker@uchospitals.edu"),
                ews::response_type::accept,
                ews::date_time("2004-11-11T11:11:11Z")));
        vec.push_back(
            ews::attendee(
                ews::mailbox("pam@nursery.org"),
                ews::response_type::no_response_received,
                ews::date_time("2004-12-24T08:00:00Z")));
        cal.set_optional_attendees(vec);
        auto result = cal.get_optional_attendees();
        ASSERT_FALSE(result.empty());
        EXPECT_TRUE(contains_if(result, [&](const ews::attendee& a)
        {
            return a.get_mailbox().value() == "pam@nursery.org"
                && a.get_response_type() == ews::response_type::no_response_received
                && a.get_last_response_time() == ews::date_time("2004-12-24T08:00:00Z");
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
        vec.push_back(
            ews::attendee(
                ews::mailbox("pam@nursery.org"),
                ews::response_type::accept,
                ews::date_time("2004-12-24T10:00:00Z")));
        auto prop = ews::property(
            ews::calendar_property_path::optional_attendees,
            vec);
        auto new_id = service().update_item(cal.get_item_id(), prop);
        cal = service().get_calendar_item(new_id);
        EXPECT_EQ(initial_count + 1, cal.get_optional_attendees().size());

        // Remove all again
        prop = ews::property(
            ews::calendar_property_path::optional_attendees,
            std::vector<ews::attendee>());
        new_id = service().update_item(cal.get_item_id(), prop);
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
        vec.push_back(
            ews::attendee(
                ews::mailbox("gaylord.focker@uchospitals.edu"),
                ews::response_type::accept,
                ews::date_time("2004-11-11T11:11:11Z")));
        vec.push_back(
            ews::attendee(
                ews::mailbox("pam@nursery.org"),
                ews::response_type::no_response_received,
                ews::date_time("2004-12-24T08:00:00Z")));
        cal.set_resources(vec);
        auto result = cal.get_resources();
        ASSERT_FALSE(result.empty());
        EXPECT_TRUE(contains_if(result, [&](const ews::attendee& a)
        {
            return a.get_mailbox().value() == "pam@nursery.org"
                && a.get_response_type() == ews::response_type::no_response_received
                && a.get_last_response_time() == ews::date_time("2004-12-24T08:00:00Z");
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
        vec.push_back(
            ews::attendee(
                ews::mailbox("pam@nursery.org"),
                ews::response_type::accept,
                ews::date_time("2004-12-24T10:00:00Z")));
        auto prop = ews::property(
            ews::calendar_property_path::resources,
            vec);
        auto new_id = service().update_item(cal.get_item_id(), prop);
        cal = service().get_calendar_item(new_id);
        EXPECT_EQ(initial_count + 1, cal.get_resources().size());

        // Remove all again
        prop = ews::property(
            ews::calendar_property_path::resources,
            std::vector<ews::attendee>());
        new_id = service().update_item(cal.get_item_id(), prop);
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
        EXPECT_TRUE(cal.get_timezone().empty());
    }

    // <AppointmentReplyTime/>
    TEST(OfflineCalendarItemTest, AppointmentReplyTimePropertyInitialValue)
    {
        auto cal = ews::calendar_item();
        EXPECT_FALSE(cal.get_appointment_reply_time().is_set());
    }

    // <AppointmentSequenceNumber/>
    TEST(OfflineCalendarItemTest,
         AppointmentSequenceNumberPropertyInitialValue)
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
        auto prop = ews::property(
            ews::calendar_property_path::conference_type, 2);
        auto new_id = service().update_item(cal.get_item_id(), prop);
        cal = service().get_calendar_item(new_id);
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
        cal = service().get_calendar_item(new_id);
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
        auto prop = ews::property(
            ews::calendar_property_path::is_online_meeting, true);
        auto new_id = service().update_item(cal.get_item_id(), prop);
        cal = service().get_calendar_item(new_id);
        EXPECT_TRUE(cal.is_online_meeting());

        prop = ews::property(
            ews::calendar_property_path::is_online_meeting, false);
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
        cal = service().get_calendar_item(new_id);
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

        ews::internal::on_scope_exit remove_items([&]
        {
            service().delete_item(idA);
            service().delete_item(idB);
            service().delete_item(idC);
        });

        const ews::distinguished_folder_id calendar_folder
            = ews::standard_folder::calendar;
        // 11 AM - 12 PM -> A, B
        auto view1 = ews::calendar_view(
                                ews::date_time("2016-01-12T11:00:00Z"),
                                ews::date_time("2016-01-12T12:00:00Z"));
        auto result = service().find_item(view1, calendar_folder);
        ASSERT_EQ(2, result.size());

        // 11:01 AM - 12 PM -> B
        auto view2 = ews::calendar_view(
                                ews::date_time("2016-01-12T11:01:00Z"),
                                ews::date_time("2016-01-12T12:00:00Z"));
        result = service().find_item(view2, calendar_folder);
        ASSERT_EQ(1, result.size());
        EXPECT_STREQ("Appointment B", result[0].get_subject().c_str());

        // 11 AM - 12:01 PM -> A, B, C
        auto view3 = ews::calendar_view(
                                ews::date_time("2016-01-12T11:00:00Z"),
                                ews::date_time("2016-01-12T12:01:00Z"));
        result = service().find_item(view3, calendar_folder);
        ASSERT_EQ(3, result.size());
    }
}

// vim:et ts=4 sw=4 noic cc=80
