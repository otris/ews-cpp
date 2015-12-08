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
}

// vim:et ts=4 sw=4 noic cc=80
