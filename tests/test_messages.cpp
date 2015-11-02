#include "fixtures.hpp"

namespace tests
{
    TEST_F(MessageTest, GetMessageWithInvalidIdThrows)
    {
        auto invalid_id = ews::item_id();
        EXPECT_THROW(service().get_message(invalid_id), ews::exchange_error);
    }

    TEST_F(MessageTest, GetMessageWithInvalidIdExceptionResponse)
    {
        auto invalid_id = ews::item_id();
        try
        {
            service().get_message(invalid_id);
            FAIL() << "Expected an exception";
        }
        catch (ews::exchange_error& exc)
        {
            EXPECT_EQ(ews::response_code::error_invalid_id_empty, exc.code());
            EXPECT_STREQ("ErrorInvalidIdEmpty", exc.what());
        }
    }

    TEST_F(MessageTest, CreateAndDeleteMessage)
    {
        const ews::distinguished_folder_id drafts =
            ews::standard_folder::drafts;
        const auto initial_count = service().find_item(drafts).size();

        auto message = ews::message();
        message.set_subject("You are hiding again, aren't you?");
        std::vector<ews::email_address> recipients;
        recipients.push_back(
            ews::email_address("darkwing.duck@duckburg.com")
        );
        message.set_to_recipients(recipients);
        auto item_id = service().create_item(
                message,
                ews::message_disposition::save_only);
        message = service().get_message(item_id);
        recipients = message.get_to_recipients();
        ASSERT_EQ(1U, recipients.size());
        EXPECT_STREQ("darkwing.duck@duckburg.com",
                     recipients.front().value().c_str());
        service().delete_message(std::move(message));
        EXPECT_TRUE(message.get_subject().empty()); // Check sink argument

        auto messages = service().find_item(drafts);
        EXPECT_EQ(initial_count, messages.size());
    }
}

// vim:et ts=4 sw=4 noic cc=80
