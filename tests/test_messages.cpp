
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
    const ews::distinguished_folder_id drafts = ews::standard_folder::drafts;
    const auto initial_count = service().find_item(drafts).size();

    auto message = ews::message();
    message.set_subject("You are hiding again, aren't you?");
    std::vector<ews::mailbox> recipients;
    recipients.push_back(ews::mailbox("darkwing.duck@duckburg.com"));
    message.set_to_recipients(recipients);
    auto item_id =
        service().create_item(message, ews::message_disposition::save_only);
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

// <IsRead/>
TEST(OfflineMessageTest, IsReadPropertyInitialValue)
{
    auto msg = ews::message();
    EXPECT_FALSE(msg.is_read());
}

TEST(OfflineMessageTest, SetIsReadProperty)
{
    auto msg = ews::message();
    msg.set_read(true);
    EXPECT_TRUE(msg.is_read());
}

TEST_F(MessageTest, UpdateIsReadProperty)
{
    auto& msg = test_message();
    auto prop = ews::property(ews::message_property_path::is_read, true);
    auto new_id = service().update_item(msg.get_item_id(), prop);
    msg = service().get_message(new_id);
    EXPECT_TRUE(msg.is_read());
}

TEST_F(MessageTest, SendItem)
{
    auto& msg = test_message();
    EXPECT_NO_THROW({ service().send_item(msg.get_item_id()); });
}
}

// vim:et ts=4 sw=4
