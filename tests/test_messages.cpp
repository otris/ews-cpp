
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
    ews::internal::on_scope_exit remove_message([&]() {
        try
        {
            service().delete_item(item_id);
        }
        catch (std::exception&)
        { /* ... */
        }
    });

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

// <CcRecipients>
TEST(OfflineMessageTest, CcRecipientsPropertyInitialValue)
{
    auto msg = ews::message();
    EXPECT_TRUE(msg.get_cc_recipients().empty());
}

TEST(OfflineMessageTest, SetCcRecipientsProperty)
{
    auto msg = ews::message();
    auto recipients = std::vector<ews::mailbox>();
    recipients.push_back(ews::mailbox("a@example.com"));
    recipients.push_back(ews::mailbox("b@example.com"));
    msg.set_cc_recipients(recipients);
    auto result = msg.get_cc_recipients();
    EXPECT_EQ(2U, result.size());
}

TEST_F(MessageTest, UpdateCcRecipientsProperty)
{
    auto recipients = std::vector<ews::mailbox>();
    recipients.push_back(ews::mailbox("a@example.com"));
    recipients.push_back(ews::mailbox("b@example.com"));

    auto& msg = test_message();
    auto prop =
        ews::property(ews::message_property_path::cc_recipients, recipients);
    auto new_id = service().update_item(msg.get_item_id(), prop);
    msg = service().get_message(new_id);

    auto result = msg.get_cc_recipients();
    ASSERT_EQ(2U, result.size());
    EXPECT_STREQ("a@example.com", result[0].value().c_str());
    EXPECT_STREQ("b@example.com", result[1].value().c_str());
}

// <BccRecipients>
TEST(OfflineMessageTest, BccRecipientsPropertyInitialValue)
{
    auto msg = ews::message();
    EXPECT_TRUE(msg.get_bcc_recipients().empty());
}

TEST(OfflineMessageTest, SetBccRecipientsProperty)
{
    auto msg = ews::message();
    auto recipients = std::vector<ews::mailbox>();
    recipients.push_back(ews::mailbox("a@example.com"));
    recipients.push_back(ews::mailbox("b@example.com"));
    msg.set_bcc_recipients(recipients);
    auto result = msg.get_bcc_recipients();
    EXPECT_EQ(2U, result.size());
}

TEST_F(MessageTest, UpdateBccRecipientsProperty)
{
    auto recipients = std::vector<ews::mailbox>();
    recipients.push_back(ews::mailbox("peter@example.com"));

    auto& msg = test_message();
    auto prop =
        ews::property(ews::message_property_path::bcc_recipients, recipients);
    auto new_id = service().update_item(msg.get_item_id(), prop);
    msg = service().get_message(new_id);

    auto result = msg.get_bcc_recipients();
    ASSERT_EQ(1U, result.size());
    EXPECT_STREQ("peter@example.com", result[0].value().c_str());
}

// <IsRead>
TEST(OfflineMessageTest, IsReadPropertyInitialValue)
{
    auto msg = ews::message();
    EXPECT_FALSE(msg.is_read());
}

TEST(OfflineMessageTest, SetIsReadProperty)
{
    auto msg = ews::message();
    msg.set_is_read(true);
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

// <From>
TEST(OfflineMessageTest, FromPropertyInitialValue)
{
    auto msg = ews::message();
    EXPECT_TRUE(msg.get_from().none());
}

TEST(OfflineMessageTest, SetFromProperty)
{
    auto msg = ews::message();
    msg.set_from(ews::mailbox("chuck@example.com"));
    ASSERT_FALSE(msg.get_from().none());
    EXPECT_STREQ("chuck@example.com", msg.get_from().value().c_str());
}

TEST_F(MessageTest, UpdateFromProperty)
{
    auto& msg = test_message();
    auto prop = ews::property(ews::message_property_path::from,
                              ews::mailbox("chuck@example.com"));
    auto new_id = service().update_item(msg.get_item_id(), prop);
    msg = service().get_message(new_id);
    EXPECT_STREQ("chuck@example.com", msg.get_from().value().c_str());
}

// <InternetMessageId>
TEST(OfflineMessageTest, InternetMessageIdPropertyInitialValue)
{
    auto msg = ews::message();
    EXPECT_TRUE(msg.get_internet_message_id().empty());
}

TEST(OfflineMessageTest, SetInternetMessageIdProperty)
{
    auto msg = ews::message();
    msg.set_internet_message_id("xyz");
    EXPECT_STREQ("xyz", msg.get_internet_message_id().c_str());
}

TEST_F(MessageTest, CreateMessageWithInternetMessageId)
{
    auto msg = ews::message();
    msg.set_internet_message_id("xxxxxxxx-xxxx-mxxx-nxxx-xxxxxxxxxxxx");
    auto id = service().create_item(msg, ews::message_disposition::save_only);
    ews::internal::on_scope_exit remove_message(
        [&]() { service().delete_item(id); });
    msg = service().get_message(id, ews::base_shape::all_properties);
    EXPECT_STREQ("xxxxxxxx-xxxx-mxxx-nxxx-xxxxxxxxxxxx",
                 msg.get_internet_message_id().c_str());
}
} // namespace tests
