
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

#include <fstream>
#include <ios>
#include <iostream>
#include <iterator>
#include <stdexcept>
#include <utility>

using ews::internal::on_scope_exit;

namespace tests
{

#ifdef EWS_HAS_FILESYSTEM_HEADER
class GetItemRequestTest : public FakeServiceFixture
{
private:
    void SetUp()
    {
        FakeServiceFixture::SetUp();

        const auto assets = std::filesystem::path(
            ews::test::global_data::instance().assets_dir);
        const auto file_path = assets / "get_item_response_message.xml";
        std::ifstream ifstr(file_path.string(),
                            std::ifstream::in | std::ios::binary);
        if (!ifstr.is_open())
        {
            throw std::runtime_error("Could not open file for reading: " +
                                     file_path.string());
        }
        ifstr.unsetf(std::ios::skipws);
        ifstr.seekg(0, std::ios::end);
        const auto file_size = ifstr.tellg();
        ifstr.seekg(0, std::ios::beg);

        std::vector<char> buf;
        buf.reserve(file_size);
        buf.insert(begin(buf), std::istream_iterator<char>(ifstr),
                   std::istream_iterator<char>());
        ifstr.close();
        buf.push_back('\0');

        set_next_fake_response(std::move(buf));
    }
};

TEST_F(GetItemRequestTest, WithAdditionalProperties)
{
    auto& serv = service();
    auto fake_id = ews::item_id();

    auto additional_props = std::vector<ews::property_path>();
    additional_props.push_back(ews::item_property_path::body);
    ews::item_shape shape(std::move(additional_props));
    auto cal_item = serv.get_calendar_item(fake_id, shape);
    (void)cal_item;

    EXPECT_NE(get_last_request().request_string().find(
                  "<t:AdditionalProperties>"
                  "<t:FieldURI FieldURI=\"item:Body\"/>"
                  "</t:AdditionalProperties>"),
              std::string::npos);
}
#endif // EWS_HAS_FILESYSTEM_HEADER

class SoapHeader : public FakeServiceFixture
{
private:
    void SetUp()
    {
        FakeServiceFixture::SetUp();
        set_next_fake_response(
            "<s:Envelope "
            "xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\">\n"
            "    <s:Header>\n"
            "        <h:ServerVersionInfo MajorVersion=\"15\" "
            "MinorVersion=\"0\" MajorBuildNumber=\"847\" "
            "MinorBuildNumber=\"31\" Version=\"V2_8\" "
            "xmlns:h=\"http://schemas.microsoft.com/exchange/services/2006/"
            "types\" "
            "xmlns=\"http://schemas.microsoft.com/exchange/services/2006/"
            "types\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" "
            "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>\n"
            "    </s:Header>\n"
            "    <s:Body "
            "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" "
            "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\n"
            "        <m:CreateItemResponse "
            "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/"
            "messages\" "
            "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/"
            "types\">\n"
            "            <m:ResponseMessages>\n"
            "                <m:CreateItemResponseMessage "
            "ResponseClass=\"Success\">\n"
            "                    <m:ResponseCode>NoError</m:ResponseCode>\n"
            "                    <m:Items>\n"
            "                        <t:Message>\n"
            "                            <t:ItemId "
            "Id="
            "\"AAMkAGRhYmQ5Njg0LTNhMjEtNDZkOS1hN2QyLTUzZTI3MjdhN2ZkYgBGAAAA"
            "AAC5LuzvcattTqJiAgNAfv18BwDKOL2xzF+"
            "1SL9YsnXMX2cZAAAAAAEQAADKOL2xzF+1SL9YsnXMX2cZAAADQIPPAAA=\" "
            "ChangeKey=\"CQAAABYAAADKOL2xzF+1SL9YsnXMX2cZAAADQKty\"/>\n"
            "                        </t:Message>\n"
            "                    </m:Items>\n"
            "                </m:CreateItemResponseMessage>\n"
            "            </m:ResponseMessages>\n"
            "        </m:CreateItemResponse>\n"
            "    </s:Body>\n"
            "</s:Envelope>\n");
    }
};

TEST_F(SoapHeader, SupportsExchange2007)
{
    auto& serv = service();
    serv.set_request_server_version(ews::server_version::exchange_2007);
    auto task = ews::task();
    task.set_subject("Random To-Do item");
    serv.create_item(task);
    auto request = get_last_request();
    EXPECT_TRUE(request.header_contains(
        "<t:RequestServerVersion Version=\"Exchange2007\"/>"));
}

TEST_F(SoapHeader, SupportsExchange2007_SP1)
{
    auto& serv = service();
    serv.set_request_server_version(ews::server_version::exchange_2007_sp1);
    auto task = ews::task();
    task.set_subject("Another To-Do item");
    serv.create_item(task);
    auto request = get_last_request();
    EXPECT_TRUE(request.header_contains(
        "<t:RequestServerVersion Version=\"Exchange2007_SP1\"/>"));
}

TEST_F(SoapHeader, SupportsExchange2010)
{
    auto& serv = service();
    serv.set_request_server_version(ews::server_version::exchange_2010);
    auto task = ews::task();
    task.set_subject("Yet another To-Do item");
    serv.create_item(task);
    auto request = get_last_request();
    EXPECT_TRUE(request.header_contains(
        "<t:RequestServerVersion Version=\"Exchange2010\"/>"));
}

TEST_F(SoapHeader, SupportsExchange2010_SP1)
{
    auto& serv = service();
    serv.set_request_server_version(ews::server_version::exchange_2010_sp1);
    auto task = ews::task();
    task.set_subject("Feed the cat");
    serv.create_item(task);
    auto request = get_last_request();
    EXPECT_TRUE(request.header_contains(
        "<t:RequestServerVersion Version=\"Exchange2010_SP1\"/>"));
}

TEST_F(SoapHeader, SupportsExchange2010_SP2)
{
    auto& serv = service();
    serv.set_request_server_version(ews::server_version::exchange_2010_sp2);
    auto task = ews::task();
    task.set_subject("Do something about code duplication");
    serv.create_item(task);
    auto request = get_last_request();
    EXPECT_TRUE(request.header_contains(
        "<t:RequestServerVersion Version=\"Exchange2010_SP2\"/>"));
}

TEST_F(SoapHeader, SupportsExchange2013)
{
    auto& serv = service();
    serv.set_request_server_version(ews::server_version::exchange_2013);
    auto task = ews::task();
    task.set_subject("Buy new shoes");
    serv.create_item(task);
    auto request = get_last_request();
    EXPECT_TRUE(request.header_contains(
        "<t:RequestServerVersion Version=\"Exchange2013\"/>"));
}

TEST_F(SoapHeader, SupportsExchange2013_SP1)
{
    auto& serv = service();
    serv.set_request_server_version(ews::server_version::exchange_2013_sp1);
    auto task = ews::task();
    task.set_subject("Get some milk from the store");
    serv.create_item(task);
    auto request = get_last_request();
    EXPECT_TRUE(request.header_contains(
        "<t:RequestServerVersion Version=\"Exchange2013_SP1\"/>"));
}

TEST_F(SoapHeader, DefaultServerVersionIs2013_SP1)
{
    auto serv = service();
    EXPECT_EQ(serv.get_request_server_version(),
              ews::server_version::exchange_2013_sp1);
}

TEST_F(SoapHeader, SpecifyTimeZone)
{
    auto& serv = service();
    serv.set_time_zone(ews::time_zone::w_europe_standard_time);
    auto task = ews::task();
    task.set_subject("Get some milk from the store");
    serv.create_item(task);
    auto request = get_last_request();
    EXPECT_TRUE(request.header_contains(
        "<t:TimeZoneContext>"
        "<t:TimeZoneDefinition Id=\"W. Europe Standard Time\"/>"
        "</t:TimeZoneContext>"));
    EXPECT_EQ(serv.get_time_zone(), ews::time_zone::w_europe_standard_time);
}

TEST_F(SoapHeader, ImpersonateAsAnotherUser)
{
    const auto someone = ews::connecting_sid(
        ews::connecting_sid::type::smtp_address, "batman@gothamcity.com");
    auto& serv = service();

    auto task = ews::task();
    task.set_subject("Random To-Do item");

    serv.impersonate(someone).create_item(task);
    auto request = get_last_request();
    EXPECT_TRUE(request.header_contains(
        "<t:ExchangeImpersonation><t:ConnectingSID><t:SmtpAddress>batman@"
        "gothamcity.com</t:SmtpAddress></t:ConnectingSID></"
        "t:ExchangeImpersonation>"));

    serv.impersonate();
    serv.create_item(task);
    request = get_last_request();
    EXPECT_FALSE(request.header_contains(
        "<t:ExchangeImpersonation><t:ConnectingSID><t:SmtpAddress>batman@"
        "gothamcity.com</t:SmtpAddress></t:ConnectingSID></"
        "t:ExchangeImpersonation>"));
}

class ServiceTest : public ContactTest
{
};

TEST_F(ServiceTest, UpdateItemOfReadOnlyPropertyThrows)
{
    auto minnie = test_contact();
    ASSERT_FALSE(minnie.has_attachments());
    auto prop = ews::property(ews::item_property_path::has_attachments, true);
    EXPECT_THROW({ service().update_item(minnie.get_item_id(), prop); },
                 ews::exchange_error);
}

TEST_F(ServiceTest, UpdateItemOfReadOnlyPropertyThrowsWhatMessage)
{
    auto minnie = test_contact();
    ASSERT_FALSE(minnie.has_attachments());
    auto prop = ews::property(ews::item_property_path::has_attachments, true);
    try
    {
        service().update_item(minnie.get_item_id(), prop);
        FAIL() << "Expected exception to be thrown";
    }
    catch (ews::exchange_error& exc)
    {
        EXPECT_EQ(ews::response_code::error_invalid_property_set, exc.code());
    }
}

TEST_F(ServiceTest, UpdateItemWithSetItemField)
{
    // <SetItemField> is used whenever we want to replace an existing value.
    // If none exists yet, it is created

    auto minnie = test_contact();

    EXPECT_STREQ("Mickey", minnie.get_spouse_name().c_str());
    auto spouse_name_property =
        ews::property(ews::contact_property_path::spouse_name, "Mickey");
    auto new_id =
        service().update_item(minnie.get_item_id(), spouse_name_property,
                              ews::conflict_resolution::auto_resolve);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Mickey", minnie.get_spouse_name().c_str());

    spouse_name_property =
        ews::property(ews::contact_property_path::spouse_name, "Peg-Leg Pedro");
    new_id = service().update_item(minnie.get_item_id(), spouse_name_property,
                                   ews::conflict_resolution::auto_resolve);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Peg-Leg Pedro", minnie.get_spouse_name().c_str());
}

TEST_F(ServiceTest, UpdateItemWithDeleteItemField)
{
    // <DeleteItemField> is simply just a FieldURI to the property that we
    // want to delete. Not much to it. It is automatically deduced by
    // service::update_item if the user has not provided any value to the
    // property he wanted to update

    auto minnie = test_contact();
    ASSERT_FALSE(minnie.get_given_name().empty());
    auto prop = ews::property(ews::contact_property_path::given_name, "");
    auto update = ews::update(prop, ews::update::operation::delete_item_field);
    auto new_id = service().update_item(minnie.get_item_id(), update);
    minnie = service().get_contact(new_id);
    EXPECT_TRUE(minnie.get_given_name().empty());
}

TEST_F(ServiceTest, UpdateItemWithAppendToItemField)
{
    // <AppendToItemField> is automatically deduced by service::update_item
    // so user does not need to bother. It is only valid for following
    // properties (at least in EWS 2007 slang):
    //
    // - calendar:OptionalAttendees
    // - calendar:RequiredAttendees
    // - calendar:Resources
    // - item:Body
    // - message:ToRecipients
    // - message:CcRecipients
    // - message:BccRecipients
    // - message:ReplyTo

    auto message = ews::message();
    message.set_subject("You are hiding again, aren't you?");
    std::vector<ews::mailbox> recipients;
    recipients.push_back(ews::mailbox("darkwing.duck@duckburg.com"));
    message.set_to_recipients(recipients);
    auto item_id =
        service().create_item(message, ews::message_disposition::save_only);
    on_scope_exit delete_message(
        [&] { service().delete_message(std::move(message)); });
    message = service().get_message(item_id);
    recipients = message.get_to_recipients();
    EXPECT_EQ(1U, recipients.size());

    std::vector<ews::mailbox> additional_recipients;
    additional_recipients.push_back(ews::mailbox("gus.goose@duckburg.com"));
    auto prop = ews::property(ews::message_property_path::to_recipients,
                              additional_recipients);
    auto change =
        ews::update(prop, ews::update::operation::append_to_item_field);
    auto new_id = service().update_item(message.get_item_id(), change);
    message = service().get_message(item_id);
    ASSERT_EQ(2U, message.get_to_recipients().size());
}

TEST_F(ServiceTest, SendItem)
{
    auto msg = ews::message();
    msg.set_subject("You are hiding again, aren't you?");
    std::vector<ews::mailbox> recipients;
    recipients.push_back(ews::mailbox("darkwing.duck@duckburg.com"));
    msg.set_to_recipients(recipients);
    const auto item_id =
        service().create_item(msg, ews::message_disposition::save_only);
    msg = service().get_message(item_id);
    EXPECT_NO_THROW({ service().send_item(msg.get_item_id()); });
}
} // namespace tests
