
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
#include <utility>

namespace tests
{
    TEST_F(ContactTest, GetContactWithInvalidIdThrows)
    {
        auto invalid_id = ews::item_id();
        EXPECT_THROW(service().get_contact(invalid_id), ews::exchange_error);
    }

    TEST_F(ContactTest, GetContactWithInvalidIdExceptionResponse)
    {
        auto invalid_id = ews::item_id();
        try
        {
            service().get_contact(invalid_id);
            FAIL() << "Expected an exception";
        }
        catch (ews::exchange_error& exc)
        {
            EXPECT_EQ(ews::response_code::error_invalid_id_empty, exc.code());
            EXPECT_STREQ("ErrorInvalidIdEmpty", exc.what());
        }
    }

    TEST_F(ContactTest, UpdateEmailAddressProperty)
    {
        auto minnie = test_contact();

        EXPECT_STREQ("", minnie.get_email_address_1().c_str());
        EXPECT_STREQ("", minnie.get_email_address_2().c_str());
        EXPECT_STREQ("", minnie.get_email_address_3().c_str());
        EXPECT_TRUE(minnie.get_email_addresses().empty());

        minnie.set_email_address_1(ews::mailbox("minnie.mouse@duckburg.com"));
        EXPECT_STREQ("minnie.mouse@duckburg.com",
                     minnie.get_email_address_1().c_str());
        EXPECT_STREQ("", minnie.get_email_address_2().c_str());
        EXPECT_STREQ("", minnie.get_email_address_3().c_str());
        auto addresses = minnie.get_email_addresses();
        ASSERT_FALSE(addresses.empty());
        EXPECT_STREQ("minnie.mouse@duckburg.com",
                     addresses.front().value().c_str());

        // TODO: delete an email address entry from a contact
    }

    TEST_F(ContactTest, GetCompleteNameProperty)
    {
        auto minnie = test_contact();

        const auto complete_name = minnie.get_complete_name();

        EXPECT_STREQ("", complete_name.get_title().c_str());
        EXPECT_STREQ("Minnie", complete_name.get_first_name().c_str());
        EXPECT_STREQ("", complete_name.get_middle_name().c_str());
        EXPECT_STREQ("Mouse", complete_name.get_last_name().c_str());
        EXPECT_STREQ("", complete_name.get_suffix().c_str());
        EXPECT_STREQ("", complete_name.get_initials().c_str());
        EXPECT_STREQ("Minnie Mouse", complete_name.get_full_name().c_str());
        EXPECT_STREQ("", complete_name.get_nickname().c_str());
    }
}

// vim:et ts=4 sw=4 noic cc=80
