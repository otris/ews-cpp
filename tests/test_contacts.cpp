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

        minnie.set_email_address_1(
                ews::mailbox("minnie.mouse@duckburg.com")
            );
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
}

// vim:et ts=4 sw=4 noic cc=80
