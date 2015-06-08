#include "fixture.hpp"
#include <utility>

namespace tests
{
    class ContactTest : public ServiceFixture {};

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

    // TODO
#if 0
    TEST_F(ContactTest, CreateAndDelete)
    {
        auto contact = ews::contact();
        contact.set_given_name("General");
        contact.set_surname("Snozzie");
        const auto item_id = service().create_item(contact);
        auto snozzie = service().get_contact(item_id);
        ews::internal::on_scope_exit delete_item([&]
        {
            service().delete_contact(std::move(snozzie));
        });
    }
#endif

    TEST_F(ContactTest, GetEmailAddress)
    {
        auto contact = ews::contact();
        contact.set_given_name("General");
        contact.set_surname("Snozzie");
        const auto item_id = service().create_item(contact);
        auto general_snozzie = service().get_contact(item_id);
        ews::internal::on_scope_exit delete_item([&]
        {
            service().delete_contact(std::move(general_snozzie));
        });

        EXPECT_STREQ("", general_snozzie.get_email_address_1().c_str());
        EXPECT_STREQ("", general_snozzie.get_email_address_2().c_str());
        EXPECT_STREQ("", general_snozzie.get_email_address_3().c_str());
        EXPECT_TRUE(general_snozzie.get_email_addresses().empty());

        general_snozzie.set_email_address_1(
                ews::email_address("snozzie@duckburg.com")
            );
        EXPECT_STREQ("snozzie@duckburg.com",
                     general_snozzie.get_email_address_1().c_str());
        EXPECT_STREQ("", general_snozzie.get_email_address_2().c_str());
        EXPECT_STREQ("", general_snozzie.get_email_address_3().c_str());
        auto addresses = general_snozzie.get_email_addresses();
        ASSERT_FALSE(addresses.empty());
        EXPECT_STREQ("snozzie@duckburg.com", addresses.front().value().c_str());

        // TODO: delete an email address entry from a contact
    }
}
