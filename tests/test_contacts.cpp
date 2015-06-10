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

    TEST_F(ContactTest, UpdateItemWithSetItemField)
    {
        // SetItemField is used whenever we want to replace an existing value.
        // If none exists yet, it is created

        auto contact = ews::contact();
        contact.set_given_name("Minnie");
        contact.set_surname("Mouse");
        const auto item_id = service().create_item(contact);
        auto minnie = service().get_contact(item_id);
        ews::internal::on_scope_exit delete_item([&]
        {
            service().delete_contact(std::move(minnie));
        });

        EXPECT_STREQ("", minnie.get_spouse_name().c_str());
        auto contact_property = ews::contact_property_path();
        auto spouse_name_property = ews::property(contact_property.spouse_name,
                                                  "Mickey");
        auto new_id =
            service().update_item(minnie.get_item_id(),
                                  spouse_name_property,
                                  ews::conflict_resolution::auto_resolve);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("Mickey", minnie.get_spouse_name().c_str());

        spouse_name_property = ews::property(contact_property.spouse_name,
                                             "Peg-Leg Pedro");
        new_id =
            service().update_item(minnie.get_item_id(),
                                  spouse_name_property,
                                  ews::conflict_resolution::auto_resolve);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("Peg-Leg Pedro", minnie.get_spouse_name().c_str());
    }

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
