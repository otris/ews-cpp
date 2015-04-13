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
            EXPECT_STREQ("Request failed", exc.what());
        }
    }

    TEST_F(ContactTest, CreateAndDeleteContact)
    {
        auto contact = ews::contact();
        contact.set_given_name("Donald");
        contact.set_surname("Duck");
        const auto item_id = service().create_item(contact);

        auto created_contact = service().get_contact(item_id);
        EXPECT_STREQ("Donald", created_contact.get_given_name().c_str());
        EXPECT_STREQ("Duck",   created_contact.get_surname().c_str());

        ASSERT_NO_THROW(
        {
            service().delete_contact(std::move(created_contact));
        });
        EXPECT_STREQ("", created_contact.get_surname().c_str());

        // Check if it is still in store
        EXPECT_THROW(service().get_contact(item_id), ews::exchange_error);
    }
}
