#include "fixture.hpp"
#include <utility>

namespace tests
{
    class ServiceTest : public ServiceFixture
    {
    public:
        void SetUp()
        {
            ServiceFixture::SetUp();

            contact_.set_given_name("Minnie");
            contact_.set_surname("Mouse");
            const auto item_id = service().create_item(contact_);
            contact_ = service().get_contact(item_id);
        }

        void TearDown()
        {
            service().delete_contact(std::move(contact_));

            ServiceFixture::TearDown();
        }

        ews::contact& test_contact() { return contact_; }

    private:
        ews::contact contact_;
    };

    TEST_F(ServiceTest, UpdateItemWithSetItemField)
    {
        // <SetItemField> is used whenever we want to replace an existing value.
        // If none exists yet, it is created

        auto minnie = test_contact();

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

    TEST_F(ServiceTest, UpdateItemWithDeleteItemField)
    {
        // <DeleteItemField> is simply just a FieldURI to the property that we
        // want to delete. Not much to it. It is automatically deduced by
        // service::update_item if the user has not provided any value to the
        // property he wanted to update

        auto minnie = test_contact();
        auto contact_property = ews::contact_property_path();
        auto prop = ews::property(contact_property.given_name);
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("", minnie.get_given_name().c_str());
    }

#if 0
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

        auto item_property = ews::item_property_path();
        auto minnie = test_contact();

        // TODO: property c'tor needs overload for other property types, too!
        //
        // - find a way to serialize those properties to a simple string
        //   including attributes and all
        // - then maybe get rid of property_path::property_name

        // auto body = ews::body("Some text", ews::body_type::plain_text);
        // auto prop = ews::property(item_property.body, body);
        // auto new_id = service().update_item(minnie.get_item_id(), prop);
    }
#endif
}
