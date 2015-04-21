#include "fixture.hpp"
#include <ews/ews.hpp>

using ews::item;
using ews::internal::on_scope_exit;

namespace tests
{
    class ItemTest : public ServiceFixture {};

    TEST_F(ItemTest, DefaultConstruction)
    {
        auto i = item();
        EXPECT_STREQ("", i.get_mime_content().c_str());
        EXPECT_STREQ("", i.get_subject().c_str());
        EXPECT_FALSE(i.get_item_id().valid());
    }

    TEST_F(ItemTest, Getter)
    {
        // A contact is an item
        auto contact = ews::contact();
        auto s = service();
        const auto item_id = s.create_item(contact);
        on_scope_exit delete_task([&]
        {
            s.delete_contact(std::move(contact));
        });
        contact = s.get_contact(item_id);
        EXPECT_STRNE("", contact.get_mime_content().c_str());
    }
}
