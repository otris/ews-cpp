#include "fixture.hpp"

#include <ews/ews.hpp>

namespace tests
{
    TEST(BodyTest, DefaultConstruction)
    {
        auto b = ews::body();
        EXPECT_EQ(ews::body_type::plain_text, b.type());
        EXPECT_FALSE(b.is_truncated());
    }

    TEST(BodyTest, PlainTextToXML)
    {
        auto b = ews::body("Here is some plain text",
                           ews::body_type::plain_text);
        EXPECT_EQ(ews::body_type::plain_text, b.type());
        EXPECT_FALSE(b.is_truncated());
        const char* const expected =
            R"(<Body BodyType="Text">Here is some plain text</Body>)";
        EXPECT_STREQ(expected, b.to_xml().c_str());
    }

    TEST(BodyTest, MakeSureHTMLIsWrappedWithCDATA)
    {
        auto b = ews::body("<b>Here is some HTML</b>", ews::body_type::html);
        EXPECT_EQ(ews::body_type::html, b.type());
        EXPECT_FALSE(b.is_truncated());
        const char* const expected =
R"(<x:Body BodyType="HTML"><![CDATA[<b>Here is some HTML</b>]]></x:Body>)";
        EXPECT_STREQ(expected, b.to_xml("x").c_str());
    }

    TEST(ItemTest, DefaultConstruction)
    {
        auto i = ews::item();
        EXPECT_TRUE(i.get_mime_content().none());
        EXPECT_STREQ("", i.get_subject().c_str());
        EXPECT_FALSE(i.get_item_id().valid());
    }

    TEST(ItemTest, GetAndSetBodyProperty)
    {
        auto item = ews::item();

        // TODO: better: what to do!? EXPECT_FALSE(item.get_body().none());
        // boost::optional desperately missing

        auto original =
            ews::body("<p>Some of the finest Vogon poetry</p>",
                      ews::body_type::html);
        item.set_body(original);

        auto actual = item.get_body();
        EXPECT_EQ(original.type(), actual.type());
        EXPECT_EQ(original.is_truncated(), actual.is_truncated());
        EXPECT_STREQ(original.content().c_str(),
                     actual.content().c_str());
    }

    class ItemTestAgainstExchangeServer : public ServiceFixture {};

    TEST_F(ItemTestAgainstExchangeServer, GetMimeContentProperty)
    {
        // A contact is an item
        auto contact = ews::contact();
        auto s = service();
        const auto item_id = s.create_item(contact);
        ews::internal::on_scope_exit delete_task([&]
        {
            s.delete_contact(std::move(contact));
        });
        contact = s.get_contact(item_id);
        EXPECT_FALSE(contact.get_mime_content().none());
    }
}
