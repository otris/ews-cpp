#include <ews/ews.hpp>
#include <gtest/gtest.h>
#include <vector>
#include <string>
#include <sstream>
#include <algorithm>
#include <iterator>
#include <cstring>

namespace
{
    inline std::vector<std::string> split(const std::string& str, char delim)
    {
        std::stringstream sstr(str);
        std::string item;
        std::vector<std::string> elems;
        while (std::getline(sstr, item, delim))
        {
            elems.push_back(item);
        }
        return elems;
    }

    const std::string contact_card =
"<s:Envelope xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\">\n"
"    <s:Body xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\n"
"        <m:GetItemResponse xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n"
"            <m:ResponseMessages>\n"
"                <m:GetItemResponseMessage ResponseClass=\"Success\">\n"
"                    <m:ResponseCode>NoError</m:ResponseCode>\n"
"                    <m:Items>\n"
"                        <t:Contact>\n"
"                            <t:DateTimeSent>2015-05-21T10:13:28Z</t:DateTimeSent>\n"
"                            <t:DateTimeCreated>2015-05-21T10:13:28Z</t:DateTimeCreated>\n"
"                            <t:EffectiveRights>\n"
"                                <t:Delete>true</t:Delete>\n"
"                                <t:Modify>true</t:Modify>\n"
"                                <t:Read>true</t:Read>\n"
"                            </t:EffectiveRights>\n"
"                            <t:Culture>en-US</t:Culture>\n"
"                        </t:Contact>\n"
"                    </m:Items>\n"
"                </m:GetItemResponseMessage>\n"
"            </m:ResponseMessages>\n"
"        </m:GetItemResponse>\n"
"    </s:Body>\n"
"</s:Envelope>";

    const std::string malformed_xml_1 =
"<html>\n"
"  <head>\n"
"  </head>\n"
"  <body>\n"
"    <p>\n"
"      <h1</h1>\n"
// line 6   ~
"    </p>\n"
"  </body>\n"
"</html>";

    const std::string malformed_xml_2 =
"<s:Envelope xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\">\n"
"    <s:Body xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\n"
"        <m:GetItemResponse xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n"
"            <m:ResponseMessages>\n"
"                <m:GetItemResponseMessage ResponseClass=\"Success\">\n"
"                    <m:ResponseCode>NoError</m:ResponseCode>\n"
"                    <m:Items>\n"
"                        <t:Contact>\n"
"                            <t:Cultu</t:Culture>\n"
// line 9                              ~
"                        </t:Contact>\n"
"                    </m:Items>\n"
"                </m:GetItemResponseMessage>\n"
"            </m:ResponseMessages>\n"
"        </m:GetItemResponse>\n"
"    </s:Body>\n"
"</s:Envelope>";

}

namespace tests
{
    TEST(InternalTest, NamespaceURIs)
    {
        typedef ews::internal::uri<> uri;
        EXPECT_EQ(uri::microsoft::errors_size,
                  std::strlen(uri::microsoft::errors()));
        EXPECT_EQ(uri::microsoft::types_size,
                  std::strlen(uri::microsoft::types()));
        EXPECT_EQ(uri::microsoft::messages_size,
                  std::strlen(uri::microsoft::messages()));

        EXPECT_EQ(uri::soapxml::envelope_size,
                  std::strlen(uri::soapxml::envelope()));
    }

    TEST(InternalTest, MimeContentDefaultConstruction)
    {
        typedef ews::mime_content mime_content;

        auto m = mime_content();
        EXPECT_TRUE(m.none());
        EXPECT_EQ(0U, m.len_bytes());
        EXPECT_TRUE(m.character_set().empty());
        EXPECT_EQ(nullptr, m.bytes());
    }

    TEST(InternalTest, MimeContentConstructionWithData)
    {
        using ews::mime_content;
        const char* content = "SGVsbG8sIHdvcmxkPw==";

        auto m = mime_content();

        {
            const auto b = std::string(content);
            m = mime_content("UTF-8", b.data(), b.length());
            EXPECT_FALSE(m.none());
            EXPECT_EQ(20U, m.len_bytes());
            EXPECT_FALSE(m.character_set().empty());
            EXPECT_STREQ("UTF-8", m.character_set().c_str());
            EXPECT_NE(nullptr, m.bytes());
        }

        // b was destructed

#if EWS_HAS_ROBUST_NONMODIFYING_SEQ_OPS
        EXPECT_TRUE(std::equal(m.bytes(), m.bytes() + m.len_bytes(),
                               content, content + std::strlen(content)));
#else
        EXPECT_TRUE(m.len_bytes() == std::strlen(content) &&
                    std::equal(m.bytes(), m.bytes() + m.len_bytes(),
                               content));
#endif
    }

    TEST(InternalTest, ExtractSubTreeFromDocument)
    {
        using namespace ews::internal;

        rapidxml::xml_document<> doc;
        auto str = doc.allocate_string(contact_card.c_str());
        doc.parse<0>(str);
        auto effective_rights_element =
            get_element_by_qname(doc,
                                 "EffectiveRights",
                                 uri<>::microsoft::types());
        auto effective_rights = xml_subtree(*effective_rights_element);
        EXPECT_STREQ(
            "<t:EffectiveRights><t:Delete>true</t:Delete><t:Modify>true</t:Modify><t:Read>true</t:Read></t:EffectiveRights>",
            effective_rights.to_string().c_str());
    }

    TEST(InternalTest, GetNonExistingElementReturnsNullptr)
    {
        using namespace ews::internal;

        rapidxml::xml_document<> doc;
        auto str = doc.allocate_string(
            "<html lang=\"en\"><head><meta charset=\"utf-8\"/><title>Welcome</title></head><body><h1>Greetings</h1><p>Hello!</p></body></html>");
        doc.parse<0>(str);
        auto body_element = get_element_by_qname(doc, "body", "");
        auto body = xml_subtree(*body_element);
        rapidxml::xml_node<char>* node = nullptr;
        ASSERT_NO_THROW(
        {
            node = body.get_node("h1");
        });
        EXPECT_EQ(nullptr, node);
    }

    TEST(InternalTest, GetElementFromSubTree)
    {
        using namespace ews::internal;

        rapidxml::xml_document<> doc;
        auto str = doc.allocate_string(contact_card.c_str());
        doc.parse<0>(str);
        auto effective_rights_element =
            get_element_by_qname(doc,
                                 "EffectiveRights",
                                 uri<>::microsoft::types());
        auto effective_rights = xml_subtree(*effective_rights_element);

        EXPECT_STREQ("true", effective_rights.get_node("Delete")->value());
    }

    TEST(InternalTest, UpdateSubTreeElement)
    {
        using namespace ews::internal;

        rapidxml::xml_document<> doc;
        auto str = doc.allocate_string(contact_card.c_str());
        doc.parse<0>(str);
        auto effective_rights_element =
            get_element_by_qname(doc,
                                 "EffectiveRights",
                                 uri<>::microsoft::types());
        auto effective_rights = xml_subtree(*effective_rights_element);

        effective_rights.set_or_update("Modify", "false");
        EXPECT_STREQ("false", effective_rights.get_node("Modify")->value());
        EXPECT_STREQ(
            "<t:EffectiveRights><t:Delete>true</t:Delete><t:Modify>false</t:Modify><t:Read>true</t:Read></t:EffectiveRights>",
            effective_rights.to_string().c_str());
    }

    // TODO: test size_hint parameter of xml_subtree reduces/eliminates reallocs

    TEST(InternalTest, XMLParseErrorMessageShort)
    {
        std::vector<char> data;
        std::copy(begin(malformed_xml_1), end(malformed_xml_1),
                  std::back_inserter(data));
        data.emplace_back('\0');
        rapidxml::xml_document<> doc;
        try
        {
            doc.parse<0>(&data[0]);
            FAIL() << "Expected rapidxml::parse_error to be raised";
        }
        catch (rapidxml::parse_error& exc)
        {
            const auto str = ews::xml_parse_error::error_message_from(exc,
                                                                      data);
            ASSERT_FALSE(str.empty());
            const auto output = split(str, '\n');
            EXPECT_EQ(4U, output.size());
            EXPECT_STREQ("in line 6:", output[0].c_str());
            EXPECT_STREQ("   </head>   <body>     <p>       <h1</h1>     </p>   </body> </html>",
                         output[2].c_str());
            EXPECT_STREQ("                                       ~",
                         output[3].c_str());
        }
    }

    TEST(InternalTest, XMLParseErrorMessageLong)
    {
        std::vector<char> data;
        std::copy(begin(malformed_xml_2), end(malformed_xml_2),
                  std::back_inserter(data));
        data.emplace_back('\0');
        rapidxml::xml_document<> doc;
        try
        {
            doc.parse<0>(&data[0]);
            FAIL() << "Expected rapidxml::parse_error to be raised";
        }
        catch (rapidxml::parse_error& exc)
        {
            const auto str = ews::xml_parse_error::error_message_from(exc,
                                                                      data);
            ASSERT_FALSE(str.empty());
            const auto output = split(str, '\n');
            EXPECT_EQ(4U, output.size());
            EXPECT_STREQ("in line 9:", output[0].c_str());
            EXPECT_STREQ("                             <t:Cultu</t:Culture>                         </t:",
                         output[2].c_str());
            EXPECT_STREQ("                                       ~",
                         output[3].c_str());
        }
    }
}
