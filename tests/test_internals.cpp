
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

#include <algorithm>
#include <ews/ews.hpp>
#include <ews/rapidxml/rapidxml_print.hpp>
#include <gtest/gtest.h>
#include <iterator>
#include <sstream>
#include <string.h>
#include <string>
#include <vector>

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
    "    <s:Body xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" "
    "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\n"
    "        <m:GetItemResponse "
    "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/"
    "messages\" "
    "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/"
    "types\">\n"
    "            <m:ResponseMessages>\n"
    "                <m:GetItemResponseMessage ResponseClass=\"Success\">\n"
    "                    <m:ResponseCode>NoError</m:ResponseCode>\n"
    "                    <m:Items>\n"
    "                        <t:Contact>\n"
    "                            "
    "<t:DateTimeSent>2015-05-21T10:13:28Z</t:DateTimeSent>\n"
    "                            "
    "<t:DateTimeCreated>2015-05-21T10:13:28Z</t:DateTimeCreated>\n"
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

const std::string appointment =
    "<s:Envelope xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\">\n"
    "    <s:Body xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" "
    "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\n"
    "        <m:GetItemResponse "
    "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/"
    "messages\" "
    "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/"
    "types\">\n"
    "            <m:ResponseMessages>\n"
    "                <m:GetItemResponseMessage ResponseClass=\"Success\">\n"
    "                    <m:ResponseCode>NoError</m:ResponseCode>\n"
    "                    <m:Items>\n"
    "                        <t:CalendarItem>\n"
    "                            <t:ItemId Id=\"xyz\" ChangeKey=\"zyx\" />\n"
    "                            <t:Subject>Customer X Installation "
    "Prod.-System</t:Subject>\n"
    "                            <t:Start>2017-09-07T15:00:00Z</t:Start>\n"
    "                            <t:End>2017-09-07T16:30:00Z</t:End>\n"
    "                        </t:CalendarItem>\n"
    "                    </m:Items>\n"
    "                </m:GetItemResponseMessage>\n"
    "            </m:ResponseMessages>\n"
    "        </m:GetItemResponse>\n"
    "    </s:Body>\n"
    "</s:Envelope>";

const std::string malformed_xml_1 = "<html>\n"
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
    "    <s:Body xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" "
    "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\n"
    "        <m:GetItemResponse "
    "xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/"
    "messages\" "
    "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/"
    "types\">\n"
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
} // namespace

namespace tests
{
TEST(InternalTest, NamespaceURIs)
{
    typedef ews::internal::uri<> uri;
    EXPECT_EQ(uri::microsoft::errors_size, strlen(uri::microsoft::errors()));
    EXPECT_EQ(uri::microsoft::types_size, strlen(uri::microsoft::types()));
    EXPECT_EQ(uri::microsoft::messages_size,
              strlen(uri::microsoft::messages()));

    EXPECT_EQ(uri::soapxml::envelope_size, strlen(uri::soapxml::envelope()));
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
    EXPECT_TRUE(std::equal(m.bytes(), m.bytes() + m.len_bytes(), content,
                           content + strlen(content)));
#else
    EXPECT_TRUE(m.len_bytes() == strlen(content) &&
                std::equal(m.bytes(), m.bytes() + m.len_bytes(), content));
#endif
}

TEST(InternalTest, TraversalIsDepthFirst)
{
    const auto xml = std::string("<a>\n"
                                 "   <b>\n"
                                 "       <c>\n"
                                 "           <d>Some value</d>\n"
                                 "       </c>\n"
                                 "   </b>\n"
                                 "   <e>\n"
                                 "       <f>More</f>\n"
                                 "       <g>Even more</g>\n"
                                 "   </e>\n"
                                 "   <h>Done</h>\n"
                                 "</a>");

#ifdef EWS_HAS_INITIALIZER_LISTS
    const std::vector<std::string> expected_order{"d", "c", "b", "f",
                                                  "g", "e", "h", "a"};
#else
    std::vector<std::string> expected_order;
    expected_order.push_back("d");
    expected_order.push_back("c");
    expected_order.push_back("b");
    expected_order.push_back("f");
    expected_order.push_back("g");
    expected_order.push_back("e");
    expected_order.push_back("h");
    expected_order.push_back("a");
#endif

    std::vector<std::string> actual_order;

    rapidxml::xml_document<> doc;
    auto str = doc.allocate_string(xml.c_str());
    doc.parse<0>(str);

    ews::internal::traverse_elements(
        doc, [&](const rapidxml::xml_node<>& node) -> bool {
            actual_order.push_back(node.name());
            return false;
        });

#if EWS_HAS_ROBUST_NONMODIFYING_SEQ_OPS
    auto ordered = std::equal(begin(actual_order), end(actual_order),
                              begin(expected_order), end(expected_order));
#else
    auto ordered = actual_order.size() == expected_order.size() &&
                   std::equal(begin(actual_order), end(actual_order),
                              begin(expected_order));
#endif

    std::stringstream sstr;
    sstr << "Expected: ";
    for (const auto& e : expected_order)
    {
        sstr << e << ", ";
    }
    sstr << "\nActual: ";
    for (const auto& e : actual_order)
    {
        sstr << e << ", ";
    }
    EXPECT_TRUE(ordered) << sstr.str();
}

TEST(InternalTest, GetElementByQualifiedName)
{
    const char* noxmlns = "";
    const auto xml = std::string("<a>\n"
                                 "   <b>\n"
                                 "       <c>\n"
                                 "           <d>Some value</d>\n"
                                 "       </c>\n"
                                 "   </b>\n"
                                 "   <e>\n"
                                 "       <f>More</f>\n"
                                 "       <g>Even more</g>\n"
                                 "   </e>\n"
                                 "   <h>Done</h>\n"
                                 "</a>");

    rapidxml::xml_document<> doc;
    auto str = doc.allocate_string(xml.c_str());
    doc.parse<0>(str);

    auto root = doc.first_node();
    auto node = ews::internal::get_element_by_qname(*root, "g", noxmlns);
    ASSERT_TRUE(node);
    EXPECT_STREQ("Even more", node->value());
}

TEST(InternalTest, GetElementByQualifiedNameMultiplePaths)
{
    const char* noxmlns = "";
    const auto xml = std::string("<a>\n"
                                 "   <b>\n"
                                 "       <c>\n"
                                 "           <d>Some value</d>\n"
                                 "       </c>\n"
                                 "   </b>\n"
                                 "   <d>Some other value</d>\n"
                                 "</a>");

    rapidxml::xml_document<> doc;
    auto str = doc.allocate_string(xml.c_str());
    doc.parse<0>(str);

    auto root = doc.first_node();
    auto node = ews::internal::get_element_by_qname(*root, "d", noxmlns);
    ASSERT_TRUE(node);
    EXPECT_STREQ("Some value", node->value());
}

TEST(InternalTest, ExtractSubTreeFromDocument)
{
    using namespace ews::internal;

    rapidxml::xml_document<> doc;
    auto str = doc.allocate_string(contact_card.c_str());
    doc.parse<0>(str);
    auto effective_rights_element =
        get_element_by_qname(doc, "EffectiveRights", uri<>::microsoft::types());
    auto effective_rights = xml_subtree(*effective_rights_element);
    EXPECT_STREQ("<t:EffectiveRights><t:Delete>true</"
                 "t:Delete><t:Modify>true</t:Modify><t:Read>true</t:Read></"
                 "t:EffectiveRights>",
                 effective_rights.to_string().c_str());
}

TEST(InternalTest, GetNonExistingElementReturnsNullptr)
{
    using namespace ews::internal;

    rapidxml::xml_document<> doc;
    auto str = doc.allocate_string("<html lang=\"en\"><head><meta "
                                   "charset=\"utf-8\"/><title>Welcome</"
                                   "title></head><body><h1>Greetings</"
                                   "h1><p>Hello!</p></body></html>");
    doc.parse<0>(str);
    auto body_element = get_element_by_qname(doc, "body", "");
    auto body = xml_subtree(*body_element);
    rapidxml::xml_node<char>* node = nullptr;
    ASSERT_NO_THROW({ node = body.get_node("h1"); });
    EXPECT_EQ(nullptr, node);
}

TEST(InternalTest, GetElementFromSubTree)
{
    using namespace ews::internal;

    rapidxml::xml_document<> doc;
    auto str = doc.allocate_string(contact_card.c_str());
    doc.parse<0>(str);
    auto effective_rights_element =
        get_element_by_qname(doc, "EffectiveRights", uri<>::microsoft::types());
    auto effective_rights = xml_subtree(*effective_rights_element);

    EXPECT_STREQ("true", effective_rights.get_node("Delete")->value());
}

TEST(InternalTest, UpdateSubTreeElement)
{
    using namespace ews::internal;

    rapidxml::xml_document<> doc;
    auto str = doc.allocate_string(contact_card.c_str());
    doc.parse<0>(str);
    auto contact_element =
        get_element_by_qname(doc, "Contact", uri<>::microsoft::types());
    auto cont = xml_subtree(*contact_element);

    cont.set_or_update("Modify", "false");
    EXPECT_STREQ("false", cont.get_node("Modify")->value());
    EXPECT_STREQ("<t:Contact>"
                 "<t:DateTimeSent>2015-05-21T10:13:28Z</t:DateTimeSent>"
                 "<t:DateTimeCreated>2015-05-21T10:13:28Z</t:DateTimeCreated>"
                 "<t:EffectiveRights>"
                 "<t:Delete>true</t:Delete>"
                 "<t:Modify>false</t:Modify>"
                 "<t:Read>true</t:Read>"
                 "</t:EffectiveRights>"
                 "<t:Culture>en-US</t:Culture>"
                 "</t:Contact>",
                 cont.to_string().c_str());
}

TEST(InternalTest, UpdateElementAttribute)
{
    using namespace ews::internal;

    rapidxml::xml_document<> doc;
    auto str = doc.allocate_string(appointment.c_str());
    doc.parse<0>(str);
    auto calitem =
        get_element_by_qname(doc, "CalendarItem", uri<>::microsoft::types());
    auto cont = xml_subtree(*calitem);

    std::vector<xml_subtree::attribute> attrs;
    xml_subtree::attribute attr1 = {"Id", "W. Europe Standard Time"};
    attrs.push_back(attr1);
    xml_subtree::attribute attr2 = {
        "Name", "(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna"};
    attrs.push_back(attr2);

    cont.set_or_update("StartTimeZone", attrs);

    const auto node = cont.get_node("StartTimeZone");
    auto count = 0;
    for_each_attribute(*node,
                       [&](const rapidxml::xml_attribute<>&) { count++; });
    EXPECT_EQ(count, 2);
}

TEST(InternalTest, AppendSubTreeToExistingXMLDocument)
{
    using namespace ews::internal;

    rapidxml::xml_document<> doc;
    {
        const auto xml = std::string("<a>"
                                     "<b>"
                                     "</b>"
                                     "</a>");
        auto str = doc.allocate_string(xml.c_str());
        doc.parse<0>(str);
    }
    auto target_node = get_element_by_qname(doc, "b", "");

    rapidxml::xml_document<> src;
    const auto xml = std::string("<c>"
                                 "<d>"
                                 "</d>"
                                 "</c>");
    auto str = src.allocate_string(xml.c_str());
    src.parse<0>(str);
    auto subtree = xml_subtree(*src.first_node());

    subtree.append_to(*target_node);

    std::string actual;
    rapidxml::print(std::back_inserter(actual), doc,
                    rapidxml::print_no_indenting);
    EXPECT_STREQ("<a>"
                 "<b>"
                 "<c>"
                 "<d/>"
                 "</c>"
                 "</b>"
                 "</a>",
                 actual.c_str());
}

TEST(InternalTest, SubTreeCopyAndAssignment)
{
    using namespace ews::internal;

    rapidxml::xml_document<> doc;
    const auto xml = std::string("<a>\n"
                                 "   <b>\n"
                                 "   </b>\n"
                                 "</a>");
    auto str = doc.allocate_string(xml.c_str());
    doc.parse<0>(str);
    auto a = xml_subtree(*doc.first_node());

    // Self-assignment
    a = a;
    EXPECT_STREQ("<a><b/></a>", a.to_string().c_str());

    auto b = a;
    EXPECT_STREQ("<a><b/></a>", b.to_string().c_str());
    EXPECT_STREQ("<a><b/></a>", a.to_string().c_str());

    auto c(a);
    EXPECT_STREQ("<a><b/></a>", c.to_string().c_str());
    EXPECT_STREQ("<a><b/></a>", a.to_string().c_str());
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
        const auto str = ews::xml_parse_error::error_message_from(exc, data);
        ASSERT_FALSE(str.empty());
        const auto output = split(str, '\n');
        EXPECT_EQ(4U, output.size());
        EXPECT_STREQ("in line 6:", output[0].c_str());
        EXPECT_STREQ("   </head>   <body>     <p>       <h1</h1>     </p>  "
                     " </body> </html>",
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
        const auto str = ews::xml_parse_error::error_message_from(exc, data);
        ASSERT_FALSE(str.empty());
        const auto output = split(str, '\n');
        EXPECT_EQ(4U, output.size());
        EXPECT_STREQ("in line 9:", output[0].c_str());
        EXPECT_STREQ("                             <t:Cultu</t:Culture>    "
                     "                     </t:",
                     output[2].c_str());
        EXPECT_STREQ("                                       ~",
                     output[3].c_str());
    }
}

TEST(InternalTest, HTTPErrorContainsStatusCodeString)
{
    auto error = ews::http_error(404);
    EXPECT_EQ(404, error.code());
    EXPECT_STREQ("HTTP status code: 404 (Not Found)", error.what());
}

#ifdef EWS_HAS_INITIALIZER_LISTS
// References calculated with: https://www.base64encode.org/
TEST(Base64, FillByteNone)
{
    std::vector<unsigned char> text{'a', 'b', 'c'};
    EXPECT_STREQ(ews::internal::base64::encode(text).c_str(), "YWJj");

    const auto decoded = ews::internal::base64::decode("YWJj");
    EXPECT_EQ(decoded, text);
}

TEST(Base64, FillByteOne)
{
    std::vector<unsigned char> text{'a', 'b', 'c', 'd', 'e'};
    EXPECT_STREQ(ews::internal::base64::encode(text).c_str(), "YWJjZGU=");

    const auto decoded = ews::internal::base64::decode("YWJjZGU=");
    EXPECT_EQ(decoded, text);
}

TEST(Base64, FillByteTwo)
{
    std::vector<unsigned char> text{'a', 'b', 'c', 'd'};
    EXPECT_STREQ(ews::internal::base64::encode(text).c_str(), "YWJjZA==");

    const auto decoded = ews::internal::base64::decode("YWJjZA==");
    EXPECT_EQ(decoded, text);
}
#endif

TEST(DateTime, ToEpochInUTCNoDST)
{
    ews::date_time dt("2018-01-08T10:03:30Z");
    EXPECT_EQ(1515405810, dt.to_epoch());
}

TEST(DateTime, ToEpochInUTCAndDST)
{
    ews::date_time dt("2001-10-26T19:32:52Z");
    EXPECT_EQ(1004124772, dt.to_epoch());
}

TEST(DateTime, ToEpochWithOffsetAndDST)
{
    ews::date_time dt1("2001-10-26T21:32:52+02:00");
    EXPECT_EQ(1004124772, dt1.to_epoch());

    ews::date_time dt2("2001-10-26T14:32:52-05:00");
    EXPECT_EQ(1004124772, dt2.to_epoch());
}

TEST(DateTime, ToEpochWithZeroOffsetAndDST)
{
    ews::date_time dt("2001-10-26T19:32:52+00:00");
    EXPECT_EQ(1004124772, dt.to_epoch());
}

TEST(DateTime, ToEpochWithPOSIXReferencePoint)
{
    ews::date_time dt("1986-12-31T23:59:59Z");
    EXPECT_EQ(536457599, dt.to_epoch());
}

TEST(DateTime, ToEpochWithEmptyDateTimeThrows)
{
    EXPECT_THROW(ews::date_time("").to_epoch(), ews::exception);
}

TEST(DateTime, EpochToDateTime)
{
    EXPECT_STREQ("1986-12-31T23:59:59Z",
                 ews::date_time::from_epoch(536457599).to_string().c_str());
    EXPECT_STREQ("2001-10-26T19:32:52Z",
                 ews::date_time::from_epoch(1004124772).to_string().c_str());
    EXPECT_STREQ("2018-01-08T10:03:30Z",
                 ews::date_time::from_epoch(1515405810).to_string().c_str());
}
} // namespace tests
