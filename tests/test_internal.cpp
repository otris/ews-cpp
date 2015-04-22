#include <ews/ews.hpp>
#include <gtest/gtest.h>
#include <vector>
#include <string>
#include <sstream>
#include <algorithm>
#include <cstring>

namespace
{
    std::vector<std::string> split(const std::string& str, char sep='\n')
    {
        std::stringstream sstr(str);
        std::string line;
        std::vector<std::string> result;
        while (std::getline(sstr, line, sep))
        {
            result.push_back(line);
        }
        return result;
    }
}

namespace tests
{
    TEST(InternalTest, NamespaceURIs)
    {
        using uri = ews::internal::uri<>;

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
        using mime_content = ews::mime_content;

        auto m = mime_content();
        EXPECT_TRUE(m.none());
        EXPECT_EQ(0U, m.len_bytes());
        EXPECT_TRUE(m.character_set().empty());
        EXPECT_EQ(nullptr, m.bytes());
    }

    TEST(InternalTest, MimeContentConstructionWithData)
    {
        using mime_content = ews::mime_content;
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

        EXPECT_TRUE(std::equal(m.bytes(), m.bytes() + m.len_bytes(),
                               content, content + std::strlen(content)));
    }

    TEST(InternalTest, PropertyClass)
    {
        using property_path = ews::property_path;
        using property = ews::property<property_path::contact>;

        auto p = property(property_path::contact::surname, "Duck");
        EXPECT_EQ(property_path::contact::surname, p.path());
        EXPECT_STREQ("Duck", p.value().c_str());
        EXPECT_STREQ("Surname", p.element_name().c_str());
        EXPECT_STREQ("<Surname>Duck</Surname>", p.to_xml().c_str());
        EXPECT_STREQ("<e:Surname>Duck</e:Surname>", p.to_xml("e").c_str());
    }

    TEST(InternalTest, PropertyMapClass)
    {
        using path = ews::property_path::contact;
        using property_map = ews::internal::property_map<path>;

        auto map = property_map();
        EXPECT_STREQ("", map.get(path::surname).c_str());
        map.set_or_update(path::surname, "Nobody");
        EXPECT_STREQ("Nobody", map.get(path::surname).c_str());
        map.set_or_update(path::surname, "Raskolnikov");
        EXPECT_STREQ("Raskolnikov", map.get(path::surname).c_str());
        map.set_or_update(path::given_name, "Rodion Romanovich");

        auto strlist = split(map.to_xml());
        ASSERT_EQ(2U, strlist.size());
        EXPECT_EQ(1U, std::count(begin(strlist), end(strlist),
                    "<Surname>Raskolnikov</Surname>"));
        EXPECT_EQ(1U, std::count(begin(strlist), end(strlist),
                    "<GivenName>Rodion Romanovich</GivenName>"));

        strlist = split(map.to_xml("n"));
        ASSERT_EQ(2U, strlist.size());
        EXPECT_EQ(1U, std::count(begin(strlist), end(strlist),
                    "<n:Surname>Raskolnikov</n:Surname>"));
        EXPECT_EQ(1U, std::count(begin(strlist), end(strlist),
                    "<n:GivenName>Rodion Romanovich</n:GivenName>"));
    }
}
