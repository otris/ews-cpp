#include "fixtures.hpp"

namespace tests
{
    TEST(FolderTest, DistinguishedFolderIdFromWellKnownName)
    {
        ews::distinguished_folder_id id = ews::standard_folder::inbox;
        EXPECT_STREQ("<x:DistinguishedFolderId Id=\"inbox\" />",
                     id.to_xml("x").c_str());
    }

    TEST(FolderTest, ConstructWithIdOnly)
    {
        auto a = ews::folder_id("abcde");
        EXPECT_STREQ("abcde", a.id().c_str());
    }

    TEST(FolderTest, DefaultConstruction)
    {
        auto a = ews::folder_id();
        EXPECT_FALSE(a.valid());
        EXPECT_STREQ("", a.id().c_str());
    }

    TEST(FolderTest, ConstructWithIdAndChangeKey)
    {
        auto a = ews::folder_id("abcde", "edcba");
        EXPECT_STREQ("abcde", a.id().c_str());
        EXPECT_STREQ("edcba", a.change_key().c_str());
    }

#pragma warning(suppress: 6262)
    TEST(FolderTest, FromParentFolderIdXMLNode)
    {
        char buf[] = "<ParentFolderId Id=\"abcde\" ChangeKey=\"edcba\" />";
        rapidxml::xml_document<> doc;
        doc.parse<0>(buf);
        auto node = doc.first_node();
        auto a = ews::folder_id::from_xml_element(*node);
        EXPECT_STREQ("abcde", a.id().c_str());
        EXPECT_STREQ("edcba", a.change_key().c_str());
    }

    TEST(FolderTest, ToXMLWithNamespace)
    {
        const char* expected = "<t:FolderId Id=\"abcde\" ChangeKey=\"edcba\"/>";
        const auto a = ews::folder_id("abcde", "edcba");
        EXPECT_STREQ(expected, a.to_xml("t").c_str());
    }

    TEST(FolderTest, DistinguishedFolderIsAlwaysValid)
    {
        auto folder = ews::distinguished_folder_id(ews::standard_folder::inbox);
        EXPECT_TRUE(folder.valid());
    }
}
