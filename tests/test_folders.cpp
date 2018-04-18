
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

#include "fixtures.hpp"

namespace tests
{
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

#ifdef _MSC_VER
#    pragma warning(suppress : 6262)
#endif
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
    EXPECT_STREQ(expected, a.to_xml().c_str());
}

TEST(FolderTest, ToXMLWithoutChangeKey)
{
    const char* expected = "<t:FolderId Id=\"abcde\"/>";
    const auto a = ews::folder_id("abcde");
    EXPECT_STREQ(expected, a.to_xml().c_str());
}

TEST(FolderTest, DistinguishedFolderIdToXML)
{
    const char* expected = "<t:DistinguishedFolderId Id=\"contacts\"/>";
    const auto folder =
        ews::distinguished_folder_id(ews::standard_folder::contacts);
    EXPECT_STREQ(expected, folder.to_xml().c_str());
}

TEST(FolderTest, DistinguishedFolderIdToXMLWithChangeKey)
{
    const char* expected =
        "<t:DistinguishedFolderId Id=\"tasks\" ChangeKey=\"abcde\"/>";
    const auto folder =
        ews::distinguished_folder_id(ews::standard_folder::tasks, "abcde");
    EXPECT_STREQ(expected, folder.to_xml().c_str());
}

TEST(FolderTest, DistinguishedFolderIdAndMailboxOwnerToXML)
{
    const char* expected = "<t:DistinguishedFolderId Id=\"inbox\">"
                           "<t:Mailbox><t:EmailAddress>test@example.com</"
                           "t:EmailAddress></t:Mailbox>"
                           "</t:DistinguishedFolderId>";
    const auto owner = ews::mailbox("test@example.com");
    const auto folder =
        ews::distinguished_folder_id(ews::standard_folder::inbox, owner);
    EXPECT_STREQ(expected, folder.to_xml().c_str());
}

TEST(FolderTest, DistinguishedFolderIsAlwaysValid)
{
    auto folder = ews::distinguished_folder_id(ews::standard_folder::inbox);
    EXPECT_TRUE(folder.valid());
}

TEST(FolderTest, DistinguishedFolderIdFromWellKnownName)
{
    // Test conversion c'tor
    ews::distinguished_folder_id id = ews::standard_folder::inbox;
    EXPECT_STREQ("<t:DistinguishedFolderId Id=\"inbox\"/>",
                 id.to_xml().c_str());
}

TEST(FolderTest, DistinguishedFolderIdAndChangeKeyAttribute)
{
    auto folder = ews::distinguished_folder_id(ews::standard_folder::calendar);
    EXPECT_STREQ("calendar", folder.id().c_str());
}
} // namespace tests
