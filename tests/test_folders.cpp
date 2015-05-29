#include "fixture.hpp"

using folder_id = ews::folder_id;
using distinguished_folder_id = ews::distinguished_folder_id;
using standard_folder = ews::standard_folder;

namespace tests
{
    TEST(FolderTest, DistinguishedFolderIdFromWellKnownName)
    {
        distinguished_folder_id id = standard_folder::inbox;
        EXPECT_STREQ("<x:DistinguishedFolderId Id=\"inbox\" />",
                     id.to_xml("x").c_str());
    }
}
