#include "fixtures.hpp"

namespace tests
{
    TEST(FolderTest, DistinguishedFolderIdFromWellKnownName)
    {
        ews::distinguished_folder_id id = ews::standard_folder::inbox;
        EXPECT_STREQ("<x:DistinguishedFolderId Id=\"inbox\" />",
                     id.to_xml("x").c_str());
    }
}
