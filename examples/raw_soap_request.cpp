#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <ews/rapidxml/rapidxml_print.hpp>

#include <string>
#include <iostream>
#include <ostream>
#include <exception>
#include <cstdlib>

namespace
{
    std::string request(
R"(<m:GetFolder>
<m:FolderShape>
    <t:BaseShape>IdOnly</t:BaseShape>
    <t:AdditionalProperties>
    <t:FieldURI FieldURI="folder:DisplayName" />
    <t:FieldURI FieldURI="folder:ChildFolderCount" />
    </t:AdditionalProperties>
</m:FolderShape>
<m:FolderIds>
    <t:DistinguishedFolderId Id="root" />
</m:FolderIds>
</m:GetFolder>)");

    const auto soap_headers = std::vector<std::string> {
        "<t:RequestServerVersion Version=\"Exchange2013_SP1\"/>"
    };
}

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    try
    {
        const auto env = ews::test::get_from_environment();
        auto response = ews::internal::make_raw_soap_request(env.server_uri,
                                                             env.username,
                                                             env.password,
                                                             env.domain,
                                                             request,
                                                             soap_headers);
        const auto& doc = response.payload();
        std::cout << doc << std::endl;

        // hack hack ...

    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}
