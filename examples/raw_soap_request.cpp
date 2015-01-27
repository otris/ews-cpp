#include <ews/ews.hpp>
#include <ews/rapidxml/rapidxml_print.hpp>

#include <string>
#include <iostream>
#include <ostream>
#include <exception>
#include <cstdlib>

namespace
{
    const std::string server_uri = "https://columbus.test.otris.de/ews/Exchange.asmx";
    const std::string domain = "TEST";
    const std::string password = "12345aA!";
    const std::string username = "mini";

    std::string request{ R"(
<m:GetFolder>
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
</m:GetFolder>)" };

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
        auto response = ews::internal::make_raw_soap_request(server_uri,
                username, password, domain, request, soap_headers);
        const auto& doc = response.payload();

        std::cout << doc << std::endl;

        // hack hack ...

        ews::internal::raise_exception_if_soap_fault(response);
    }
    catch (std::exception& exc)
    {
        std::cout << exc.what() << std::endl;
        res = EXIT_FAILURE;
    }

    ews::tear_down();
    return res;
}
