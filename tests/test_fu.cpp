#include <ews/ews.hpp>
#include <ews/rapidxml/rapidxml_print.hpp>

#include <string>
#include <iostream>
#include <ostream>
#include <exception>
#include <cstdlib>

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();
    std::string server_uri = "https://columbus.test.otris.de/ews/Exchange.asmx";
    std::string domain = "TEST";
    std::string password = "12345aA!";
    std::string username = "mini";

    std::string request{ R"(
<GetFolder>
    <FolderShape>
        <t:BaseShape>AllProperties</t:BaseShape>
    </FolderShape>
    <FolderIds>
        <t:DistinguishedFolderId Id="inbox" />
    </FolderIds>
</GetFolder>)" };

    auto soap_headers = std::vector<std::string>{};

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
