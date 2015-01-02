#include <ews/plumbing.hpp>
#include <ews/porcelain.hpp>

#include <string>
#include <iostream>
#include <ostream>
#include <exception>
#include <cstdlib>

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    std::string request{
        R"(
        <GetFolder>
            <FolderShape>
                <t:BaseShape>AllProperties</t:BaseShape>
            </FolderShape>
            <FolderIds>
                <t:DistinguishedFolderId Id="inbox" />
            </FolderIds>
        </GetFolder>
        )"};

    auto soap_headers = std::vector<std::string>{};

    try
    {
        ews::plumbing::make_raw_soap_request(
            "https://192.168.56.2/ews/Exchange.asmx", "mini", "12345aA!",
            "TEST", request, soap_headers);
        //(void)doc;

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
