#include <ews/plumbing.hpp>
#include <ews/porcelain.hpp>

#include <string>
#include <iostream>
#include <ostream>

int main()
{
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
        auto doc = ews::plumbing::make_raw_soap_request(
            "https://yourserver/ews/Exchange.asmx", "username", "yourPassword",
            "yourDomain", request, soap_headers);
        (void)doc;

        // hack hack ...

    }
    catch (ews::curl::curl_error& e)
    {
        std::cout << e.what() << std::endl;
    }

    ews::tear_down();
    return 0;
}
