#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <ews/rapidxml/rapidxml_print.hpp>

#include <string>
#include <iostream>
#include <ostream>
#include <exception>
#include <cstdlib>

using namespace ews::internal;

namespace
{
    std::string request(
        "<m:GetFolder>\n"
        "<m:FolderShape>\n"
        "    <t:BaseShape>IdOnly</t:BaseShape>\n"
        "    <t:AdditionalProperties>\n"
        "    <t:FieldURI FieldURI=\"folder:DisplayName\" />\n"
        "    <t:FieldURI FieldURI=\"folder:ChildFolderCount\" />\n"
        "    </t:AdditionalProperties>\n"
        "</m:FolderShape>\n"
        "<m:FolderIds>\n"
        "    <t:DistinguishedFolderId Id=\"root\" />\n"
        "</m:FolderIds>\n"
        "</m:GetFolder>\n");
}

int main()
{
    int res = EXIT_SUCCESS;
    ews::set_up();

    auto soap_headers = std::vector<std::string>();
    soap_headers.emplace_back(
        "<t:RequestServerVersion Version=\"Exchange2013_SP1\"/>"
    );

    try
    {
        const auto env = ews::test::environment();
        auto response = make_raw_soap_request<http_request>(env.server_uri,
                                                            env.username,
                                                            env.password,
                                                            env.domain,
                                                            request,
                                                            soap_headers);
        const auto doc = parse_response(std::move(response));
        std::cout << *doc << std::endl;

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

// vim:et ts=4 sw=4 noic cc=80
