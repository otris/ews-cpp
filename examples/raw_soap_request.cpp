
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

#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <ews/rapidxml/rapidxml_print.hpp>
#include <exception>
#include <iostream>
#include <ostream>
#include <stdlib.h>
#include <string>

using namespace ews::internal;

namespace
{
std::string request("<m:GetFolder>\n"
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
        "<t:RequestServerVersion Version=\"Exchange2013_SP1\"/>");

    try
    {
        const auto env = ews::test::environment();
        auto response = make_raw_soap_request<http_request>(
            env.server_uri, env.username, env.password, env.domain, request,
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
