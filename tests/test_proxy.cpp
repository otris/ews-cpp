
//   Copyright 2016-2020 otris software AG
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

#include <ews/ews.hpp>

namespace tests
{
TEST_F(ProxyTest, NoTunnel)
{
    const auto& proxy = get_proxy_uri();

    if (!proxy.empty())
    {
        service().set_proxy(proxy);
        service().set_http_proxy_tunnel(false);
        auto response =
            service().resolve_names("name", ews::search_scope::active_directory);
        EXPECT_EQ(response.resolutions.size(), 0);
    }
}
TEST_F(ProxyTest, Tunnel)
{
    const auto& proxy = get_proxy_uri();

    if (!proxy.empty())
    {
        service().set_proxy(proxy);
        service().set_http_proxy_tunnel(true);
        auto response =
            service().resolve_names("name", ews::search_scope::active_directory);
        EXPECT_EQ(response.resolutions.size(), 0);
    }
}
} // namespace tests
