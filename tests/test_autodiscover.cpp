
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
//
//   This project is hosted at https://github.com/otris

#ifdef EWS_HAS_FILESYSTEM_HEADER

#    include "fixtures.hpp"

#    include <cstring>
#    include <string>
#    include <vector>

namespace tests
{
class AutodiscoverTest : public TemporaryDirectoryFixture
{
public:
    AutodiscoverTest() : creds_("", ""), smtp_address_() {}

    ews::basic_credentials& credentials() { return creds_; }

    std::string& address() { return smtp_address_; }

    void set_next_fake_response(const std::vector<char>& bytes)
    {
        auto& storage = http_request_mock::storage::instance();
        storage.fake_response = bytes;
        storage.fake_response.push_back('\0');
    }

    http_request_mock get_last_request() { return http_request_mock(); }

private:
    ews::basic_credentials creds_;
    std::string smtp_address_;

    void SetUp()
    {
        TemporaryDirectoryFixture::SetUp();
        creds_ =
            ews::basic_credentials("dduck@duckburg.onmicrosoft.com", "quack");
        smtp_address_ = "dduck@duckburg.onmicrosoft.com";
    }
};

TEST_F(AutodiscoverTest, EmptyAddressThrows)
{
    set_next_fake_response(
        read_file(assets_dir() / "autodiscover_response.xml"));

    EXPECT_THROW(
        {
            auto result = ews::get_exchange_web_services_url<http_request_mock>(
                "", credentials());
        },
        ews::exception);
}

TEST_F(AutodiscoverTest, EmptyAddressExceptionText)
{
    set_next_fake_response(
        read_file(assets_dir() / "autodiscover_response.xml"));

    try
    {
        auto result = ews::get_exchange_web_services_url<http_request_mock>(
            "", credentials());
        FAIL() << "Expected an exception";
    }
    catch (ews::exception& exc)
    {
        EXPECT_STREQ("Empty SMTP address given", exc.what());
    }
}

TEST_F(AutodiscoverTest, InvalidAddressThrows)
{
    set_next_fake_response(
        read_file(assets_dir() / "autodiscover_response.xml"));

    EXPECT_THROW(
        {
            auto result = ews::get_exchange_web_services_url<http_request_mock>(
                "typo", credentials());
        },
        ews::exception);
}

TEST_F(AutodiscoverTest, InvalidAddressExceptionText)
{
    set_next_fake_response(
        read_file(assets_dir() / "autodiscover_response.xml"));

    try
    {
        auto result = ews::get_exchange_web_services_url<http_request_mock>(
            "typo", credentials());
        FAIL() << "Expected an exception";
    }
    catch (ews::exception& exc)
    {
        EXPECT_STREQ("No valid SMTP address given", exc.what());
    }
}

TEST_F(AutodiscoverTest, GetExchangeWebServicesURL)
{
    set_next_fake_response(
        read_file(assets_dir() / "autodiscover_response.xml"));

    // Internal should return ASUrl element's value in EXCH protocol
    auto result = ews::get_exchange_web_services_url<http_request_mock>(
        address(), credentials());
    EXPECT_STREQ("https://outlook.office365.com/EWS/Exchange.asmx",
                 result.internal_ews_url.c_str());

    // External should return ASUrl element's value in EXPR protocol
    EXPECT_STREQ("https://outlook.another.office365.com/EWS/Exchange.asmx",
                 result.external_ews_url.c_str());
}

TEST_F(AutodiscoverTest, GetExchangeWebServicesURLWithHint)
{
    auto& s = http_request_mock::storage::instance();
    ews::autodiscover_hints hints;
    hints.autodiscover_url =
        "https://some.domain.com/autodiscover/autodiscover.xml";
    auto result = ews::get_exchange_web_services_url<http_request_mock>(
        address(), credentials(), hints);
    EXPECT_STREQ("https://some.domain.com/autodiscover/autodiscover.xml",
                 s.url.c_str());
}

TEST_F(AutodiscoverTest, GetExchangeWebServicesURLWithoutHint)
{
    auto& s = http_request_mock::storage::instance();
    auto result = ews::get_exchange_web_services_url<http_request_mock>(
        address(), credentials());
    // address: dduck@duckburg.onmicrosoft.com
    EXPECT_STREQ(
        "https://duckburg.onmicrosoft.com/autodiscover/autodiscover.xml",
        s.url.c_str());
}

TEST_F(AutodiscoverTest, GetExchangeWebServicesURLThrowsOnError)
{
    // A response that is returned by Autodiscover if the SMTP
    // address is unknown to the system. HTTP status code is 200, but XML
    // content indicates an error
    set_next_fake_response(
        read_file(assets_dir() / "autodiscover_response_error.xml"));

    // External should return ASUrl element's value in EXPR protocol
    EXPECT_THROW(
        {
            ews::get_exchange_web_services_url<http_request_mock>(
                address(), credentials());
        },
        ews::exception);
}

TEST_F(AutodiscoverTest, GetExchangeWebServicesURLExceptionText)
{
    set_next_fake_response(
        read_file(assets_dir() / "autodiscover_response_error.xml"));

    try
    {
        ews::get_exchange_web_services_url<http_request_mock>(address(),
                                                              credentials());
        FAIL() << "Expected an exception";
    }
    catch (ews::exception& exc)
    {
        EXPECT_STREQ("The email address can't be found. (error code: 500)",
                     exc.what());
    }
}
} // namespace tests

#endif // EWS_HAS_FILESYSTEM_HEADER
