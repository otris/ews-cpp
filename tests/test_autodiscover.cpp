#ifdef EWS_USE_BOOST_LIBRARY

#include "fixtures.hpp"

#include <string>
#include <vector>
#include <cstring>

namespace tests
{
    class AutodiscoverTest : public AssetsFixture
    {
    public:
        AutodiscoverTest()
            : creds_("", ""),
              smtp_address_()
        {
        }

        ews::basic_credentials& credentials() { return creds_; }

        std::string& address() { return smtp_address_; }

        void set_next_fake_response(const std::vector<char>& bytes)
        {
            auto& storage = http_request_mock::storage::instance();
            storage.fake_response = bytes;
            storage.fake_response.push_back('\0');
        }

        http_request_mock get_last_request()
        {
            return http_request_mock();
        }

    private:
        ews::basic_credentials creds_;
        std::string smtp_address_;

        void SetUp()
        {
            AssetsFixture::SetUp();
            const auto& env = ews::test::global_data::instance().env;
            creds_ = ews::basic_credentials(env.autodiscover_smtp_address,
                                            env.autodiscover_password);
            smtp_address_ = env.autodiscover_smtp_address;
        }
    };

    TEST_F(AutodiscoverTest, GetExchangeWebServicesURL)
    {
        set_next_fake_response(read_file(
            assets_dir() / "autodiscover_response.xml"));

        // Internal should return ASUrl element's value in EXCH protocol
        auto result =
            ews::get_exchange_web_services_url<http_request_mock>(
                address(),
                ews::autodiscover_protocol::internal,
                credentials());
        EXPECT_STREQ("https://outlook.office365.com/EWS/Exchange.asmx",
                     result.c_str());

        // External should return ASUrl element's value in EXPR protocol
        result =
            ews::get_exchange_web_services_url<http_request_mock>(
                address(),
                ews::autodiscover_protocol::external,
                credentials());
        EXPECT_STREQ("https://outlook.another.office365.com/EWS/Exchange.asmx",
                     result.c_str());
    }

    TEST_F(AutodiscoverTest, GetExchangeWebServicesURLThrowsOnError)
    {
        // A response that is returned by Autodiscover if the SMTP
        // address is unknown to the system. HTTP status code is 200, but XML
        // content indicates an error
        set_next_fake_response(read_file(
            assets_dir() / "autodiscover_response_error.xml"));

        // External should return ASUrl element's value in EXPR protocol
        EXPECT_THROW(
        {
            ews::get_exchange_web_services_url<http_request_mock>(
                address(),
                ews::autodiscover_protocol::internal,
                credentials());
        }, ews::exception);
    }

    TEST_F(AutodiscoverTest, GetExchangeWebServicesURLExceptionText)
    {
        set_next_fake_response(read_file(
            assets_dir() / "autodiscover_response_error.xml"));

        try
        {
            ews::get_exchange_web_services_url<http_request_mock>(
                address(),
                ews::autodiscover_protocol::internal,
                credentials());
            FAIL() << "Expected an exception";
        }
        catch (ews::exception& exc)
        {
            EXPECT_STREQ(
                "The email address can't be found. (error code: 500)",
                exc.what());
        }
    }
}

#endif // EWS_USE_BOOST_LIBRARY

// vim:et ts=4 sw=4 noic cc=80
