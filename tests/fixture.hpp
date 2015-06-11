#pragma once

#include <memory>

#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <gtest/gtest.h>

namespace tests
{
    // Global data used in tests; initialized at program-start
    struct Environment
    {
        static const ews::test::credentials& credentials()
        {
            static auto creds = ews::test::get_from_environment();
            return creds;
        }
    };

    // Per-test-case set-up and tear-down
    class BaseFixture : public ::testing::Test
    {
    protected:
        static void SetUpTestCase()
        {
            ews::set_up();
        }

        static void TearDownTestCase()
        {
            ews::tear_down();
        }
    };

    // Set-up and tear down a service object
    class ServiceFixture : public BaseFixture
    {
    public:
        virtual ~ServiceFixture() = default;

        ews::service& service()
        {
            return *service_ptr_;
        }

    protected:
        virtual void SetUp()
        {
            BaseFixture::SetUp();
            const auto& creds = Environment::credentials();
#ifdef EWS_HAS_MAKE_UNIQUE
            service_ptr_ = std::make_unique<ews::service>(creds.server_uri,
                                                          creds.domain,
                                                          creds.username,
                                                          creds.password);
#else
            service_ptr_ = std::unique_ptr<ews::service>(
                                         new ews::service(creds.server_uri,
                                                          creds.domain,
                                                          creds.username,
                                                          creds.password));
#endif
        }

        virtual void TearDown()
        {
            service_ptr_.reset();
            BaseFixture::TearDown();
        }

    private:
        std::unique_ptr<ews::service> service_ptr_;
    };
}
