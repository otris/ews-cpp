#pragma once

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
}
