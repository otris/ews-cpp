#include <ews/ews.hpp>
#include <gtest/gtest.h>
#include <cstring>

namespace tests
{
    TEST(InternalConstantsTest, NamespaceURIs)
    {
        using uri = ews::internal::uri<>;

        EXPECT_EQ(uri::microsoft::errors_size,
                  std::strlen(uri::microsoft::errors()));
        EXPECT_EQ(uri::microsoft::types_size,
                  std::strlen(uri::microsoft::types()));
        EXPECT_EQ(uri::microsoft::messages_size,
                  std::strlen(uri::microsoft::messages()));

        EXPECT_EQ(uri::soapxml::envelope_size,
                  std::strlen(uri::soapxml::envelope()));
    }
}
