#include <ews/ews.hpp>
#include <gtest/gtest.h>
#include <cstring>

TEST(ConstantsTest, NamespaceURIs)
{
    using uri = ews::internal::uri<>;

    ASSERT_EQ(uri::microsoft::errors_size,
              std::strlen(uri::microsoft::errors()));
    ASSERT_EQ(uri::microsoft::types_size,
              std::strlen(uri::microsoft::types()));
    ASSERT_EQ(uri::microsoft::messages_size,
              std::strlen(uri::microsoft::messages()));

    ASSERT_EQ(uri::soapxml::envelope_size,
              std::strlen(uri::soapxml::envelope()));
}
