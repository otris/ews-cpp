// Global data used in tests and examples; initialized at program-start

#pragma once

#include <string>
#include <stdexcept>
#include <cstdlib>

#ifdef _MSC_VER
# pragma warning(push)
# pragma warning(disable: 4996)
#endif

namespace ews
{
    namespace test
    {
        namespace internal
        {
            inline std::string getenv_or_throw(const char* name)
            {
                const char* const env = getenv(name);
                if (env == nullptr)
                {
                    const auto msg
                        = std::string("Missing environment variable ") + name;
                    throw std::runtime_error(msg);
                }
                return std::string(env);
            }
        }

        struct credentials
        {
            std::string domain;
            std::string username;
            std::string password;
            std::string server_uri;
            std::string autodiscover_smtp_address;
            std::string autodiscover_password;
        };

        inline credentials get_from_environment()
        {
            using internal::getenv_or_throw;

            credentials cred;
            cred.domain = getenv_or_throw("EWS_TEST_DOMAIN");
            cred.username = getenv_or_throw("EWS_TEST_USERNAME");
            cred.password = getenv_or_throw("EWS_TEST_PASSWORD");
            cred.server_uri = getenv_or_throw("EWS_TEST_URI");
            cred.autodiscover_smtp_address =
                getenv("EWS_TEST_AUTODISCOVER_SMTP_ADDRESS");
            cred.autodiscover_password =
                getenv("EWS_TEST_AUTODISCOVER_PASSWORD");
            return cred;
        }

        struct global_data
        {
            std::string assets_dir;

            static global_data& instance()
            {
                static global_data inst;
                return inst;
            }
        };
    }
}

#ifdef _MSC_VER
# pragma warning(pop)
#endif

// vim:et ts=4 sw=4 noic cc=80
