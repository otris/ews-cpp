// Global data used in tests and examples; initialized at program-start

#pragma once

#include <string>
#include <stdexcept>
#include <cstdlib>

namespace ews
{
    namespace test
    {
        namespace internal
        {
            inline std::string get_or_throw_env(const char* name)
            {
#ifdef _MSC_VER
# pragma warning(push)
# pragma warning(disable: 4996)
#endif

                const char* const env = getenv(name);

#ifdef _MSC_VER
# pragma warning(pop)
#endif

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
        };

        inline credentials get_from_environment()
        {
            using internal::get_or_throw_env;

            credentials cred;
            cred.domain = get_or_throw_env("EWS_TEST_DOMAIN");
            cred.username = get_or_throw_env("EWS_TEST_USERNAME");
            cred.password = get_or_throw_env("EWS_TEST_PASSWORD");
            cred.server_uri = get_or_throw_env("EWS_TEST_URI");
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

// vim:et ts=4 sw=4 noic cc=80
