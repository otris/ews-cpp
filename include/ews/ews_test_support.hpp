// Global data used in tests and examples; initialized at program-start

#pragma once

#include <string>
#include <stdexcept>
#include <cstdlib>

namespace ews
{
    namespace test
    {
        class environment final
        {
        public:
            std::string domain;
            std::string username;
            std::string password;
            std::string server_uri;
            std::string autodiscover_smtp_address;
            std::string autodiscover_password;

            environment()
                : domain(getenv_or_throw("EWS_TEST_DOMAIN")),
                  username(getenv_or_throw("EWS_TEST_USERNAME")),
                  password(getenv_or_throw("EWS_TEST_PASSWORD")),
                  server_uri(getenv_or_throw("EWS_TEST_URI")),
                  autodiscover_smtp_address(getenv_or_empty_string(
                      "EWS_TEST_AUTODISCOVER_SMTP_ADDRESS")),
                  autodiscover_password(
                      getenv_or_empty_string("EWS_TEST_AUTODISCOVER_PASSWORD"))
            {
            }

        private:

#ifdef _MSC_VER
# pragma warning(push)
# pragma warning(disable : 4996)
#endif
            static std::string getenv_or_throw(const char* name)
            {
                const char* const val = getenv(name);
                if (val == nullptr)
                {
                    const auto msg =
                        std::string("Missing environment variable ") + name;
                    throw std::runtime_error(msg);
                }
                return std::string(val);
            }

            static std::string getenv_or_empty_string(const char* name)
            {
                const char* const val = getenv(name);
                return val == nullptr ? "" : std::string(val);
            }
#ifdef _MSC_VER
# pragma warning(pop)
#endif
        };

        struct global_data final
        {
            std::string assets_dir;
            environment env;

            static global_data& instance()
            {
                static global_data inst;
                return inst;
            }
        };
    }
}

// vim:et ts=4 sw=4 noic cc=80
