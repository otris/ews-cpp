// Wrapper around curl.h header file. Wraps everything in that header file with
// C++ namespace ews::curl; also defines exception for cURL related runtime
// errors.

#pragma once

#include <stdexcept>

namespace ews
{
    namespace curl
    {

#include <curl/curl.h>

        class curl_error : public std::runtime_error
        {
        public:
            explicit curl_error(const std::string& what)
                : std::runtime_error(what)
            {
            }
            explicit curl_error(const char* what) : std::runtime_error(what) {}
        };

        class curl_ptr
        {
        public:
            curl_ptr() : handle_{curl_easy_init()}
            {
                if (!handle_)
                {
                    throw curl_error{"Could not start libcurl session"};
                }
            }

            ~curl_ptr() { curl_easy_cleanup(handle_); }

            curl_ptr(const curl_ptr&) = delete; // Could use curl_easy_duphandle
                                                // for this
            curl_ptr& operator=(const curl_ptr&) = delete;

            CURL* get() const { return handle_; }

        private:
            CURL* handle_;
        };
    }
}
