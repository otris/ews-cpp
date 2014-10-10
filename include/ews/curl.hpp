// Wrapper around curl.h header file. Wraps everything in that header file with
// C++ namespace ews::curl; also defines exception for cURL related runtime
// errors.

#pragma once

#include <stdexcept>
#include <string>


// Visual Studio 12 does not support noexcept specifier.
#ifndef _MSC_VER
# define NOEXCEPT noexcept
#else
# define EWS_NOEXCEPT 
#endif

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

        // Helper function; constructs an exception with a meaningful error
        // message from the given result code for the most recent cURL API call.
        //
        // msg: A string that prepends the actual cURL error message.
        // rescode: The result code of a failed cURL operation.
        inline curl_error make_error(const std::string& msg, CURLcode rescode)
        {
            std::string reason{curl_easy_strerror(rescode)};
            return curl_error{msg + ": \'" + reason + "\'"};
        }

        // RAII helper class for CURL* handles.
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

            // Could use curl_easy_duphandle for copying
            curl_ptr(const curl_ptr&) = delete;
            curl_ptr& operator=(const curl_ptr&) = delete;

            CURL* get() const EWS_NOEXCEPT { return handle_; }

        private:
            CURL* handle_;
        };

        // RAII wrapper class around cURLs slist construct.
        class curl_string_list
        {
        public:
            curl_string_list() EWS_NOEXCEPT : slist_{ nullptr }{}

            ~curl_string_list() { curl_slist_free_all(slist_); }

            void append(const char* str) EWS_NOEXCEPT
            {
                slist_ = curl_slist_append(slist_, str);
            }

            curl_slist* get() const EWS_NOEXCEPT { return slist_; }

        private:
            curl_slist* slist_;
        };
    }
}
