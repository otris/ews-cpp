#pragma once

#include <string>
#include <vector>
#include <sstream>
#include <utility>
#include <memory>
#include <cstddef>

#include "curl.hpp"
#include "assert.hpp"

namespace ews
{
    namespace plumbing
    {
        class http_web_request;

        // This ought to be a DOM wrapper; usually around a web response
        class xml_document final
        {
        };

        class credentials
        {
        public:
            virtual ~credentials() = default;
            virtual void certify(http_web_request*) const = 0;
        };

        class ntlm_credentials : public credentials
        {
        public:
            ntlm_credentials(std::string username, std::string password,
                             std::string domain)
                : username_{std::move(username)},
                  password_{std::move(password)},
                  domain_{std::move(domain)}
            {
            }

        private:
            void certify(http_web_request*) const override; // implemented below

            std::string username_;
            std::string password_;
            std::string domain_;
        };

        class http_web_request final
        {
        public:
            enum class method
            {
                POST
            };

            // Create a new HTTP request to the given URL.
            explicit http_web_request(const std::string& url)
            {
                set_option(curl::CURLOPT_URL, url.c_str());
            }

            // Set the HTTP method (only POST supported).
            void set_method(method)
            {
                // Method can only be a regular POST in our use case
                set_option(curl::CURLOPT_POST, 1);
            }

            // Set this HTTP request's content type.
            void set_content_type(const std::string& content_type)
            {
                curl::curl_string_list chunk;
                chunk.append(content_type.c_str());
                set_option(curl::CURLOPT_HTTPHEADER, chunk.get());
            }

            // Set credentials for authentication.
            void set_credentials(const credentials& creds)
            {
                creds.certify(this);
            }

            // Small wrapper around curl_easy_setopt(3).
            //
            // Converts return codes to exceptions of type curl_error and allows
            // objects other than http_web_requests to set transfer options
            // without direct access to the internal CURL handle.
            template <typename... Args>
            void set_option(curl::CURLoption option, Args... args)
            {
                auto retcode = curl::curl_easy_setopt(
                    handle_.get(), option, std::forward<Args>(args)...);
                switch (retcode)
                {
                case curl::CURLE_OK:
                    return;

                case curl::CURLE_FAILED_INIT:
                {
                    throw curl::make_error(
                        "curl_easy_setopt: unsupported option", retcode);
                }

                default:
                {
                    throw curl::make_error(
                        "curl_easy_setopt: failed setting option", retcode);
                }
                };
            }

            // Perform the HTTP request.
            //
            // request: The complete request string; you must make sure that the
            // data is encoded the way you want the server to receive it.
            // response_handler: A callable function or function pointer that
            // takes a std::string (the response) object as only argument. The
            // function is invoked as soon as a response is received.
            //
            // Throws curl_error if operation could not be completed.
            template <typename Function>
            void send(const std::string& request, Function response_handler)
            {
                // Set complete request string for HTTP POST method; note: no
                // encoding here

                set_option(curl::CURLOPT_POSTFIELDS, request.c_str());
                set_option(curl::CURLOPT_POSTFIELDSIZE, request.length());

                // Set response handler; wrap passed handler function in
                // something that we can cast to a void* (we cannot cast
                // function pointers to void*) to make sure this works with
                // lambdas, std::function instances, function pointers, and
                // functors

                struct wrapper
                {
                    Function callable;

                    explicit wrapper(Function function) : callable{function} {}
                };

                wrapper wrp{response_handler};

                auto callback =
                    [](char* ptr, std::size_t size, std::size_t nmemb,
                       void* userdata) -> std::size_t
                {
                    // Call user supplied response handler with a std::string
                    // (contains the complete response). Always indicate
                    // success; user has no chance of indicating an error to the
                    // cURL library.

                    wrapper* w = reinterpret_cast<wrapper*>(userdata);
                    const auto count = size * nmemb;
                    std::string response{ptr, count};
                    w->callable(std::move(response));
                    return count;
                };

                set_option(
                    curl::CURLOPT_WRITEFUNCTION,
                    static_cast<std::size_t (*)(char*, std::size_t, std::size_t,
                                                void*)>(callback));
                set_option(curl::CURLOPT_WRITEDATA, std::addressof(wrp));

                auto retcode = curl::curl_easy_perform(handle_.get());
                if (retcode != 0)
                {
                    throw curl::make_error("curl_easy_perform", retcode);
                }
            }

        private:
            curl::curl_ptr handle_;
        };

        // Makes a raw SOAP request.
        //
        // url: The URL of the server to talk to.
        // username: The username of user.
        // password: The user's secret password, plain-text.
        // domain: The user's Windows domain.
        // soap_body: The contents of the SOAP body (minus the body element);
        // this is the actual EWS request.
        // soap_headers: Any SOAP headers to add.
        //
        // Returns a DOM wrapper around the response.
        inline xml_document make_raw_soap_request(
            const std::string& url, const std::string& username,
            const std::string& password, const std::string& domain,
            const std::string& soap_body,
            const std::vector<std::string>& soap_headers)
        {
            http_web_request request{url};
            request.set_method(http_web_request::method::POST);
            request.set_content_type("text/xml;utf-8");

            ntlm_credentials creds{username, password, domain};
            request.set_credentials(creds);

            std::stringstream request_stream;
            request_stream << R"(
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
    >
)";

            // Add SOAP headers if present
            if (!soap_headers.empty())
            {
                request_stream << "<soap:Header>\n";
                for (const auto& header : soap_headers)
                {
                    request_stream << header;
                }
                request_stream << "</soap:Header>\n";
            }

            request_stream << "<soap:Body>\n";
            // Add the passed request
            request_stream << soap_body;
            request_stream << "</soap:Body>\n";
            request_stream << "</soap:Envelope>\n";

            xml_document doc;
            request.send(request_stream.str(), [&doc](std::string response)
            {
                // TODO: load XML from response into doc
                // Note: response is a copy and can therefore be moved
                (void)response;
            });

            return doc;
        }

        // Implementations

        void ntlm_credentials::certify(http_web_request* request) const
        {
            EWS_ASSERT(request != nullptr);

            // CURLOPT_USERPWD: domain\username:password
            std::string login = domain_ + "\\" + username_ + "" + password_;
            request->set_option(curl::CURLOPT_USERPWD, login.c_str());
        }
    }
}
