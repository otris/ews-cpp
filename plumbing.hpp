#pragma once

#include <string>
#include <vector>
#include <sstream>
#include <utility>
#include <cassert> // TODO: wrap assert with another macro so that we can make
                   // assertions a configuration option

#include "curl.hpp"

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
            ntlm_credentials(const std::string& username,
                             const std::string& password,
                             const std::string& domain)
                : username_{username}, password_{password}, domain_{domain}
            {
            }

        private:
            std::string username_;
            std::string password_;
            std::string domain_;

            void certify(http_web_request*) const override; // implemented below
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

            // Set the HTTP method. Only POST supported
            void set_method(method)
            {
                // Method can only be a regular POST in our use case
                set_option(curl::CURLOPT_POST, 1);
            }

            // Set this HTTP request's content type.
            void set_content_type(const std::string& content_type)
            {
                curl::curl_slist* chunk = nullptr;
                chunk = curl::curl_slist_append(chunk, content_type.c_str());
                // FIXME: this is not exception-safe, set_option can throw
                set_option(curl::CURLOPT_HTTPHEADER, chunk);
                curl::curl_slist_free_all(chunk);
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
                    curl_.get(), option, std::forward<Args>(args)...);
                switch (retcode)
                {
                case curl::CURLE_OK:
                    return;

                case curl::CURLE_FAILED_INIT:
                    throw curl::curl_error{
                        "curl_easy_setopt: unsupported option"};

                default:
                    throw curl::curl_error{
                        "curl_easy_setopt: failed setting option"};
                };
            }

        private:
            curl::curl_ptr curl_;
        };

        // Makes a raw SOAP request.
        //
        // url: the URL of the server to talk to
        // username: username of user
        // password: the user's secret password, plain-text
        // domain: the user's domain
        // ews_request: the contents of the SOAP body (minus the body element)
        // soap_headers: any SOAP headers to add
        //
        // Returns a DOM wrapper around the response
        inline xml_document make_raw_soap_request(
            const std::string& url, const std::string& username,
            const std::string& password, const std::string& domain,
            const std::string& ews_request,
            const std::vector<std::string>& soap_headers)
        {
            http_web_request request{url};
            request.set_method(http_web_request::method::POST);
            request.set_content_type("text/xml;utf-8");

            ntlm_credentials creds{username, password, domain};
            request.set_credentials(creds);

            std::stringstream sstr;
            sstr << R"(
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
                sstr << "<soap:Header>\n";
                for (const auto& header : soap_headers)
                {
                    sstr << header;
                }
                sstr << "</soap:Header>\n";
            }

            sstr << "<soap:Body>\n";
            // Add the passed request
            sstr << ews_request;
            sstr << "</soap:Body>\n";
            sstr << "</soap:Envelope>\n";

            return xml_document{};
        }

        // Implementations

        void ntlm_credentials::certify(http_web_request* request) const
        {
            assert(request != nullptr);

            // CURLOPT_USERPWD: domain\username:password
            std::string login = domain_ + "\\" + username_ + "" + password_;
            request->set_option(curl::CURLOPT_USERPWD, login.c_str());
        }
    }
}
