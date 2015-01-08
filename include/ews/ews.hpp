#pragma once

#include <exception>
#include <stdexcept>
#include <string>
#include <vector>
#include <sstream>
#include <algorithm>
#include <functional>
#include <utility>
#include <memory>
#ifndef _MSC_VER
#include <type_traits>
#endif
#include <cstddef>
#include <cstring>

#include "rapidxml/rapidxml.hpp"

#include <curl/curl.h>

// Macro for verifying expressions at run-time. Calls assert() with 'expr'.
// Allows turning assertions off, even if -DNDEBUG wasn't given at
// compile-time.  This macro does nothing unless EWS_ENABLE_ASSERTS was defined
// during compilation.
#define EWS_ASSERT(expr) ((void)0)
#ifndef NDEBUG
# ifdef EWS_ENABLE_ASSERTS
#  include <cassert>
#  undef EWS_ASSERT
#  define EWS_ASSERT(expr) assert(expr)
# endif
#endif // !NDEBUG

// Visual Studio 12 does not support noexcept specifier.
#ifndef _MSC_VER
# define EWS_NOEXCEPT noexcept
#else
# define EWS_NOEXCEPT 
#endif

namespace ews
{
    // A SOAP fault occurred due to a bad request
    class soap_fault : public std::runtime_error
    {
    public:
        explicit soap_fault(const std::string& what)
            : std::runtime_error(what)
        {
        }
        explicit soap_fault(const char* what) : std::runtime_error(what) {}
    };

    // A SOAP fault that is raised when we sent invalid XML.
    //
    // This is an internal error and indicates a bug in this library, thus
    // should never happen.
    //
    // Note: because this exception is due to a SOAP fault (sometimes recognized
    // before any server-side XML parsing finished) any included failure message
    // is likely not localized according to any MailboxCulture SOAP header.
    class schema_validation_error final : public soap_fault
    {
    public:
        schema_validation_error(unsigned long line_number,
                                unsigned long line_position,
                                std::string violation)
            : soap_fault("The request failed schema validation"),
              violation_(std::move(violation)),
              line_pos_(line_position),
              line_no_(line_number)
        {
        }

        // Line number in request string where the error was found
        unsigned long line_number() const EWS_NOEXCEPT { return line_no_; }

        // Column number in request string where the error was found
        unsigned long line_position() const EWS_NOEXCEPT { return line_pos_; }

        // A more detailed explanation of what went wrong
        const std::string& violation() const EWS_NOEXCEPT { return violation_; }

    private:
        std::string violation_;
        unsigned long line_pos_;
        unsigned long line_no_;
    };

    namespace internal
    {
        // Exception for libcurl related runtime errors
        class curl_error final : public std::runtime_error
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
        inline curl_error make_curl_error(const std::string& msg, CURLcode rescode)
        {
            std::string reason{ curl_easy_strerror(rescode) };
#ifdef NDEBUG
            (void)msg;
            return curl_error(reason);
#else
            return curl_error(msg + ": \'" + reason + "\'");
#endif
        }

        // RAII helper class for CURL* handles.
        class curl_ptr final
        {
        public:
            curl_ptr() : handle_{ curl_easy_init() }
            {
                if (!handle_)
                {
                    throw curl_error{ "Could not start libcurl session" };
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
        class curl_string_list final
        {
        public:
            curl_string_list() EWS_NOEXCEPT : slist_{ nullptr } {}

            ~curl_string_list() { curl_slist_free_all(slist_); }

            void append(const char* str) EWS_NOEXCEPT
            {
                slist_ = curl_slist_append(slist_, str);
            }

            curl_slist* get() const EWS_NOEXCEPT { return slist_; }

        private:
            curl_slist* slist_;
        };

        // Obligatory scope guard helper.
        class on_scope_exit final
        {
        public:
            template <typename Function>
            on_scope_exit(Function destructor_function) try
                : func_(std::move(destructor_function))
            {}
            catch (std::exception&)
            {
                destructor_function();
            }

            on_scope_exit() = delete;
            on_scope_exit(const on_scope_exit&) = delete;
            on_scope_exit& operator=(const on_scope_exit&) = delete;

            ~on_scope_exit()
            {
                if (func_)
                {
                    try
                    {
                        func_();
                    }
                    catch (std::exception&)
                    {
                        // Swallow, abort(3) if suitable
                        EWS_ASSERT(false);
                    }
                }
            }

            void release() { func_ = nullptr; }

        private:
            std::function<void(void)> func_;
        };

#ifndef _MSC_VER
        // <type_traits> is broken in Visual Studio 12, disregard
        static_assert(!std::is_copy_constructible<on_scope_exit>::value, "");
        static_assert(!std::is_copy_assignable<on_scope_exit>::value, "");
        static_assert(!std::is_default_constructible<on_scope_exit>::value, "");
#endif

        // Raised when a response from a server could not be parsed.
        //
        // TODO: remove parse_error, use rapidxml's
        class parse_error final : public std::runtime_error
        {
        public:
            explicit parse_error(const std::string& what)
                : std::runtime_error(what)
            {
            }
            explicit parse_error(const char* what) : std::runtime_error(what) {}
        };

        // Forward declarations
        class http_request;

        // String constants
        // TODO: sure this can't be done easier?
        template <typename Ch = char> struct uri
        {
            struct microsoft
            {
                enum { errors_size = 58U };
                static const Ch* errors()
                {
                    static const Ch* val =
                        "http://schemas.microsoft.com/exchange/services/2006/errors";
                    return val;
                }
                enum { types_size = 57U };
                static const Ch* types()
                {
                    static const Ch* const val =
                        "http://schemas.microsoft.com/exchange/services/2006/types";
                    return val;
                }
                enum { messages_size = 60U };
                static const Ch* messages()
                {
                    static const Ch* const val =
                        "http://schemas.microsoft.com/exchange/services/2006/messages";
                    return val;
                }
            };
            struct soapxml
            {
                enum { envelope_size = 41U };
                static const Ch* envelope()
                {
                    static const Ch* const val =
                        "http://schemas.xmlsoap.org/soap/envelope/";
                    return val;
                }
            };
        };

        // This ought to be a DOM wrapper; usually around a web response
        //
        // This class basically wraps rapidxml::xml_document because the parsed
        // data must persist for the lifetime of the rapidxml::xml_document.
        class http_response final
        {
        public:
            http_response(long code, std::vector<char>&& data)
                : data_(std::move(data)), doc_(), code_(code), parsed_(false)
            {
                EWS_ASSERT(!data_.empty());
            }

            http_response(const http_response&) = delete;
            http_response& operator=(const http_response&) = delete;

            http_response(http_response&& other)
                : data_(std::move(other.data_)),
                  doc_(std::move(other.doc_)),
                  code_(std::move(other.code_)),
                  parsed_(std::move(other.parsed_))
            {
                other.code_ = 0U;
            }

            http_response& operator=(http_response&& rhs)
            {
                if (&rhs != this)
                {
                    data_ = std::move(rhs.data_);
                    doc_ = std::move(rhs.doc_);
                    code_ = std::move(rhs.code_);
                    parsed_ = std::move(rhs.parsed_);
                }
                return *this;
            }

            // Returns the SOAP payload in this response.
            //
            // Parses the payload (if it hasn't already) and returns it as
            // xml_document.
            //
            // Note: we are using a mutable temporary buffer internally because
            // we are using RapidXml in destructive mode (the parser modifies
            // source text during the parsing process). Hence, we need to make
            // sure that parsing is done only once!
            const rapidxml::xml_document<char>& payload()
            {
                if (!parsed_)
                {
                    on_scope_exit set_parsed([&]()
                    {
                        parsed_ = true;
                    });
                    parse();
                }
                return doc_;
            }

            // Returns the response code of the HTTP request.
            long code() const EWS_NOEXCEPT
            {
                return code_;
            }

            // Returns whether the response is a SOAP fault.
            //
            // This means the server responded with status code 500 and
            // indicates that the entire request failed (not just a normal EWS
            // error). This can happen e.g. when the request we sent was not
            // schema compliant.
            bool is_soap_fault() const EWS_NOEXCEPT
            {
                return code() == 500U;
            }

            // Returns whether the HTTP respone code is 200 (OK).
            bool ok() const EWS_NOEXCEPT
            {
                return code() == 200U;
            }

        private:
            // Here we handle the server's response. We load the SOAP payload
            // from response into the xml_document.
            void parse()
            {
                try
                {
                    static const int flags = 0;
                    doc_.parse<flags>(&data_[0]);
                }
                catch (rapidxml::parse_error& e)
                {
                    // Swallow and erase type
                    throw parse_error(e.what());
                }
            }

            std::vector<char> data_;
            rapidxml::xml_document<char> doc_;
            long code_;
            bool parsed_;
        };

        // Traverse elements, depth first, beginning with given node.
        //
        // Applies given function to every element during traversal, stopping as
        // soon as that function returns true.
        template <typename Function>
        inline void traverse_elements(const rapidxml::xml_node<>& node,
                                      Function func)
        {
            for (auto child = node.first_node(); child != nullptr;
                 child = child->next_sibling())
            {
                traverse_elements(*child, func);
                if (child->type() == rapidxml::node_element)
                {
                    if (func(*child))
                    {
                        return;
                    }
                }
            }
        }

        // Select element by qualified name, nullptr if there is no such element
        // TODO: what if namespace_uri is empty, can we do that cheaper?
        inline rapidxml::xml_node<>*
        get_element_by_qname(const rapidxml::xml_node<>& node,
                             const char* local_name,
                             const char* namespace_uri)
        {
            EWS_ASSERT(local_name && namespace_uri);

            rapidxml::xml_node<>* element = nullptr;
            const auto local_name_len = std::strlen(local_name);
            const auto namespace_uri_len = std::strlen(namespace_uri);
            traverse_elements(node, [&](rapidxml::xml_node<>& elem) -> bool
            {
                using rapidxml::internal::compare;

                if (compare(elem.namespace_uri(),
                            elem.namespace_uri_size(),
                            namespace_uri,
                            namespace_uri_len)
                    &&
                    compare(elem.local_name(),
                            elem.local_name_size(),
                            local_name,
                            local_name_len))
                {
                    element = std::addressof(elem);
                    return true;
                }
                return false;
            });
            return element;
        }

        // Does nothing if given response is not a SOAP fault
        inline void raise_exception_if_soap_fault(http_response& response)
        {
            if (!response.is_soap_fault())
            {
                return;
            }
            using rapidxml::internal::compare;
            const auto& doc = response.payload();
            auto elem = get_element_by_qname(doc,
                                             "ResponseCode",
                                             uri<>::microsoft::errors());
            EWS_ASSERT(elem &&
                "Expected SOAP faults to always have a <ResponseCode> element");
            if (compare(elem->value(),
                        elem->value_size(),
                        "ErrorSchemaValidation",
                        std::strlen("ErrorSchemaValidation")))
            {
                // Get some more helpful details
                elem = get_element_by_qname(doc,
                                            "LineNumber",
                                            uri<>::microsoft::types());
                EWS_ASSERT(elem &&
                        "Expected <LineNumber> element in response");
                const auto line_number =
                    std::stoul(std::string(elem->value(), elem->value_size()));

                elem = get_element_by_qname(doc,
                                            "LinePosition",
                                            uri<>::microsoft::types());
                EWS_ASSERT(elem &&
                        "Expected <LinePosition> element in response");
                const auto line_position =
                    std::stoul(std::string(elem->value(), elem->value_size()));

                elem = get_element_by_qname(doc,
                                            "Violation",
                                            uri<>::microsoft::types());
                EWS_ASSERT(elem &&
                        "Expected <Violation> element in response");
                throw schema_validation_error(
                        line_number,
                        line_position,
                        std::string(elem->value(), elem->value_size()));
            }
            else
            {
                elem = get_element_by_qname(doc, "faultstring", "");
                EWS_ASSERT(elem &&
                        "Expected <faultstring> element in response");
                throw soap_fault(elem->value());
            }
        }

        class credentials
        {
        public:
            virtual ~credentials() = default;
            virtual void certify(http_request*) const = 0;
        };

        class ntlm_credentials final : public credentials
        {
        public:
            ntlm_credentials(std::string username, std::string password,
                             std::string domain)
                : username_(std::move(username)),
                  password_(std::move(password)),
                  domain_(std::move(domain))
            {
            }

        private:
            void certify(http_request*) const override; // implemented below

            std::string username_;
            std::string password_;
            std::string domain_;
        };

        class http_request final
        {
        public:
            enum class method
            {
                POST
            };

            // Create a new HTTP request to the given URL.
            explicit http_request(const std::string& url)
            {
                set_option(CURLOPT_URL, url.c_str());
            }

            // Set the HTTP method (only POST supported).
            void set_method(method)
            {
                // Method can only be a regular POST in our use case
                set_option(CURLOPT_POST, 1);
            }

            // Set this HTTP request's content type.
            void set_content_type(const std::string& content_type)
            {
                const std::string str = "Content-Type: " + content_type;
                headers_.append(str.c_str());
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
            void set_option(CURLoption option, Args... args)
            {
                auto retcode = curl_easy_setopt(
                    handle_.get(), option, std::forward<Args>(args)...);
                switch (retcode)
                {
                case CURLE_OK:
                    return;

                case CURLE_FAILED_INIT:
                {
                    throw make_curl_error(
                        "curl_easy_setopt: unsupported option", retcode);
                }

                default:
                {
                    throw make_curl_error(
                        "curl_easy_setopt: failed setting option", retcode);
                }
                };
            }

            // Perform the HTTP request and returns the response. This function
            // blocks until the complete response is received or a timeout is
            // reached. Throws curl_error if operation could not be completed.
            //
            // request: The complete request string; you must make sure that
            // the data is encoded the way you want the server to receive it.
            http_response send(const std::string& request)
            {
                auto callback =
                    [](char* ptr, std::size_t size, std::size_t nmemb,
                       void* userdata) -> std::size_t
                {
                    std::vector<char>* buf
                        = reinterpret_cast<std::vector<char>*>(userdata);
                    const auto realsize = size * nmemb;
                    try
                    {
                        buf->reserve(realsize + 1); // plus 0-terminus
                    }
                    catch (std::bad_alloc&)
                    {
                        // Out of memory, indicate error to libcurl
                        return 0U;
                    }
                    std::copy(ptr, ptr + realsize, std::back_inserter(*buf));
                    return realsize;
                };

#ifndef NDEBUG
                // Print HTTP headers to stderr
                set_option(CURLOPT_VERBOSE, 1L);
#endif

                // Set complete request string for HTTP POST method; note: no
                // encoding here
                set_option(CURLOPT_POSTFIELDS, request.c_str());
                set_option(CURLOPT_POSTFIELDSIZE, request.length());

                // Finally, set HTTP headers. We do this as last action here
                // because we want to overwrite implicitly set header lines due
                // to the options set above with our own header lines
                set_option(CURLOPT_HTTPHEADER, headers_.get());

                std::vector<char> response_data;
                set_option(
                    CURLOPT_WRITEFUNCTION,
                    static_cast<std::size_t (*)(char*, std::size_t, std::size_t,
                                                void*)>(callback));
                set_option(CURLOPT_WRITEDATA,
                        std::addressof(response_data));

#ifndef NDEBUG
                // Turn-off verification of the server's authenticity
                set_option(CURLOPT_SSL_VERIFYPEER, 0);
#endif

                auto retcode = curl_easy_perform(handle_.get());
                if (retcode != 0)
                {
                    throw make_curl_error("curl_easy_perform", retcode);
                }
                long response_code = 0U;
                curl_easy_getinfo(handle_.get(),
                    CURLINFO_RESPONSE_CODE, &response_code);
                response_data.emplace_back('\0');
                return http_response(response_code, std::move(response_data));
            }

        private:
            curl_ptr handle_;
            curl_string_list headers_;
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
        // Returns the response.
        inline http_response make_raw_soap_request(
            const std::string& url, const std::string& username,
            const std::string& password, const std::string& domain,
            const std::string& soap_body,
            const std::vector<std::string>& soap_headers)
        {
            http_request request{url};
            request.set_method(http_request::method::POST);
            request.set_content_type("text/xml; charset=utf-8");

            ntlm_credentials creds{username, password, domain};
            request.set_credentials(creds);

            std::stringstream request_stream;
            request_stream <<
R"(<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
>)";

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

            return request.send(request_stream.str());
        }

        // Implementations

        void ntlm_credentials::certify(http_request* request) const
        {
            EWS_ASSERT(request != nullptr);

            // CURLOPT_USERPWD: domain\username:password
            std::string login = domain_ + "\\" + username_ + ":" + password_;
            request->set_option(CURLOPT_USERPWD, login.c_str());
            request->set_option(CURLOPT_HTTPAUTH, CURLAUTH_NTLM);
        }

    } // internal

    // Function is not thread-safe; should be set-up when application is still
    // in single-threaded context. Calling this function more than once does no
    // harm.
    inline void set_up() { curl_global_init(CURL_GLOBAL_DEFAULT); }

    // Function is not thread-safe; you should call this function only when no
    // other thread is running (see libcurl(3) man-page or
    // http://curl.haxx.se/libcurl/c/libcurl.html)
    inline void tear_down() EWS_NOEXCEPT { curl_global_cleanup(); }
}
