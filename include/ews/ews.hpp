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

#include <curl/curl.h>

#include "rapidxml/rapidxml.hpp"

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
    // Forward declarations
    class item_id;

    enum class response_class { error, success, warning };

    enum class response_code { no_error }; // there are hundreds of these

    enum class base_shape { id_only, default_shape, all_properties };

    inline std::string base_shape_str(base_shape shape)
    {
        switch (shape)
        {
            case base_shape::id_only: return "IdOnly";
            case base_shape::default_shape: return "Default";
            case base_shape::all_properties: return "AllProperties";
            // TODO: make one base exception class
            default: throw std::runtime_error("Bad enum value");
        }
    }

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

        // Scope guard helper.
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
        // TODO: sure this can't be done easier within a header file?
        // We need better handling of static strings (URIs, XML node names,
        // values and stuff, see usage of rapidxml::internal::compare)
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
            using rapidxml::internal::compare;

            if (!response.is_soap_fault())
            {
                return;
            }
            const auto& doc = response.payload();
            auto elem = get_element_by_qname(doc,
                                             "ResponseCode",
                                             uri<>::microsoft::errors());
            EWS_ASSERT(elem &&
                "Expected SOAP faults to always have a <ResponseCode> element");
            if (compare(elem->value(),
                        elem->value_size(),
                        "ErrorSchemaValidation",
                        21))
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
# ifdef EWS_ENABLE_VERBOSE
                // Print HTTP headers to stderr
                set_option(CURLOPT_VERBOSE, 1L);
# endif
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

        // Helper function for parsing response messages.
        //
        // Code seems to be common for all response messages.
        //
        // Returns response class and response code and executes given function
        // for each item in the response's <Items> array.
        //
        // response: The HTTP response retrieved from the server.
        // response_message_element_name: One of GetItemResponseMessage,
        // CreateItemResponseMessage, DeleteItemResponseMessage, or
        // UpdateItemResponseMessage.
        // func: A callable that is invoked for each item in the response
        // message's <Items> array. A const rapidxml::xml_node& is passed to
        // that callable.
        template <typename Func>
        inline std::pair<response_class, response_code>
        for_each_item_in(http_response& response,
                         const char* response_message_element_name,
                         Func func)
        {
            using uri = internal::uri<>;
            using internal::get_element_by_qname;
            using rapidxml::internal::compare;

            const auto& doc = response.payload();
            auto elem = get_element_by_qname(doc, response_message_element_name,
                    uri::microsoft::messages());
            EWS_ASSERT(elem && "Expected element, got nullptr");

            // ResponseClass
            auto cls = response_class::success;
            auto response_class_attr = elem->first_attribute("ResponseClass");
            if (compare(response_class_attr->value(),
                        response_class_attr->value_size(),
                        "Error",
                        5))
            {
                cls = response_class::error;
            }
            else if (compare(response_class_attr->value(),
                             response_class_attr->value_size(),
                             "Warning",
                             7))
            {
                cls = response_class::warning;
            }

            // ResponseCode
            auto code = response_code::no_error;
            auto response_code_elem =
                elem->first_node_ns(uri::microsoft::messages(), "ResponseCode");
            EWS_ASSERT(response_code_elem && "Expected <ResponseCode> element");
            if (!compare(response_code_elem->value(),
                         response_code_elem->value_size(),
                         "NoError",
                         7))
            {
                // TODO: there are more possible response codes
                EWS_ASSERT(false && "Unexpected <ResponseCode> value");
            }

            // Items
            auto items_elem =
                elem->first_node_ns(uri::microsoft::messages(), "Items");
            EWS_ASSERT(response_code_elem && "Expected <Items> element");

            for (auto item_elem = items_elem->first_node(); item_elem;
                 item_elem = item_elem->next_sibling())
            {
                EWS_ASSERT(item_elem && "Expected an element, got nullptr");
                func(*item_elem);
            }

            return std::make_pair(cls, code);
        }

        // Base-class for various response messages
        //
        // The ItemType template parameter denotes the type of all items in the
        // returned array. The choice for a compile-time parameter has following
        // implications and restrictions:
        //
        // - Microsoft EWS allows for different types of items in the returned
        //   array. However, this implementation forces you to only issue
        //   requests that return only one type of item in a single response at
        //   a time.
        //
        // - You need to know the type of the item returned by a request
        //   up-front at compile time. Microsoft EWS would allow to deal with
        //   different types of items in a single response dynamically.
        template <typename ItemType>
        class response_message_base
        {
        public:
            typedef ItemType item_type;

            response_message_base(response_class cls,
                                  response_code code,
                                  std::vector<item_type> items)
                : items_(std::move(items)),
                  cls_(cls),
                  code_(code)
            {
            }

            response_class get_response_class() const EWS_NOEXCEPT
            {
                return cls_;
            }

            bool success() const EWS_NOEXCEPT
            {
                return get_response_class() == response_class::success;
            }

            response_code get_response_code() const EWS_NOEXCEPT
            {
                return code_;
            }

            const std::vector<item_type>& items() const EWS_NOEXCEPT
            {
                return items_;
            }

        private:
            std::vector<item_type> items_;
            response_class cls_;
            response_code code_;
        };

        class create_item_response_message final
            : public response_message_base<item_id>
        {
        public:
            // implemented below
            static create_item_response_message parse(http_response&);

        private:
            create_item_response_message(response_class cls,
                                         response_code code,
                                         std::vector<item_id> items)
                : response_message_base<item_id>(cls, code, std::move(items))
            {
            }
        };

        template <typename ItemType>
        class get_item_response_message final
            : public response_message_base<ItemType>
        {
        public:
            // implemented below
            static get_item_response_message parse(http_response&);

        private:
            get_item_response_message(response_class cls,
                                      response_code code,
                                      std::vector<ItemType> items)
                : response_message_base<ItemType>(cls, code, std::move(items))
            {
            }
        };

    } // internal

    // Function is not thread-safe; should be set-up when application is still
    // in single-threaded context. Calling this function more than once does no
    // harm.
    inline void set_up() { curl_global_init(CURL_GLOBAL_DEFAULT); }

    // Function is not thread-safe; you should call this function only when no
    // other thread is running (see libcurl(3) man-page or
    // http://curl.haxx.se/libcurl/c/libcurl.html)
    inline void tear_down() EWS_NOEXCEPT { curl_global_cleanup(); }

    // Contains the unique identifier and change key of an item in the Exchange
    // store.
    class item_id final
    {
    public:
        explicit item_id(std::string id)
            : id_(std::move(id)),
              change_key_()
        {}

        item_id(std::string id, std::string change_key)
            : id_(std::move(id)),
              change_key_(std::move(change_key))
        {}

        const std::string& id() const EWS_NOEXCEPT { return id_; }

        const std::string& change_key() const EWS_NOEXCEPT
        {
            return change_key_;
        }

        std::string to_xml(const char* xmlns=nullptr) const
        {
            if (xmlns)
            {
                return std::string("<") + xmlns + ":ItemId Id=\"" + id() +
                    "\" ChangeKey=\"" + change_key() + "\"/>";
            }
            else
            {
                return "<ItemId Id=\"" + id() +
                    "\" ChangeKey=\"" + change_key() + "\"/>";
            }
        }

        // Makes an item_id instance from an <ItemId> XML element
        static item_id from_xml_element(const rapidxml::xml_node<>& elem)
        {
            auto id_attr = elem.first_attribute("Id");
            EWS_ASSERT(id_attr && "Missing attribute Id in <ItemId>");
            auto id = std::string(id_attr->value(), id_attr->value_size());
            auto ckey_attr = elem.first_attribute("ChangeKey");
            EWS_ASSERT(ckey_attr && "Missing attribute ChangeKey in <ItemId>");
            auto ckey = std::string(ckey_attr->value(),
                    ckey_attr->value_size());
            return item_id(std::move(id), std::move(ckey));
        }

    private:
        // case-sensitive; therefore, comparisons between IDs must be
        // case-sensitive or binary
        std::string id_;

        // Identifies a specific version of an item.
        std::string change_key_;
    };

    // Note About Dates in EWS
    //
    // Microsoft EWS uses date and date/time string representations as described
    // in http://www.w3.org/TR/xmlschema-2/, notably xs:dateTime (or
    // http://www.w3.org/2001/XMLSchema:dateTime) and xs:date (also known as
    // http://www.w3.org/2001/XMLSchema:date).
    //
    // For example, the lexical representation of xs:date is
    //
    //     '-'? yyyy '-' mm '-' dd zzzzzz?
    //
    // whereas the z represents the timezone. Two examples of date strings are:
    // 2000-01-16Z and 1981-07-02 (the Z means Zulu time which is the same as
    // UTC). xs:dateTime is formatted accordingly, just with a time component;
    // you get the idea.
    //
    // This library does not interpret, parse, or in any way touch date nor
    // date/time strings in any circumstance. This library provides two classes,
    // date and date_type. Both classes, date and date_time, act solely as thin
    // wrapper to make the signatures of public API functions more type-rich and
    // easier to understand. Both types are implicitly convertible from
    // std::string.
    //
    // If your date or date/time strings are not formatted properly, Microsoft
    // EWS will likely give you a SOAP fault which this library transports to
    // you as an exception of type ews::soap_fault.

    // A date/time string wrapper class for xs:dateTime formatted strings.
    //
    // See Note About Dates in EWS above.
    class date_time final
    {
    public:
        date_time(std::string str) // intentionally not explicit
            : date_time_string_(std::move(str))
        {
        }

    private:
        std::string date_time_string_;
    };

    // A date string wrapper class for xs:date formatted strings.
    //
    // See Note About Dates in EWS above.
    class date final
    {
    public:
        date(std::string str) // intentionally not explicit
            : date_string_(std::move(str))
        {
        }

    private:
        std::string date_string_;
    };

    // Represents the actual body content of a message.
    //
    // This can be of type Best, HTML, or plain-text. See EWS XML elements
    // documentation on MSDN.
    class body final
    {
    public:
        body(const std::string& text) // intentionally not explicit
        {
            (void)text;
        }
    };

    // Represents a generic item in the Exchange store.
    //
    // Basically:
    //
    //      item
    //      ├── appointment
    //      ├── contact
    //      ├── message
    //      └── task
    class item
    {
    public:
        virtual ~item() = default;

    protected:
        // Sub-classes reimplement these functions
        virtual std::string create_item_request_string() = 0;

    private:
        friend class service;
    };

    // Represents a concrete task in the Exchange store.
    class task final : public item
    {
    public:
        task() = default;

        void set_subject(const std::string& str) { subject_ = str; }
        const std::string& get_subject() const EWS_NOEXCEPT { return subject_; }

        void set_body(const body&) {} // TODO: getter

        void set_owner(const std::string&) {}
        void set_start_date(const ews::date_time&) {}
        void set_due_date(const ews::date_time&) {}
        void set_reminder_enabled(bool) {}
        void set_reminder_due_by(const ews::date_time&) {}

        // Makes a task instance from a <Task> XML element
        static task from_xml_element(const rapidxml::xml_node<>& elem)
        {
            using uri = internal::uri<>;

            auto t = task();
            auto node = elem.first_node_ns(uri::microsoft::types(), "Subject");
            if (node)
            {
                t.set_subject(std::string(node->value(), node->value_size()));
            }

            return t;
        }

    private:
        std::string subject_;

        std::string create_item_request_string() override
        {
            return
              "<CreateItem " \
                  "xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " \
                  "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" >" \
              "<Items>" \
              "<t:Task>" \
              "<t:Subject>" + get_subject() + "</t:Subject>" \
              "</t:Task>" \
              "</Items>" \
              "</CreateItem>";
        }
    };

    // The service class contains all methods that can be performed on an
    // Exchange server.
    //
    // Will get a *huge* public interface over time, e.g.,
    //
    // - create_item
    // - find_conversation
    // - find_folder
    // - find_item
    // - find_people
    // - get_contact
    // - get_task
    //
    // and so on and so on.
    class service final
    {
    public:
        // FIXME: credentials are stored plain-text in memory
        //
        // That'll be bad. We wouldn't want random Joe at first-level support to
        // see plain-text passwords and user-names just because the process
        // crashed and some automatic mechanism sent a minidump over the wire.
        // What are our options? Security-by-obscurity: we could hash
        // credentials with a hash of the process-id or something.
        service(std::string server_uri, std::string domain,
                std::string username, std::string password)
            : server_uri_(std::move(server_uri)),
              domain_(std::move(domain)),
              username_(std::move(username)),
              password_(std::move(password)),
              server_version_("Exchange2013_SP1")
        {
        }

        // Gets a task from the Exchange store
        task get_task(item_id id)
        {
            return get_task(id, base_shape::all_properties);
        }

        // Creates local_item on the server and returns it's item_id.
        item_id create_item(item& the_item)
        {
            using internal::create_item_response_message;

            auto response = request(the_item.create_item_request_string());
            const auto response_message =
                create_item_response_message::parse(response);
            if (!response_message.success())
            {
                // TODO: throw? Also: code-dup; see get_task below
            }
            EWS_ASSERT(!response_message.items().empty()
                    && "Expected at least one item");
            return response_message.items().front();
        }

    private:
        std::string server_uri_;
        std::string domain_;
        std::string username_;
        std::string password_;
        std::string server_version_;

        // Helper for doing requests.
        //
        // Adds the right headers, credentials, and checks the response for
        // faults.
        internal::http_response request(const std::string& request_string)
        {
            // TODO: support multiple dialects depending on server version
            const auto soap_headers = std::vector<std::string> {
                "<t:RequestServerVersion Version=\"Exchange2013_SP1\"/>"
            };
            auto response = internal::make_raw_soap_request(server_uri_,
                                                            username_,
                                                            password_,
                                                            domain_,
                                                            request_string,
                                                            soap_headers);
            internal::raise_exception_if_soap_fault(response);
            return response;
        }

        task get_task(item_id id, base_shape shape)
        {
            using get_item_response_message =
                internal::get_item_response_message<task>;

            const std::string request_string =
              "<GetItem " \
                  "xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " \
                  "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" >" \
              "<ItemShape>" \
              "<t:BaseShape>" + base_shape_str(shape) + "</t:BaseShape>" \
              "</ItemShape>" \
              "<ItemIds>" + id.to_xml("t") + "</ItemIds>" \
              "</GetItem>";
            auto response = request(request_string);

           // std::cout << response.payload() << std::endl;

            const auto response_message =
                get_item_response_message::parse(response);
            if (!response_message.success())
            {
                // TODO: throw?
            }
            EWS_ASSERT(!response_message.items().empty()
                    && "Expected at least one item");
            return response_message.items().front();
        }
    };

    // Implementations

    namespace internal
    {
        inline void ntlm_credentials::certify(http_request* request) const
        {
            EWS_ASSERT(request != nullptr);

            // CURLOPT_USERPWD: domain\username:password
            std::string login = domain_ + "\\" + username_ + ":" + password_;
            request->set_option(CURLOPT_USERPWD, login.c_str());
            request->set_option(CURLOPT_HTTPAUTH, CURLAUTH_NTLM);
        }

        // FIXME: a CreateItemResponse can contain multiple ResponseMessages
        inline create_item_response_message
        create_item_response_message::parse(http_response& response)
        {
            auto item_ids = std::vector<item_id>();
            const auto result = for_each_item_in(
                    response,
                    "CreateItemResponseMessage",
                    [&item_ids](const rapidxml::xml_node<>& item_elem)
            {
                auto item_id_elem = item_elem.first_node();
                EWS_ASSERT(item_id_elem && "Expected <ItemId> element");
                item_ids.emplace_back(
                        item_id::from_xml_element(*item_id_elem));
            });
            return create_item_response_message(result.first,
                                                result.second,
                                                std::move(item_ids));
        }

        template <typename ItemType>
        inline
        get_item_response_message<ItemType>
        get_item_response_message<ItemType>::parse(http_response& response)
        {
            auto items = std::vector<ItemType>();
            const auto result = for_each_item_in(
                    response,
                    "GetItemResponseMessage",
                    [&items](const rapidxml::xml_node<>& item_elem)
            {
                items.emplace_back(ItemType::from_xml_element(item_elem));
            });
            return get_item_response_message(result.first,
                                             result.second,
                                             std::move(items));
        }
    }
}
