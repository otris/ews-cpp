#pragma once

#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <gtest/gtest.h>
#ifdef EWS_USE_BOOST_LIBRARY
# include <boost/filesystem.hpp>
#endif

#include <string>
#include <vector>
#include <memory>
#include <utility>

namespace tests
{
    // Global data used in tests; initialized at program-start
    struct Environment
    {
        static const ews::test::credentials& credentials()
        {
            static auto creds = ews::test::get_from_environment();
            return creds;
        }
    };

    // Per-test-case set-up and tear-down
    class BaseFixture : public ::testing::Test
    {
    protected:
        static void SetUpTestCase()
        {
            ews::set_up();
        }

        static void TearDownTestCase()
        {
            ews::tear_down();
        }
    };

    class FakeServiceFixture : public BaseFixture
    {
    public:
        struct http_request_mock final
        {
            struct storage
            {
                static storage& instance()
                {
                    thread_local storage inst;
                    return inst;
                }

                std::string request_string;
                std::vector<char> fake_response;
            };

            http_request_mock() = default;

            bool header_contains(const std::string& search_str) const
            {
                const auto& request_str = storage::instance().request_string;
                const auto header_end_idx = request_str.find("</soap:Header>");
                const auto search_str_idx = request_str.find(search_str);
                return     search_str_idx != std::string::npos
                        && search_str_idx < header_end_idx;
            }

            // Below same public interface as ews::internal::http_request class
            enum class method { POST };

            explicit http_request_mock(const std::string& url)
            {
                (void)url;
            }

            void set_method(method m)
            {
                (void)m;
            }

            void set_content_type(const std::string& content_type)
            {
                (void)content_type;
            }

            void set_credentials(const ews::internal::credentials& creds)
            {
                (void)creds;
            }

            template <typename... Args> void set_option(CURLoption, Args...) {}

            ews::internal::http_response send(const std::string& request)
            {
                auto& s = storage::instance();
                s.request_string = request;
                auto response_bytes = s.fake_response;
                return ews::internal::http_response(200,
                                                    std::move(response_bytes));
            }
        };

        virtual ~FakeServiceFixture() = default;

        ews::basic_service<http_request_mock>& service()
        {
            return *service_ptr_;
        }

        http_request_mock get_last_request()
        {
            return http_request_mock();
        }

        void set_next_fake_response(const char* str)
        {
            auto& storage = http_request_mock::storage::instance();
            storage.fake_response =
                std::vector<char>(str, str + std::strlen(str));
            storage.fake_response.push_back('\0');
        }

    protected:
        virtual void SetUp()
        {
            BaseFixture::SetUp();
#ifdef EWS_HAS_MAKE_UNIQUE
            service_ptr_ = std::make_unique<
                                ews::basic_service<http_request_mock>>(
                                    "https://example.com/ews/Exchange.asmx",
                                    "FAKEDOMAIN",
                                    "fakeuser",
                                    "fakepassword");
#else
            service_ptr_ = std::unique_ptr<ews::basic_service<http_request_mock>>(
                                new ews::basic_service<http_request_mock>(
                                    "https://example.com/ews/Exchange.asmx",
                                    "FAKEDOMAIN",
                                    "fakeuser",
                                    "fakepassword"));
#endif
        }

        virtual void TearDown()
        {
            service_ptr_.reset();
            BaseFixture::TearDown();
        }

    private:
        std::unique_ptr<ews::basic_service<http_request_mock>> service_ptr_;
    };

    // Set-up and tear down a service object
    class ServiceFixture : public BaseFixture
    {
    public:
        virtual ~ServiceFixture() = default;

        ews::service& service()
        {
            return *service_ptr_;
        }

    protected:
        virtual void SetUp()
        {
            BaseFixture::SetUp();
            const auto& creds = Environment::credentials();
#ifdef EWS_HAS_MAKE_UNIQUE
            service_ptr_ = std::make_unique<ews::service>(creds.server_uri,
                                                          creds.domain,
                                                          creds.username,
                                                          creds.password);
#else
            service_ptr_ = std::unique_ptr<ews::service>(
                                         new ews::service(creds.server_uri,
                                                          creds.domain,
                                                          creds.username,
                                                          creds.password));
#endif
        }

        virtual void TearDown()
        {
            service_ptr_.reset();
            BaseFixture::TearDown();
        }

    private:
        std::unique_ptr<ews::service> service_ptr_;
    };

    // Create and remove a task on the server
    class TaskTest : public ServiceFixture
    {
    public:
        void SetUp()
        {
            ServiceFixture::SetUp();

            task_.set_subject("Get some milk");
            task_.set_body(ews::body("Get some milk from the store"));
            task_.set_start_date(ews::date_time("2015-06-17T19:00:00Z"));
            task_.set_due_date(ews::date_time("2015-06-17T19:30:00Z"));
            const auto item_id = service().create_item(task_);
            task_ = service().get_task(item_id);
        }

        void TearDown()
        {
            service().delete_task(
                    std::move(task_),
                    ews::delete_type::hard_delete,
                    ews::affected_task_occurrences::all_occurrences);

            ServiceFixture::TearDown();
        }

        ews::task& test_task() { return task_; }

    private:
        ews::task task_;
    };

    // Create and remove a contact on the server
    class ContactTest : public ServiceFixture
    {
    public:
        void SetUp()
        {
            ServiceFixture::SetUp();

            contact_.set_given_name("Minnie");
            contact_.set_surname("Mouse");
            const auto item_id = service().create_item(contact_);
            contact_ = service().get_contact(item_id);
        }

        void TearDown()
        {
            service().delete_contact(std::move(contact_));

            ServiceFixture::TearDown();
        }

        ews::contact& test_contact() { return contact_; }

    private:
        ews::contact contact_;
    };

    struct MessageTest : public ServiceFixture
    {
    };

    class AttachmentTest : public ServiceFixture
    {
    public:
        void SetUp()
        {
            ServiceFixture::SetUp();

            auto msg = ews::message();
            msg.set_subject("Honorable Minister of Finance - Release Funds");
            std::vector<ews::email_address> recipients{
                ews::email_address("udom.emmanuel@zenith-bank.com.ng")
            };
            msg.set_to_recipients(recipients);
            auto item_id = service().create_item(
                    msg,
                    ews::message_disposition::save_only);
            message_ = service().get_message(item_id);
        }

        void TearDown()
        {
            service().delete_message(std::move(message_));

            ServiceFixture::TearDown();
        }

        ews::message& test_message() { return message_; }

    private:
        ews::message message_;
    };

#ifdef EWS_USE_BOOST_LIBRARY
    class FileAttachmentTest : public BaseFixture
    {
    public:
        FileAttachmentTest()
            : assets_dir_("/home/bkircher/src/ews-cpp/tests/assets")
        {
        }

        const boost::filesystem::path& assets_dir() const
        {
            return assets_dir_;
        }

    private:
        boost::filesystem::path assets_dir_;
    };
#endif // EWS_USE_BOOST_LIBRARY
}
