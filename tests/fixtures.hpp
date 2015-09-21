#pragma once

#include <string>
#include <vector>
#include <algorithm>
#include <memory>
#include <utility>
#include <stdexcept>

#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <ews/rapidxml/rapidxml.hpp>
#include <gtest/gtest.h>
#ifdef EWS_USE_BOOST_LIBRARY
# include <boost/filesystem.hpp>
#endif

namespace tests
{
    struct environment
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
#ifdef EWS_HAS_THREAD_LOCAL_STORAGE
                    thread_local storage inst;
#else
                    static storage inst;
#endif
                    return inst;
                }

                std::string request_string;
                std::vector<char> fake_response;
            };

#ifdef EWS_HAS_DEFAULT_AND_DELETE
            http_request_mock() = default;
#else
            http_request_mock() {}
#endif

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

#ifdef EWS_HAS_VARIADIC_TEMPLATES
            template <typename... Args> void set_option(CURLoption, Args...) {}
#else
            template <typename T1>
            void set_option(CURLoption option, T1) {}

            template <typename T1, typename T2>
            void set_option(CURLoption option, T1, T2) {}
#endif

            ews::internal::http_response send(const std::string& request)
            {
                auto& s = storage::instance();
                s.request_string = request;
                auto response_bytes = s.fake_response;
                return ews::internal::http_response(200,
                                                    std::move(response_bytes));
            }
        };

#ifdef EWS_HAS_DEFAULT_AND_DELETE
        virtual ~FakeServiceFixture() = default;
#else
        virtual ~FakeServiceFixture() {}
#endif

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
            service_ptr_ =
                std::unique_ptr<ews::basic_service<http_request_mock>>(
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
#ifdef EWS_HAS_DEFAULT_AN_DELETE
        virtual ~ServiceFixture() = default;
#else
        virtual ~ServiceFixture() {}
#endif

        ews::service& service()
        {
            if (!service_ptr_)
            {
                throw std::runtime_error(
                        "Cannot access service: no instance created");
            }
            return *service_ptr_;
        }

    protected:
        virtual void SetUp()
        {
            BaseFixture::SetUp();
            const auto& creds = environment::credentials();
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
            std::vector<ews::email_address> recipients;
            recipients.push_back(
                ews::email_address("udom.emmanuel@zenith-bank.com.ng")
            );
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
    class FileAttachmentTest : public AttachmentTest
    {
    public:
        FileAttachmentTest()
            : assetsdir_(ews::test::global_data::instance().assets_dir)
        {
        }

        void SetUp()
        {
            AttachmentTest::SetUp();

            olddir_ = boost::filesystem::current_path();
            workingdir_ = boost::filesystem::unique_path(
                        boost::filesystem::temp_directory_path() /
                        "%%%%-%%%%-%%%%-%%%%");
            ASSERT_TRUE(boost::filesystem::create_directory(workingdir_))
                << "Unable to create temporary working directory";
            boost::filesystem::current_path(workingdir_);
        }

        void TearDown()
        {
            EXPECT_TRUE(boost::filesystem::is_empty(workingdir_))
                << "Temporary directory not empty on TearDown";
            boost::filesystem::remove_all(workingdir_);
            boost::filesystem::current_path(olddir_);

            AttachmentTest::TearDown();
        }

        const boost::filesystem::path& assets_dir() const
        {
            return assetsdir_;
        }

        const boost::filesystem::path& cwd() const
        {
            return workingdir_;
        }

    private:
        boost::filesystem::path assetsdir_;
        boost::filesystem::path olddir_;
        boost::filesystem::path workingdir_;
    };
#endif // EWS_USE_BOOST_LIBRARY

    inline ews::task make_fake_task(const char* xml=nullptr)
    {
        typedef rapidxml::xml_document<> xml_document;

        if (!xml)
        {
            xml =
                "<t:Task\n"
                "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n"
                "    <t:ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>\n"
                "    <t:ParentFolderId Id=\"qwertz\" ChangeKey=\"ztrewq\"/>\n"
                "    <t:ItemClass>IPM.Task</t:ItemClass>\n"
                "    <t:Subject>Write poem</t:Subject>\n"
                "    <t:Sensitivity>Confidential</t:Sensitivity>\n"
                "    <t:Body BodyType=\"Text\" IsTruncated=\"false\"/>\n"
                "    <t:DateTimeReceived>2015-02-09T13:00:11Z</t:DateTimeReceived>\n"
                "    <t:Size>962</t:Size>\n"
                "    <t:Importance>Normal</t:Importance>\n"
                "    <t:IsSubmitted>false</t:IsSubmitted>\n"
                "    <t:IsDraft>false</t:IsDraft>\n"
                "    <t:IsFromMe>false</t:IsFromMe>\n"
                "    <t:IsResend>false</t:IsResend>\n"
                "    <t:IsUnmodified>false</t:IsUnmodified>\n"
                "    <t:DateTimeSent>2015-02-09T13:00:11Z</t:DateTimeSent>\n"
                "    <t:DateTimeCreated>2015-02-09T13:00:11Z</t:DateTimeCreated>\n"
                "    <t:DisplayCc/>\n"
                "    <t:DisplayTo/>\n"
                "    <t:HasAttachments>false</t:HasAttachments>\n"
                "    <t:Culture>en-US</t:Culture>\n"
                "    <t:EffectiveRights>\n"
                "            <t:CreateAssociated>false</t:CreateAssociated>\n"
                "            <t:CreateContents>false</t:CreateContents>\n"
                "            <t:CreateHierarchy>false</t:CreateHierarchy>\n"
                "            <t:Delete>true</t:Delete>\n"
                "            <t:Modify>true</t:Modify>\n"
                "            <t:Read>true</t:Read>\n"
                "            <t:ViewPrivateItems>true</t:ViewPrivateItems>\n"
                "    </t:EffectiveRights>\n"
                "    <t:LastModifiedName>Kwaltz</t:LastModifiedName>\n"
                "    <t:LastModifiedTime>2015-02-09T13:00:11Z</t:LastModifiedTime>\n"
                "    <t:IsAssociated>false</t:IsAssociated>\n"
                "    <t:Flag>\n"
                "            <t:FlagStatus>NotFlagged</t:FlagStatus>\n"
                "    </t:Flag>\n"
                "    <t:InstanceKey>AQAAAAAAARMBAAAAG4AqWQAAAAA=</t:InstanceKey>\n"
                "    <t:EntityExtractionResult/>\n"
                "    <t:ChangeCount>1</t:ChangeCount>\n"
                "    <t:IsComplete>false</t:IsComplete>\n"
                "    <t:IsRecurring>false</t:IsRecurring>\n"
                "    <t:PercentComplete>0</t:PercentComplete>\n"
                "    <t:Status>NotStarted</t:Status>\n"
                "    <t:StatusDescription>Not Started</t:StatusDescription>\n"
                "</t:Task>";
        }

        std::vector<char> buf;
        std::copy(xml, xml + std::strlen(xml), std::back_inserter(buf));
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = doc.first_node();
        return ews::task::from_xml_element(*node);
    }
}

// vim:et ts=4 sw=4 noic cc=80
