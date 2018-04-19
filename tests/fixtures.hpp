
//   Copyright 2016 otris software AG
//
//   Licensed under the Apache License, Version 2.0 (the "License");
//   you may not use this file except in compliance with the License.
//   You may obtain a copy of the License at
//
//       http://www.apache.org/licenses/LICENSE-2.0
//
//   Unless required by applicable law or agreed to in writing, software
//   distributed under the License is distributed on an "AS IS" BASIS,
//   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//   See the License for the specific language governing permissions and
//   limitations under the License.
//
//   This project is hosted at https://github.com/otris

#pragma once

#include <algorithm>
#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <ews/rapidxml/rapidxml.hpp>
#include <memory>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

#ifdef EWS_HAS_FILESYSTEM_HEADER
#    include <cstdio>
#    include <filesystem>
#    include <fstream>
#    include <iostream>
#    include <iterator>
#endif

#include <gtest/gtest.h>

namespace tests
{
// Check if cont contains the element val
template <typename ContainerType, typename ValueType>
inline bool contains(const ContainerType& cont, const ValueType& val)
{
    return std::find(begin(cont), end(cont), val) != end(cont);
}

// Check if cont contains an element for which pred evaluates to true
template <typename ContainerType, typename Predicate>
inline bool contains_if(const ContainerType& cont, Predicate pred)
{
    return std::find_if(begin(cont), end(cont), pred) != end(cont);
}

#ifdef EWS_HAS_FILESYSTEM_HEADER
// Read file contents into a buffer
inline std::vector<char> read_file(const std::filesystem::path& path)
{
    std::ifstream ifstr(path.string(), std::ifstream::in | std::ios::binary);
    if (!ifstr.is_open())
    {
        throw std::runtime_error("Could not open file for reading: " +
                                 path.string());
    }

    ifstr.unsetf(std::ios::skipws);

    ifstr.seekg(0, std::ios::end);
    const auto file_size = ifstr.tellg();
    ifstr.seekg(0, std::ios::beg);

    auto contents = std::vector<char>();
    contents.reserve(file_size);

    contents.insert(begin(contents),
                    std::istream_iterator<unsigned char>(ifstr),
                    std::istream_iterator<unsigned char>());
    ifstr.close();
    contents.push_back('\0');
    return contents;
}
#endif // EWS_HAS_FILESYSTEM_HEADER

// Allows you to test requests w/o sending anything to the server. See also
// FakeServiceFixture.
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
        std::string url;
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
        return search_str_idx != std::string::npos &&
               search_str_idx < header_end_idx;
    }

    const std::string& request_string() const
    {
        return storage::instance().request_string;
    }

    // Below same public interface as ews::internal::http_request class
    enum class method
    {
        POST
    };

    explicit http_request_mock(const std::string& url)
    {
        auto& s = storage::instance();
        s.url = url;
    }

    void set_method(method) {}

    void set_content_type(const std::string&) {}

    void set_content_length(size_t) {}

    void set_credentials(const ews::internal::credentials&) {}

#ifdef EWS_HAS_VARIADIC_TEMPLATES
    template <typename... Args> void set_option(CURLoption, Args...) {}
#else
    template <typename T1> void set_option(CURLoption option, T1) {}

    template <typename T1, typename T2>
    void set_option(CURLoption option, T1, T2)
    {
    }
#endif

    ews::internal::http_response send(const std::string& request)
    {
        auto& s = storage::instance();
        s.request_string = request;
        auto response_bytes = s.fake_response;
        return ews::internal::http_response(200, std::move(response_bytes));
    }
};

// Per-test-case set-up and tear-down
class BaseFixture : public testing::Test
{
public:
    BaseFixture() : assets_(ews::test::global_data::instance().assets_dir) {}

    const std::string& assets() const { return assets_; }

protected:
    static void SetUpTestCase() { ews::set_up(); }

    static void TearDownTestCase() { ews::tear_down(); }

private:
    std::string assets_;
};

class FakeServiceFixture : public BaseFixture
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    virtual ~FakeServiceFixture() = default;
#else
    virtual ~FakeServiceFixture() {}
#endif

    ews::basic_service<http_request_mock>& service() { return *service_ptr_; }

    http_request_mock get_last_request() { return http_request_mock(); }

    void set_next_fake_response(const char* str)
    {
        auto& storage = http_request_mock::storage::instance();
        storage.fake_response = std::vector<char>(str, str + strlen(str));
        storage.fake_response.push_back('\0');
    }

    void set_next_fake_response(std::vector<char>&& buffer)
    {
        auto& storage = http_request_mock::storage::instance();
        storage.fake_response = std::move(buffer);
    }

protected:
    virtual void SetUp()
    {
        BaseFixture::SetUp();
#ifdef EWS_HAS_MAKE_UNIQUE
        service_ptr_ = std::make_unique<ews::basic_service<http_request_mock>>(
            "https://example.com/ews/Exchange.asmx", "FAKEDOMAIN", "fakeuser",
            "fakepassword");
#else
        service_ptr_ = std::unique_ptr<ews::basic_service<http_request_mock>>(
            new ews::basic_service<http_request_mock>(
                "https://example.com/ews/Exchange.asmx", "FAKEDOMAIN",
                "fakeuser", "fakepassword"));
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
class ServiceMixin
{
public:
    ServiceMixin()
    {
        const auto& env = ews::test::global_data::instance().env;
#ifdef EWS_HAS_MAKE_UNIQUE
        service_ptr_ = std::make_unique<ews::service>(
            env.server_uri, env.domain, env.username, env.password);
#else
        service_ptr_ = std::unique_ptr<ews::service>(new ews::service(
            env.server_uri, env.domain, env.username, env.password));
#endif
    }

    ews::service& service()
    {
        if (!service_ptr_)
        {
            throw std::runtime_error(
                "Cannot access service: no instance created");
        }
        return *service_ptr_;
    }

private:
    std::unique_ptr<ews::service> service_ptr_;
};

class ItemTest : public BaseFixture, public ServiceMixin
{
};

// Create and remove a task on the server
class TaskTest : public BaseFixture, public ServiceMixin
{
public:
    void SetUp()
    {
        BaseFixture::SetUp();

        task_.set_subject("Get some milk");
        task_.set_body(ews::body("Get some milk from the store"));
        task_.set_start_date(ews::date_time("2015-06-17T19:00:00Z"));
        task_.set_due_date(ews::date_time("2015-06-17T19:30:00Z"));
        const auto item_id = service().create_item(task_);
        task_ = service().get_task(item_id, ews::base_shape::all_properties);
    }

    void TearDown()
    {
        service().delete_task(std::move(task_), ews::delete_type::hard_delete,
                              ews::affected_task_occurrences::all_occurrences);

        BaseFixture::TearDown();
    }

    ews::task& test_task() { return task_; }

private:
    ews::task task_;
};

// Create and remove a contact on the server
class ContactTest : public BaseFixture, public ServiceMixin
{
public:
    void SetUp()
    {
        BaseFixture::SetUp();

        contact_.set_given_name("Minerva");
        contact_.set_nickname("Minnie");
        contact_.set_surname("Mouse");
        contact_.set_spouse_name("Mickey");
        contact_.set_job_title("Damsel in distress");

        const auto item_id = service().create_item(contact_);
        contact_ =
            service().get_contact(item_id, ews::base_shape::all_properties);
    }

    void TearDown()
    {
        service().delete_contact(std::move(contact_));

        BaseFixture::TearDown();
    }

    ews::contact& test_contact() { return contact_; }

private:
    ews::contact contact_;
};

class MessageTest : public BaseFixture, public ServiceMixin
{
public:
    void SetUp()
    {
        BaseFixture::SetUp();

        message_.set_subject("Meet the Fockers");
        std::vector<ews::mailbox> recipients;
        recipients.push_back(ews::mailbox("nirvana@example.com"));
        message_.set_to_recipients(recipients);
        const auto item_id = service().create_item(
            message_, ews::message_disposition::save_only);
        message_ =
            service().get_message(item_id, ews::base_shape::all_properties);
    }

    void TearDown()
    {
        service().delete_message(std::move(message_));
        BaseFixture::TearDown();
    }

    ews::message& test_message() { return message_; }

private:
    ews::message message_;
};

class CalendarItemTest : public BaseFixture, public ServiceMixin
{
public:
    void SetUp()
    {
        BaseFixture::SetUp();

        calitem_.set_subject("Meet the Fockers");
        calitem_.set_start(ews::date_time("2004-12-25T10:00:00.000Z"));
        calitem_.set_end(ews::date_time("2004-12-27T10:00:00.000Z"));
        const auto item_id = service().create_item(calitem_);
        calitem_ = service().get_calendar_item(item_id,
                                               ews::base_shape::all_properties);
    }

    void TearDown()
    {
        service().delete_calendar_item(std::move(calitem_));

        BaseFixture::TearDown();
    }

    ews::calendar_item& test_calendar_item() { return calitem_; }

private:
    ews::calendar_item calitem_;
};

class AttachmentTest : public BaseFixture, public ServiceMixin
{
public:
    void SetUp()
    {
        BaseFixture::SetUp();

        auto msg = ews::message();
        msg.set_subject("Honorable Minister of Finance - Release Funds");
        std::vector<ews::mailbox> recipients;
        recipients.push_back(ews::mailbox("udom.emmanuel@zenith-bank.com.ng"));
        msg.set_to_recipients(recipients);
        auto item_id =
            service().create_item(msg, ews::message_disposition::save_only);
        message_ =
            service().get_message(item_id, ews::base_shape::all_properties);
    }

    void TearDown()
    {
        service().delete_message(std::move(message_));

        BaseFixture::TearDown();
    }

    ews::message& test_message() { return message_; }

private:
    ews::message message_;
};

#ifdef EWS_HAS_FILESYSTEM_HEADER
class ResolveNamesTest : public FakeServiceFixture
{
public:
    const std::filesystem::path assets_dir() const
    {
        return std::filesystem::path(assets());
    }
};

class SubscribeTest : public FakeServiceFixture
{
public:
    const std::filesystem::path assets_dir() const
    {
        return std::filesystem::path(assets());
    }
};

class TemporaryDirectoryFixture : public BaseFixture
{
public:
    void SetUp()
    {
        BaseFixture::SetUp();

        olddir_ = std::filesystem::current_path();
        workingdir_ = std::tmpnam(nullptr);
        ASSERT_TRUE(std::filesystem::create_directory(workingdir_))
            << "Unable to create temporary working directory";
        std::filesystem::current_path(workingdir_);
    }

    void TearDown()
    {
        EXPECT_TRUE(std::filesystem::is_empty(workingdir_))
            << "Temporary directory not empty on TearDown";
        std::filesystem::current_path(olddir_);
        std::filesystem::remove_all(workingdir_);

        BaseFixture::TearDown();
    }

    const std::filesystem::path& cwd() const { return workingdir_; }

    const std::filesystem::path assets_dir() const
    {
        return std::filesystem::path(assets());
    }

private:
    std::filesystem::path olddir_;
    std::filesystem::path workingdir_;
};
#endif // EWS_HAS_FILESYSTEM_HEADER

inline ews::task make_fake_task(const char* xml = nullptr)
{
    typedef rapidxml::xml_document<> xml_document;

    if (!xml)
    {
        xml = "<t:Task\n"
              "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/"
              "types\">\n"
              "    <t:ItemId Id=\"abcde\" ChangeKey=\"edcba\"/>\n"
              "    <t:ParentFolderId Id=\"qwertz\" ChangeKey=\"ztrewq\"/>\n"
              "    <t:ItemClass>IPM.Task</t:ItemClass>\n"
              "    <t:Subject>Write poem</t:Subject>\n"
              "    <t:Sensitivity>Confidential</t:Sensitivity>\n"
              "    <t:Body BodyType=\"Text\" IsTruncated=\"false\"/>\n"
              "    "
              "<t:DateTimeReceived>2015-02-09T13:00:11Z</"
              "t:DateTimeReceived>\n"
              "    <t:Size>962</t:Size>\n"
              "    <t:Importance>Normal</t:Importance>\n"
              "    <t:IsSubmitted>false</t:IsSubmitted>\n"
              "    <t:IsDraft>false</t:IsDraft>\n"
              "    <t:IsFromMe>false</t:IsFromMe>\n"
              "    <t:IsResend>false</t:IsResend>\n"
              "    <t:IsUnmodified>false</t:IsUnmodified>\n"
              "    <t:DateTimeSent>2015-02-09T13:00:11Z</t:DateTimeSent>\n"
              "    "
              "<t:DateTimeCreated>2015-02-09T13:00:11Z</t:DateTimeCreated>\n"
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
              "    "
              "<t:LastModifiedTime>2015-02-09T13:00:11Z</"
              "t:LastModifiedTime>\n"
              "    <t:IsAssociated>false</t:IsAssociated>\n"
              "    <t:Flag>\n"
              "            <t:FlagStatus>NotFlagged</t:FlagStatus>\n"
              "    </t:Flag>\n"
              "    "
              "<t:InstanceKey>AQAAAAAAARMBAAAAG4AqWQAAAAA=</t:InstanceKey>\n"
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
    std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
    buf.push_back('\0');
    xml_document doc;
    doc.parse<0>(&buf[0]);
    auto node = doc.first_node();
    return ews::task::from_xml_element(*node);
}

#ifdef EWS_HAS_FILESYSTEM_HEADER
inline ews::message make_fake_message(const char* xml = nullptr)
{
    typedef rapidxml::xml_document<> xml_document;

    std::vector<char> buf;

    if (xml)
    {
        // Load from given C-string

        std::copy(xml, xml + strlen(xml), std::back_inserter(buf));
        buf.push_back('\0');
    }
    else
    {
        // Load from file

        const auto assets = std::filesystem::path(
            ews::test::global_data::instance().assets_dir);
        const auto file_path =
            assets / "undeliverable_test_mail_get_item_response.xml";
        std::ifstream ifstr(file_path.string(),
                            std::ifstream::in | std::ios::binary);
        if (!ifstr.is_open())
        {
            throw std::runtime_error("Could not open file for reading: " +
                                     file_path.string());
        }
        ifstr.unsetf(std::ios::skipws);
        ifstr.seekg(0, std::ios::end);
        const auto file_size = ifstr.tellg();
        ifstr.seekg(0, std::ios::beg);
        buf.reserve(file_size);
        buf.insert(begin(buf), std::istream_iterator<char>(ifstr),
                   std::istream_iterator<char>());
        ifstr.close();
    }

    xml_document doc;
    doc.parse<0>(&buf[0]);
    auto node = ews::internal::get_element_by_qname(
        doc, "Message", ews::internal::uri<>::microsoft::types());
    return ews::message::from_xml_element(*node);
}
#endif // EWS_HAS_FILESYSTEM_HEADER
} // namespace tests
