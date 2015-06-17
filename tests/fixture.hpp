#pragma once

#include <memory>

#include <ews/ews.hpp>
#include <ews/ews_test_support.hpp>
#include <gtest/gtest.h>

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
}
