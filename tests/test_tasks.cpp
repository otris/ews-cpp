#include "fixture.hpp"

#include <memory>

namespace tests
{
    class TaskTest : public BaseFixture
    {
    public:
        virtual ~TaskTest() = default;

        ews::service& service()
        {
            return *service_ptr_;
        }

    private:
        std::unique_ptr<ews::service> service_ptr_;

        virtual void SetUp()
        {
            BaseFixture::SetUp();
            const auto& creds = Environment::credentials();
            service_ptr_ = std::make_unique<ews::service>(
                    ews::service(creds.server_uri,
                                 creds.domain,
                                 creds.username,
                                 creds.password));
        }

        virtual void TearDown()
        {
            service_ptr_.reset();
            BaseFixture::TearDown();
        }
    };

    TEST_F(TaskTest, CreateTask)
    {
        auto start_time = ews::date_time("2015-01-16T12:00:00Z");
        auto end_time   = ews::date_time("2015-01-16T12:30:00Z");

        auto task = ews::task();
        task.set_subject("Something important to do");
        task.set_body(ews::body("Some descriptive body text"));
        task.set_owner("Donald Duck");
        task.set_start_date(start_time);
        task.set_due_date(end_time);
        task.set_reminder_enabled(true);
        task.set_reminder_due_by(start_time);

        auto item_id = service().create_item(task);
    }
}
