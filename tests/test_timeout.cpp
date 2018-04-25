#include "fixtures.hpp"

#include <chrono>
#include <cstdlib>
#include <cstring>
#include <ctime>
#include <ews/ews.hpp>
#include <iterator>
#include <stdexcept>
#include <string>
#include <thread>
#include <vector>

#ifdef EWS_HAS_PYTHON_EXECUTABLE

namespace tests
{
namespace timeout_test
{
    inline void start(const std::string& dir)
    {
        static bool server_started = false;
        if (!server_started)
        {
            const auto cmd =
                std::string("python ") + dir + std::string("/wait_server.py");
            std::thread server_thread([dir, cmd] { std::system(cmd.c_str()); });
            std::this_thread::sleep_for(std::chrono::seconds(3));
            server_thread.detach();
            server_started = true;
        }
    }
} // namespace timeout_test

class TimeoutTest : public BaseFixture
{
};

TEST_F(TimeoutTest, ReachTimeout)
{
    timeout_test::start(assets());
    bool reached_timeout = false;
    try
    {
        ews::internal::http_request r("http://127.0.0.1:8080/");
        r.set_timeout(std::chrono::seconds(1));

        r.send("");
    }
    catch (std::exception&)
    {
        reached_timeout = true;
    }
    EXPECT_TRUE(reached_timeout);
}

TEST_F(TimeoutTest, NotReachTimeout)
{
    timeout_test::start(assets());

    bool reached_timeout = false;
    try
    {
        ews::internal::http_request r("http://127.0.0.1:8080/");
        r.set_timeout(std::chrono::seconds(30));

        r.send("");
    }
    catch (std::exception&)
    {
        reached_timeout = true;
    }
    EXPECT_TRUE(!reached_timeout);
}

TEST_F(TimeoutTest, NoTimeout)
{
    timeout_test::start(assets());

    bool reached_timeout = false;
    try
    {
        ews::internal::http_request r("http://127.0.0.1:8080/");

        r.send("");
    }
    catch (std::exception&)
    {
        reached_timeout = true;
    }
    EXPECT_TRUE(!reached_timeout);
}
} // namespace tests

#endif
