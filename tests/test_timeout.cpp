#include "fixtures.hpp"

#include <chrono>
#include <cstring>
#include <ctime>
#include <ews/ews.hpp>
#include <iterator>
#include <stdexcept>
#include <stdlib.h>
#include <string>
#include <thread>
#include <vector>

#ifdef EWS_HAS_PYTHON

namespace tests
{

    namespace timeout_test
    {
        bool server_started = false;
        void start(std::string dir)
        {
            if (!timeout_test::server_started)
            {
                std::thread server_thread([&dir]() {
                    (std::system((std::string("python3 ") + dir +
                                  std::string("/wait_server.py"))
                                     .c_str()));
                });
                server_thread.detach();
                std::this_thread::sleep_for(std::chrono::seconds(3));
                timeout_test::server_started = true;
            }
        }
    }

    class TimeoutTest : public AssetsFixture
    {
    };

    TEST_F(TimeoutTest, ReachTimeout)
    {
        timeout_test::start(assets_dir().string());
        bool reachedTimeout = false;
        try
        {
            ews::internal::http_request r("http://127.0.0.1:8080/");
            r.set_timeout(1);

            r.send("");
        }
        catch (std::exception& e)
        {
            reachedTimeout = true;
        }
        EXPECT_TRUE(reachedTimeout);
    }

    TEST_F(TimeoutTest, NotReachTimeout)
    {
        timeout_test::start(assets_dir().string());

        bool reachedTimeout = false;
        try
        {
            ews::internal::http_request r("http://127.0.0.1:8080/");
            r.set_timeout(30);

            r.send("");
        }
        catch (std::exception& e)
        {
            reachedTimeout = true;
        }
        EXPECT_TRUE(!reachedTimeout);
    }

    TEST_F(TimeoutTest, NoTimeout)
    {
        timeout_test::start(assets_dir().string());

        bool reachedTimeout = false;
        try
        {
            ews::internal::http_request r("http://127.0.0.1:8080/");

            r.send("");
        }
        catch (std::exception& e)
        {
            reachedTimeout = true;
        }
        EXPECT_TRUE(!reachedTimeout);
    }
}

#endif