#pragma once

#include "curl.hpp"

namespace ews
{
    // Function is not thread-safe; should be set-up when application is still
    // in single-threaded context. Calling this function more than once does no
    // harm.
    inline void set_up() { curl::curl_global_init(CURL_GLOBAL_DEFAULT); }

    // Function is not thread-safe; you should call this function only when no
    // other thread is running (see libcurl(3) man-page or
    // http://curl.haxx.se/libcurl/c/libcurl.html)
    inline void tear_down() { curl::curl_global_cleanup(); }
}
