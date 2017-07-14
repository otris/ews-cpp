#pragma once

#if defined(__clang__)
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wdeprecated"
#pragma GCC diagnostic ignored "-Wmissing-noreturn"
#pragma GCC diagnostic ignored "-Wshift-sign-overflow"
#pragma GCC diagnostic ignored "-Wused-but-marked-unused"
#endif

#if defined(__GNUC__)
#pragma GCC diagnostic ignored "-Wredundant-decls"
#endif

#include <gtest/gtest.h>

#if defined(__clang__)
#pragma GCC diagnostic pop
#endif

// vim:et ts=4 sw=4
