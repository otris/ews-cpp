// Macro for verifying expressions at run-time. Calls assert() with 'expr'.
// Allows turning assertions off, even if -DNDEBUG wasn't given at
// compile-time.  This macro does nothing unless EWS_ENABLE_ASSERTS was defined
// during compilation. No include-guard here on purpose.
#define EWS_ASSERT(expr) ((void)0)
#ifndef NDEBUG
# ifdef EWS_ENABLE_ASSERTS
#  include <cassert>
#  undef EWS_ASSERT
#  define EWS_ASSERT(expr) assert(expr)
# endif
#endif // !NDEBUG
