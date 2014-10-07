#include "plumbing.hpp"
#include "porcelain.hpp"

int main()
{
    ews::set_up();

    auto doc = ews::plumbing::make_raw_soap_request("", "", "", "", "",
                                                    std::vector<std::string>{});
    (void)doc;

    // hack hack ...

    ews::tear_down();
    return 0;
}
