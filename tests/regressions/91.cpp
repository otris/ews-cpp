
//   Copyright 2018 otris software AG
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

#include <ews/ews.hpp>
#include <gtest/gtest.h>

#include "fixtures.hpp"

namespace tests
{

typedef ItemTest Issue91Test;

TEST_F(Issue91Test, MakeSureUpdateEscapesXMLValues)
{
    auto msg = ews::message();
    msg.set_subject("some text with ampersand &");
    auto id = service().create_item(msg, ews::message_disposition::save_only);
    ews::internal::on_scope_exit remove_message(
        [&]() { service().delete_item(id); });

    ews::property prop(ews::item_property_path::subject,
                       "this should work too &");
    ews::update update(prop, ews::update::operation::set_item_field);
    id = service().update_item(id, update);
}
}


