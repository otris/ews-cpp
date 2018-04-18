
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

#include "fixtures.hpp"

#include <utility>

namespace tests
{
TEST_F(ContactTest, GetContactWithInvalidIdThrows)
{
    auto invalid_id = ews::item_id();
    EXPECT_THROW(service().get_contact(invalid_id), ews::exchange_error);
}

TEST_F(ContactTest, GetContactWithInvalidIdExceptionResponse)
{
    auto invalid_id = ews::item_id();
    try
    {
        service().get_contact(invalid_id);
        FAIL() << "Expected an exception";
    }
    catch (ews::exchange_error& exc)
    {
        EXPECT_EQ(ews::response_code::error_invalid_id_empty, exc.code());
    }
}

TEST(OfflineContactTest, InitialFileAs)
{
    auto minnie = ews::contact();
    EXPECT_STREQ("", minnie.get_file_as().c_str());
}

TEST(OfflineContactTest, SetFileAs)
{
    auto minnie = ews::contact();
    minnie.set_file_as("Minnie Mouse");
    EXPECT_STREQ("Minnie Mouse", minnie.get_file_as().c_str());
}

TEST_F(ContactTest, UpdateFileAs)
{
    auto minnie = test_contact();
    auto prop =
        ews::property(ews::contact_property_path::file_as, "Minnie Mouse");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Minnie Mouse", minnie.get_file_as().c_str());
}

TEST(OfflineContactTest, InitialFileAsMappingValue)
{
    auto minnie = ews::contact();
    EXPECT_EQ(ews::internal::file_as_mapping::none,
              minnie.get_file_as_mapping());
}

TEST(OfflineContactTest, SetFileAsMappingValue)
{
    auto minnie = ews::contact();
    minnie.set_file_as_mapping(
        ews::internal::file_as_mapping::last_comma_first);
    EXPECT_EQ(ews::internal::file_as_mapping::last_comma_first,
              minnie.get_file_as_mapping());
}

TEST_F(ContactTest, UpdateFileAsMappingValue)
{
    auto minnie = test_contact();
    auto prop =
        ews::property(ews::contact_property_path::file_as, "Minnie Mouse");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Minnie Mouse", minnie.get_file_as().c_str());

    prop = ews::property(ews::contact_property_path::file_as_mapping,
                         "LastSpaceFirst");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_EQ(ews::internal::file_as_mapping::last_space_first,
              minnie.get_file_as_mapping());
    EXPECT_STREQ("Mouse Minerva", minnie.get_file_as().c_str());
}

TEST(OfflineContactTest, InitialEmailAddressProperty)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_email_addresses().empty());
}

TEST(OfflineContactTest, SetEmailAddressProperty)
{
    auto minnie = ews::contact();
    auto email = ews::email_address(ews::email_address::key::email_address_1,
                                    "minnie.mouse@duckburg.com");
    minnie.set_email_address(email);
    auto new_mail = minnie.get_email_addresses();
    ASSERT_EQ(1U, new_mail.size());
    EXPECT_EQ(email, new_mail[0]);
}

TEST_F(ContactTest, UpdateEmailAddressProperty)
{
    auto minnie = test_contact();
    auto mail_address = ews::email_address(
        ews::email_address::key::email_address_1, "minnie.mouse@duckburg.com");
    auto prop = ews::property(ews::contact_property_path::email_address_1,
                              mail_address);
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    auto new_mail = minnie.get_email_addresses();
    ASSERT_EQ(1U, new_mail.size());
    EXPECT_EQ(mail_address, new_mail[0]);
}

TEST_F(ContactTest, DeleteEmailAddress)
{
    auto minnie = test_contact();
    auto mail_address = ews::email_address(
        ews::email_address::key::email_address_1, "minnie.mouse@duckburg.com");
    auto prop = ews::property(ews::contact_property_path::email_address_1,
                              mail_address);
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    auto new_mail = minnie.get_email_addresses();
    ASSERT_EQ(1U, new_mail.size());
    EXPECT_EQ(mail_address, new_mail[0]);

    auto update = ews::update(prop, ews::update::operation::delete_item_field);
    new_id = service().update_item(minnie.get_item_id(), update);
    minnie = service().get_contact(new_id);
    EXPECT_TRUE(minnie.get_email_addresses().empty());
}

TEST(OfflineContactTest, InitialGivenNameValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_given_name().empty());
}

TEST(OfflineContactTest, SetGivenNameValue)
{
    auto minnie = ews::contact();
    minnie.set_given_name("Minnie");
    EXPECT_STREQ("Minnie", minnie.get_given_name().c_str());
}

TEST_F(ContactTest, UpdateGivenNameValue)
{
    auto minnie = test_contact();
    auto prop = ews::property(ews::contact_property_path::given_name, "Minnie");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Minnie", minnie.get_given_name().c_str());

    prop = ews::property(ews::contact_property_path::given_name, "Money");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Money", minnie.get_given_name().c_str());
}

TEST(OfflineContactTest, InitialDisplayNameValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_display_name().empty());
}

TEST(OfflineContactTest, SetDisplayNameValue)
{
    auto minnie = ews::contact();
    minnie.set_display_name("Money Maker");
    EXPECT_STREQ("Money Maker", minnie.get_display_name().c_str());
}

TEST_F(ContactTest, UpdateDisplayNameValue)
{
    auto minnie = test_contact();
    auto prop = ews::property(ews::contact_property_path::display_name,
                              "Minerva Mouse");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Minerva Mouse", minnie.get_display_name().c_str());

    prop =
        ews::property(ews::contact_property_path::display_name, "Money Maker");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Money Maker", minnie.get_display_name().c_str());
}

TEST(OfflineContactTest, InitialInitialsValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_initials().empty());
}

TEST(OfflineContactTest, SetInitialsValue)
{
    auto minnie = ews::contact();
    minnie.set_initials("MM");
    EXPECT_STREQ("MM", minnie.get_initials().c_str());
}

TEST_F(ContactTest, UpdateInitialsValue)
{
    auto minnie = test_contact();
    auto prop = ews::property(ews::contact_property_path::initials, "");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("", minnie.get_initials().c_str());

    prop = ews::property(ews::contact_property_path::initials, "MM");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("MM", minnie.get_initials().c_str());
}

TEST(OfflineContactTest, InitialMiddleNameValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_middle_name().empty());
}

TEST(OfflineContactTest, SetMiddleNameValue)
{
    auto minnie = ews::contact();
    minnie.set_middle_name("Money");
    EXPECT_STREQ("Money", minnie.get_middle_name().c_str());
}

TEST_F(ContactTest, UpdateMiddleNameValue)
{
    auto minnie = test_contact();
    auto prop = ews::property(ews::contact_property_path::middle_name, "Mani");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Mani", minnie.get_middle_name().c_str());

    prop = ews::property(ews::contact_property_path::middle_name, "Money");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Money", minnie.get_middle_name().c_str());
}

TEST(OfflineContactTest, InitialNicknameValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_nickname().empty());
}

TEST(OfflineContactTest, SetNicknameValue)
{
    auto minnie = ews::contact();
    minnie.set_nickname("Money");
    EXPECT_STREQ("Money", minnie.get_nickname().c_str());
}

TEST_F(ContactTest, UpdateNicknameValue)
{
    auto minnie = test_contact();
    auto prop = ews::property(ews::contact_property_path::nickname, "Money");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Money", minnie.get_nickname().c_str());

    prop = ews::property(ews::contact_property_path::nickname, "Geld");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Geld", minnie.get_nickname().c_str());
}

TEST(OfflineContactTest, InitialCompanyNameValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_company_name().empty());
}

TEST(OfflineContactTest, SetCompanyNameValue)
{
    auto minnie = ews::contact();
    minnie.set_company_name("Money Bin");
    EXPECT_STREQ("Money Bin", minnie.get_company_name().c_str());
}

TEST_F(ContactTest, UpdateCompanyNameValue)
{
    auto minnie = test_contact();
    auto prop =
        ews::property(ews::contact_property_path::company_name, "Money Bin");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Money Bin", minnie.get_company_name().c_str());

    prop = ews::property(ews::contact_property_path::company_name,
                         "Tarantinos Bar and Restaurant");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Tarantinos Bar and Restaurant",
                 minnie.get_company_name().c_str());
}

TEST(OfflineContactTest, InitialAssistantNameValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_assistant_name().empty());
}

TEST(OfflineContactTest, SetAssistantNameValue)
{
    auto minnie = ews::contact();
    minnie.set_assistant_name("Pluto");
    EXPECT_STREQ("Pluto", minnie.get_assistant_name().c_str());
}

TEST_F(ContactTest, UpdateAssistantNameValue)
{
    auto minnie = test_contact();
    auto prop =
        ews::property(ews::contact_property_path::assistant_name, "Pluto");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Pluto", minnie.get_assistant_name().c_str());

    prop =
        ews::property(ews::contact_property_path::assistant_name, "Plutocrat");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Plutocrat", minnie.get_assistant_name().c_str());
}

TEST(OfflineContactTest, InitialBusinessHomePageValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_business_homepage().empty());
}

TEST(OfflineContactTest, SetBusinessHomePageValue)
{
    auto minnie = ews::contact();
    minnie.set_business_homepage("holstensicecream.com");
    EXPECT_STREQ("holstensicecream.com",
                 minnie.get_business_homepage().c_str());
}

TEST_F(ContactTest, UpdateBusinessHomePageValue)
{
    auto minnie = test_contact();
    auto prop = ews::property(ews::contact_property_path::business_home_page,
                              "holstensicecream.com");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("holstensicecream.com",
                 minnie.get_business_homepage().c_str());

    prop = ews::property(ews::contact_property_path::business_home_page,
                         "lainchan.org");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("lainchan.org", minnie.get_business_homepage().c_str());
}

TEST(OfflineContactTest, InitalDepartmentValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_department().empty());
}

TEST(OfflineContactTest, SetDepartmentValue)
{
    auto minnie = ews::contact();
    minnie.set_department("Human Resources");
    EXPECT_STREQ("Human Resources", minnie.get_department().c_str());
}

TEST_F(ContactTest, UpdateDepartmentValue)
{
    auto minnie = test_contact();
    auto prop = ews::property(ews::contact_property_path::department,
                              "Human Resources");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Human Resources", minnie.get_department().c_str());

    prop = ews::property(ews::contact_property_path::department,
                         "Anthropomorphic Research");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Anthropomorphic Research", minnie.get_department().c_str());
}

TEST(OfflineContactTest, InitialGenerationValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_generation().empty());
}

TEST(OfflineContactTest, SetGenerationValue)
{
    auto minnie = ews::contact();
    minnie.set_generation("III");
    EXPECT_STREQ("III", minnie.get_generation().c_str());
}

TEST_F(ContactTest, UpdateGenerationValue)
{
    auto minnie = test_contact();
    auto prop = ews::property(ews::contact_property_path::generation, "III");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("III", minnie.get_generation().c_str());

    prop = ews::property(ews::contact_property_path::generation, "Jr.");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Jr.", minnie.get_generation().c_str());
}

TEST(OfflineContactTest, InitialJobTitleValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_job_title().empty());
}

TEST(OfflineContactTest, SetJobTitleValue)
{
    auto minnie = ews::contact();
    minnie.set_job_title("Unemployed");
    EXPECT_STREQ("Unemployed", minnie.get_job_title().c_str());
}

TEST_F(ContactTest, UpdateJobTitleValue)
{
    auto minnie = test_contact();
    auto prop =
        ews::property(ews::contact_property_path::job_title, "Unemployed");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Unemployed", minnie.get_job_title().c_str());

    prop = ews::property(ews::contact_property_path::job_title, "Engineer");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Engineer", minnie.get_job_title().c_str());
}

TEST(OfflineContactTest, InitialManagerValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_manager().empty());
}

TEST(OfflineContactTest, SetManagerValue)
{
    auto minnie = ews::contact();
    minnie.set_manager("Scrooge McDuck");
    EXPECT_STREQ("Scrooge McDuck", minnie.get_manager().c_str());
}

TEST_F(ContactTest, UpdateManagerValue)
{
    auto minnie = test_contact();
    auto prop =
        ews::property(ews::contact_property_path::manager, "Scrooge McDuck");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Scrooge McDuck", minnie.get_manager().c_str());

    prop = ews::property(ews::contact_property_path::manager,
                         "Flintheart Glomgold");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Flintheart Glomgold", minnie.get_manager().c_str());
}

TEST(OfflineContactTest, InitialMileageValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_mileage().empty());
}

TEST(OfflineContactTest, SetMileageValue)
{
    auto minnie = ews::contact();
    minnie.set_mileage("20km");
    EXPECT_STREQ("20km", minnie.get_mileage().c_str());
}

TEST_F(ContactTest, UpdateMileageValue)
{
    auto minnie = test_contact();
    auto prop = ews::property(ews::contact_property_path::mileage, "20km");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("20km", minnie.get_mileage().c_str());

    prop = ews::property(ews::contact_property_path::mileage, "Infinite");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Infinite", minnie.get_mileage().c_str());
}

TEST(OfflineContactTest, InitialOfficeLocationValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_office_location().empty());
}

TEST(OfflineContactTest, SetOfficeLocationValue)
{
    auto minnie = ews::contact();
    minnie.set_office_location("Duckburg");
    EXPECT_STREQ("Duckburg", minnie.get_office_location().c_str());
}

TEST_F(ContactTest, UpdateOfficeLocationValue)
{
    auto minnie = test_contact();
    auto prop =
        ews::property(ews::contact_property_path::office_location, "Duckburg");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Duckburg", minnie.get_office_location().c_str());

    prop =
        ews::property(ews::contact_property_path::office_location, "Detroit");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Detroit", minnie.get_office_location().c_str());
}

TEST(OfflineContactTest, InitialProfessionValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_profession().empty());
}

TEST(OfflineContactTest, SetProfessionValue)
{
    auto minnie = ews::contact();
    minnie.set_profession("Veterinarian");
    EXPECT_STREQ("Veterinarian", minnie.get_profession().c_str());
}

TEST_F(ContactTest, UpdateProfessionValue)
{
    auto minnie = test_contact();
    auto prop =
        ews::property(ews::contact_property_path::profession, "Veterinarian");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Veterinarian", minnie.get_profession().c_str());

    prop = ews::property(ews::contact_property_path::profession, "Engineer");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Engineer", minnie.get_profession().c_str());
}

TEST(OfflineContactTest, InitialSpouseName)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_spouse_name().empty());
}

TEST(OfflineContactTest, SetSpouseName)
{
    auto minnie = ews::contact();
    minnie.set_spouse_name("Donald");
    EXPECT_STREQ("Donald", minnie.get_spouse_name().c_str());
}

TEST_F(ContactTest, UpdateSpouseName)
{
    auto minnie = test_contact();
    auto prop =
        ews::property(ews::contact_property_path::spouse_name, "Donald");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Donald", minnie.get_spouse_name().c_str());

    prop = ews::property(ews::contact_property_path::spouse_name, "Scrooge");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("Scrooge", minnie.get_spouse_name().c_str());
}

TEST(OfflineContactTest, InitialSurnameValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_surname().empty());
}

TEST(OfflineContactTest, SetSurnameValue)
{
    auto minnie = ews::contact();
    minnie.set_surname("McDuck");
    EXPECT_STREQ("McDuck", minnie.get_surname().c_str());
}

TEST_F(ContactTest, UpdateSurnameValue)
{
    auto minnie = test_contact();
    auto prop = ews::property(ews::contact_property_path::surname, "McDuck");
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("McDuck", minnie.get_surname().c_str());

    prop = ews::property(ews::contact_property_path::surname, "Gibson");
    new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    EXPECT_STREQ("Gibson", minnie.get_surname().c_str());
}

TEST(OfflineContactTest, InitialPhysicalAddressesValues)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_physical_addresses().empty());
}

TEST(OfflineContactTest, SetPhysicalAddressesValues)
{
    auto minnie = ews::contact();
    auto address =
        ews::physical_address(ews::physical_address::key::home, "Doomroad",
                              "Doomburg", "Doom", "Doomonia", "4 15 15 13");
    minnie.set_physical_address(address);
    const auto addresses = minnie.get_physical_addresses();
    ASSERT_FALSE(addresses.empty());
    EXPECT_EQ(address, addresses[0]);
}

TEST_F(ContactTest, UpdatePhysicalAddressesValues)
{
    auto minnie = test_contact();
    auto address = ews::physical_address(ews::physical_address::key::home, "",
                                         "Duckburg", "", "", "");
    auto prop = ews::property(
        ews::contact_property_path::physical_address::city, address);
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    ASSERT_FALSE(minnie.get_physical_addresses().empty());
    auto new_address = minnie.get_physical_addresses();
    EXPECT_EQ(address, new_address[0]);
}

TEST(OfflineContactTest, InitialBirthdayValue)
{
    auto minnie = ews::contact();
    EXPECT_STREQ("", minnie.get_birthday().c_str());
}

TEST(OfflineContactTest, SetBirthdayValue)
{
    auto minnie = ews::contact();
    minnie.set_birthday("1982-08-01");
    EXPECT_STREQ("1982-08-01", minnie.get_birthday().c_str());
}

TEST_F(ContactTest, UpdateBirthdayValue)
{
    auto minnie = test_contact();
    auto birthday = ews::date_time("1994-11-03");
    auto prop = ews::property(ews::contact_property_path::birthday, birthday);
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("1994-11-03T00:00:00Z", minnie.get_birthday().c_str());
}

TEST(OfflineContactTest, InitialWeddingAnniversaryValue)
{
    auto minnie = ews::contact();
    EXPECT_STREQ("", minnie.get_wedding_anniversary().c_str());
}

TEST(OfflineContactTest, SetWeddingAnniversaryValue)
{
    auto minnie = ews::contact();
    minnie.set_wedding_anniversary("1953-03-16");
    EXPECT_STREQ("1953-03-16", minnie.get_wedding_anniversary().c_str());
}

TEST_F(ContactTest, UpdateWeddingAnniversaryValue)
{
    auto minnie = test_contact();
    auto wedding = ews::date_time("2006-06-06");
    auto prop =
        ews::property(ews::contact_property_path::wedding_anniversary, wedding);
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_STREQ("2006-06-06T00:00:00Z",
                 minnie.get_wedding_anniversary().c_str());
}

TEST(OfflineContactTest, InitialChildrenValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_children().empty());
}

TEST(OfflineContactTest, SetChildrenValue)
{
    auto minnie = ews::contact();
    std::vector<std::string> children;
    children.push_back("Andie");
    children.push_back("Bandie");
    minnie.set_children(children);
    auto first_child = children[0].c_str();
    EXPECT_STREQ("Andie", first_child);
}

TEST_F(ContactTest, UpdateChildrenValue)
{
    auto minnie = test_contact();
    std::vector<std::string> children;
    children.push_back("Ando");
    children.push_back("Bando");
    auto prop = ews::property(ews::contact_property_path::children, children);
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_FALSE(minnie.get_children().empty());
}

TEST(OfflineContactTest, InitialCompaniesValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_companies().empty());
}

TEST(OfflineContactTest, SetCompaniesValue)
{
    auto minnie = ews::contact();
    std::vector<std::string> companies;
    companies.push_back("Otris GmbH");
    companies.push_back("Aperture Science");
    minnie.set_companies(companies);
    auto first_company = companies[0].c_str();
    EXPECT_STREQ("Otris GmbH", first_company);
}

TEST_F(ContactTest, UpdateCompaniesValue)
{
    auto minnie = test_contact();
    std::vector<std::string> companies;
    companies.push_back("Otris GmbH");
    companies.push_back("Aperture Science");
    auto prop = ews::property(ews::contact_property_path::companies, companies);
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id, ews::base_shape::all_properties);
    EXPECT_FALSE(minnie.get_companies().empty());
}

TEST(OfflineContactTest, InitialImAddressesValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_im_addresses().empty());
}

TEST(OfflineContactTest, SetImAddressValue)
{
    auto minnie = ews::contact();
    auto address = ews::im_address(ews::im_address::key::imaddress1, "MMouse");
    minnie.set_im_address(address);
    auto im_addresses = minnie.get_im_addresses();
    EXPECT_EQ(address, im_addresses[0]);
}

TEST_F(ContactTest, UpdateImAddressesValue)
{
    auto minnie = test_contact();
    auto address = ews::im_address(ews::im_address::key::imaddress1, "MMouse");
    auto prop =
        ews::property(ews::contact_property_path::im_address_1, address);
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    auto new_address = minnie.get_im_addresses();
    EXPECT_EQ(address, new_address[0]);
}

TEST(OfflineContactTest, InitialPhoneNumberValue)
{
    auto minnie = ews::contact();
    EXPECT_TRUE(minnie.get_phone_numbers().empty());
}

TEST(OfflineContactTest, SetPhoneNumberValue)
{
    auto minnie = ews::contact();
    auto phone_number =
        ews::phone_number(ews::phone_number::key::home_phone, "0123456789");
    minnie.set_phone_number(phone_number);
    auto new_number = minnie.get_phone_numbers();
    ASSERT_FALSE(new_number.empty());
    EXPECT_EQ(new_number[0], phone_number);
}

TEST_F(ContactTest, UpdatePhoneNumberValue)
{
    auto minnie = test_contact();

    auto prop = ews::property(
        ews::contact_property_path::phone_number::home_phone,
        ews::phone_number(ews::phone_number::key::home_phone, "9876543210"));
    auto new_id = service().update_item(minnie.get_item_id(), prop);
    minnie = service().get_contact(new_id);
    ASSERT_FALSE(minnie.get_phone_numbers().empty());
    auto numbers = minnie.get_phone_numbers();
    EXPECT_EQ(ews::phone_number::key::home_phone, numbers[0].get_key());
    EXPECT_EQ("9876543210", numbers[0].get_value());
}

TEST_F(ContactTest, ContactSourceValue)
{
    auto minnie = test_contact();
    EXPECT_STREQ("", minnie.get_contact_source().c_str());
}

// TODO:
// PostalAddressIndex

TEST_F(ContactTest, GetCompleteNameProperty)
{
    auto minnie = test_contact();

    const auto complete_name = minnie.get_complete_name();

    EXPECT_STREQ("", complete_name.get_title().c_str());
    EXPECT_STREQ("Minerva", complete_name.get_first_name().c_str());
    EXPECT_STREQ("", complete_name.get_middle_name().c_str());
    EXPECT_STREQ("Mouse", complete_name.get_last_name().c_str());
    EXPECT_STREQ("", complete_name.get_suffix().c_str());
    EXPECT_STREQ("", complete_name.get_initials().c_str());
    EXPECT_STREQ("Minerva Mouse", complete_name.get_full_name().c_str());
    EXPECT_STREQ("Minnie", complete_name.get_nickname().c_str());
}
} // namespace tests
