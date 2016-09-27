
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
            EXPECT_STREQ("ErrorInvalidIdEmpty", exc.what());
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
        EXPECT_STREQ("", minnie.get_file_as_mapping().c_str());
    }

    TEST(OfflineContactTest, SetFileAsMappingValue)
    {
        auto minnie = ews::contact();
        minnie.set_file_as_mapping("LastCommaFirst");
        EXPECT_STREQ("LastCommaFirst", minnie.get_file_as_mapping().c_str());
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
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("LastSpaceFirst", minnie.get_file_as_mapping().c_str());
        EXPECT_STREQ("Mouse Minerva", minnie.get_file_as().c_str());
    }

    TEST(OfflineContactTest, InitialEmailAddressProperty)
    {
        auto minnie = ews::contact();
        EXPECT_STREQ("", minnie.get_email_address_1().c_str());
        EXPECT_STREQ("", minnie.get_email_address_2().c_str());
        EXPECT_STREQ("", minnie.get_email_address_3().c_str());
        EXPECT_TRUE(minnie.get_email_addresses().empty());
    }

    TEST(OfflineContactTest, SetEmailAddressProperty)
    {
        auto minnie = ews::contact();
        minnie.set_email_address_1(ews::mailbox("minnie.mouse@duckburg.com"));
        EXPECT_STREQ("minnie.mouse@duckburg.com",
                     minnie.get_email_address_1().c_str());
        EXPECT_STREQ("", minnie.get_email_address_2().c_str());
        EXPECT_STREQ("", minnie.get_email_address_3().c_str());
        auto addresses = minnie.get_email_addresses();
        ASSERT_FALSE(addresses.empty());
        EXPECT_STREQ("minnie.mouse@duckburg.com",
                     addresses.front().value().c_str());
    }

    TEST_F(ContactTest, UpdateEmailAddressProperty)
    {
        auto minnie = test_contact();
        auto prop = ews::property(ews::contact_property_path::email_address_1,
                                  "minnie.mouse@duckburg.com");
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("minnie.mouse@duckburg.com",
                     minnie.get_email_address_1().c_str());
    }

    TEST_F(ContactTest, DeleteEmailAddress)
    {
        auto minnie = test_contact();
        auto prop = ews::property(ews::contact_property_path::email_address_1,
                                  "minnie.mouse@duckburg.com");
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        ASSERT_STREQ("minnie.mouse@duckburg.com",
                     minnie.get_email_address_1().c_str());   
    
        prop = ews::property(ews::contact_property_path::email_address_1);
        new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("", minnie.get_email_address_1().c_str());
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
        auto prop =
            ews::property(ews::contact_property_path::given_name, "Minnie");
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

        prop = ews::property(ews::contact_property_path::display_name,
                             "Money Maker");
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
        auto prop =
            ews::property(ews::contact_property_path::middle_name, "Mani");
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
        auto prop =
            ews::property(ews::contact_property_path::nickname, "Money");
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
        auto prop = ews::property(ews::contact_property_path::company_name,
                                  "Money Bin");
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
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("Pluto", minnie.get_assistant_name().c_str());

        prop = ews::property(ews::contact_property_path::assistant_name,
                             "Plutocrat");
        new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
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
        auto prop =
            ews::property(ews::contact_property_path::business_home_page,
                          "holstensicecream.com");
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("holstensicecream.com",
                     minnie.get_business_homepage().c_str());

        prop = ews::property(ews::contact_property_path::business_home_page,
                             "lainchan.org");
        new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
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
        EXPECT_STREQ("Anthropomorphic Research",
                     minnie.get_department().c_str());
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
        auto prop =
            ews::property(ews::contact_property_path::generation, "III");
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("III", minnie.get_generation().c_str());

        prop = ews::property(ews::contact_property_path::generation, "Jr.");
        new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
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
        auto prop = ews::property(ews::contact_property_path::manager,
                                  "Scrooge McDuck");
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("Scrooge McDuck", minnie.get_manager().c_str());

        prop = ews::property(ews::contact_property_path::manager,
                             "Flintheart Glomgold");
        new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
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
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("20km", minnie.get_mileage().c_str());

        prop = ews::property(ews::contact_property_path::mileage, "Infinite");
        new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
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
        auto prop = ews::property(ews::contact_property_path::office_location,
                                  "Duckburg");
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("Duckburg", minnie.get_office_location().c_str());

        prop = ews::property(ews::contact_property_path::office_location,
                             "Detroit");
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
        auto prop = ews::property(ews::contact_property_path::profession,
                                  "Veterinarian");
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("Veterinarian", minnie.get_profession().c_str());

        prop =
            ews::property(ews::contact_property_path::profession, "Engineer");
        new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
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
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("Donald", minnie.get_spouse_name().c_str());

        prop =
            ews::property(ews::contact_property_path::spouse_name, "Scrooge");
        new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
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
        auto prop =
            ews::property(ews::contact_property_path::surname, "McDuck");
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
            ews::physical_address(ews::physical_address::key::home, "Duckroad",
                                  "Duckburg", "", "", "");
        auto address2 =
            ews::physical_address(ews::physical_address::key::business,
                                  "Duckstreet", "Duckburg", "", "", "");
        minnie.set_physical_address(address);
        minnie.set_physical_address(address2);
        const auto addresses = minnie.get_physical_addresses();
        ASSERT_EQ(2U, addresses.size());
        auto home = ews::internal::enum_to_str(addresses[0].get_key()).c_str();
        EXPECT_STREQ("Home", home);
        auto business =
            ews::internal::enum_to_str(addresses[1].get_key()).c_str();
        EXPECT_STREQ("Business", business);
    }

    TEST_F(ContactTest, UpdatePhysicalAddressesValues)
    {
        auto minnie = test_contact();
        auto address =
            ews::physical_address(ews::physical_address::key::home, "Duckroad",
                                  "Duckburg", "", "", "");
        auto prop = ews::property(ews::contact_property_path::city, "Duckburg");
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_FALSE(minnie.get_physical_addresses().empty());
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
        auto prop =
            ews::property(ews::contact_property_path::birthday, "1994-11-03");
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
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
        auto prop = ews::property(
            ews::contact_property_path::wedding_anniversary, "2001-09-11");
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_STREQ("2001-09-11T00:00:00Z",
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
        std::vector<std::string> children{"Andie", "Bandie"};
        minnie.set_children(children);
        auto first_child = children[0].c_str();
        EXPECT_STREQ("Andie", first_child);
    }

    TEST_F(ContactTest, UpdateChildrenValue)
    {
        auto minnie = test_contact();
        std::vector<std::string> children{"Ando", "Bando"};
        auto prop =
            ews::property(ews::contact_property_path::children, children);
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
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
        std::vector<std::string> companies{"Otris GmbH", "Aperture Science"};
        minnie.set_companies(companies);
        auto first_company = companies[0].c_str();
        EXPECT_STREQ("Otris GmbH", first_company);
    }

    TEST_F(ContactTest, UpdateCompaniesValue)
    {
        auto minnie = test_contact();
        std::vector<std::string> companies {"Otris GmbH", "Aperture Science"};
        auto prop = ews::property(ews::contact_property_path::companies, companies);
        auto new_id = service().update_item(minnie.get_item_id(), prop);
        minnie = service().get_contact(new_id);
        EXPECT_FALSE(minnie.get_companies().empty());
    }


    // TODO:
    // PhoneNumbers
    // Companies
    // ContactSource read-only!!
    // ImAddresses
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
}

// vim:et ts=4 sw=4 noic cc=80
