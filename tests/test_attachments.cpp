#include "fixtures.hpp"

#include <ews/rapidxml/rapidxml_print.hpp>
#include <string>
#include <vector>
#include <iterator>
#include <stdexcept>
#include <cstring>

typedef rapidxml::xml_document<> xml_document;

#ifdef EWS_USE_BOOST_LIBRARY
namespace
{
    inline std::vector<char> read_contents(const boost::filesystem::path& path)
    {
        std::ifstream ifstr(path.string(),
                            std::ifstream::in | std::ios::binary);
        if (!ifstr.is_open())
        {
            throw std::runtime_error("Could not open file for reading: " +
                    path.string());
        }

        ifstr.unsetf(std::ios::skipws);

        ifstr.seekg(0, std::ios::end);
        const auto file_size = ifstr.tellg();
        ifstr.seekg(0, std::ios::beg);

        auto contents = std::vector<char>();
        contents.reserve(file_size);

        contents.insert(begin(contents),
                      std::istream_iterator<unsigned char>(ifstr),
                      std::istream_iterator<unsigned char>());
        ifstr.close();
        return contents;
    }
}
#endif // EWS_USE_BOOST_LIBRARY

namespace tests
{
    TEST(AttachmentIdTest, DefaultConstruction)
    {
        auto obj = ews::attachment_id();
        EXPECT_FALSE(obj.valid());
        EXPECT_STREQ("", obj.id().c_str());
        EXPECT_FALSE(obj.root_item_id().valid());
    }

    TEST(AttachmentIdTest, ConstructFromId)
    {
        // As in GetAttachment and DeleteAttachment operations
        auto obj = ews::attachment_id("abcde");
        EXPECT_TRUE(obj.valid());
        EXPECT_STREQ("abcde", obj.id().c_str());
        EXPECT_FALSE(obj.root_item_id().valid());
    }

    TEST(AttachmentIdTest, ConstructFromIdAndRootItemId)
    {
        // As in CreateAttachment operation
        auto obj = ews::attachment_id("abcde",
                                      ews::item_id("edcba", "qwertz"));
        EXPECT_TRUE(obj.valid());
        EXPECT_STREQ("abcde", obj.id().c_str());
        EXPECT_TRUE(obj.root_item_id().valid());
        EXPECT_STREQ("edcba", obj.root_item_id().id().c_str());
        EXPECT_STREQ("qwertz", obj.root_item_id().change_key().c_str());
    }

    TEST(AttachmentIdTest, FromXMLNodeWithIdAttributeOnly)
    {
        const char* xml = "<AttachmentId Id=\"abcde\"/>";
        std::vector<char> buf(xml, xml + std::strlen(xml));
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = doc.first_node();
        auto obj = ews::attachment_id::from_xml_element(*node);
        EXPECT_TRUE(obj.valid());
        EXPECT_STREQ("abcde", obj.id().c_str());
        EXPECT_FALSE(obj.root_item_id().valid());
    }

    TEST(AttachmentIdTest, FromXMLNodeWithIdAndRootIdAttributes)
    {
        const char* xml =
"<AttachmentId Id=\"abcde\" RootItemId=\"qwertz\" RootItemChangeKey=\"edcba\"/>";
        std::vector<char> buf(xml, xml + std::strlen(xml));
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = doc.first_node();
        auto obj = ews::attachment_id::from_xml_element(*node);
        EXPECT_TRUE(obj.valid());
        EXPECT_STREQ("abcde", obj.id().c_str());
        EXPECT_TRUE(obj.root_item_id().valid());
        EXPECT_STREQ("qwertz", obj.root_item_id().id().c_str());
        EXPECT_STREQ("edcba", obj.root_item_id().change_key().c_str());
    }

    TEST(AttachmentIdTest, ToXMLWithNamespace)
    {
        const char* expected =
"<ns:AttachmentId Id=\"abcde\" RootItemId=\"qwertz\" RootItemChangeKey=\"edcba\"/>";
        const auto obj =
            ews::attachment_id("abcde", ews::item_id("qwertz", "edcba"));
        EXPECT_STREQ(expected, obj.to_xml("ns").c_str());
    }

    TEST(AttachmentIdTest, ToXML)
    {
        const char* expected =
"<AttachmentId Id=\"abcde\" RootItemId=\"qwertz\" RootItemChangeKey=\"edcba\"/>";
        const auto obj =
            ews::attachment_id("abcde", ews::item_id("qwertz", "edcba"));
        EXPECT_STREQ(expected, obj.to_xml().c_str());
    }

    TEST(AttachmentIdTest, FromAndToXMLRoundTrip)
    {
        const char* xml =
"<AttachmentId Id=\"abcde\" RootItemId=\"qwertz\" RootItemChangeKey=\"edcba\"/>";
        std::vector<char> buf(xml, xml + std::strlen(xml));
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = doc.first_node();
        auto obj = ews::attachment_id::from_xml_element(*node);
        EXPECT_STREQ(xml, obj.to_xml().c_str());
    }

    TEST_F(AttachmentTest, DefaultConstructorCreatesItemAttachment)
    {
        auto attachment = ews::attachment();
        EXPECT_EQ(ews::attachment::type::item, attachment.get_type());
    }

    TEST_F(AttachmentTest, CreateFromExistingItem)
    {
        auto& message = test_message();
        auto att = ews::attachment::from_item(message, "Arbitrary name");

        EXPECT_EQ(ews::attachment::type::item, att.get_type());
        EXPECT_FALSE(att.id().valid());
        EXPECT_STREQ("Arbitrary name", att.name().c_str());
        EXPECT_STREQ("", att.content_type().c_str());
        EXPECT_TRUE(att.content().empty());
        EXPECT_EQ(0U, att.content_size());
    }

    TEST_F(AttachmentTest, ToXML)
    {
        auto item = ews::item();
        auto item_attachment = ews::attachment::from_item(item, "Some name");
        const auto xml = item_attachment.to_xml();
        EXPECT_FALSE(xml.empty());
        EXPECT_STREQ(
            "<t:ItemAttachment><t:Name>Some name</t:Name></t:ItemAttachment>",
            xml.c_str());
    }

    // TODO: CreateFromItem
    // TODO: GetAttachmentWithInvalidIdThrows
    // TODO: GetAttachmentWithInvalidIdExceptionResponse

    TEST_F(AttachmentTest, CreateAndDeleteItemAttachmentOnServer)
    {
        // Try to attach an existing message to a new task

        auto& msg = test_message();
        auto item_attachment = ews::attachment::from_item(msg, "This message");

        auto some_task = ews::task();
        some_task.set_subject("Respond to Mike's mail!");
        auto task_id = service().create_item(some_task);
        some_task = service().get_task(task_id);
        ews::internal::on_scope_exit remove_task([&]
        {
            service().delete_task(std::move(some_task));
        });

        EXPECT_TRUE(some_task.get_attachments().empty());

        auto attachment_id = service().create_attachment(
                                                        some_task.get_item_id(),
                                                        item_attachment);
        ASSERT_TRUE(attachment_id.valid());
        item_attachment = service().get_attachment(attachment_id);
        ews::internal::on_scope_exit remove_attachment([&]
        {
            service().delete_attachment(item_attachment.id());
        });

        // RootItemId should be that of the parent task

        auto root_item_id = attachment_id.root_item_id();
        ASSERT_TRUE(root_item_id.valid());
        EXPECT_FALSE(root_item_id.change_key().empty());

        EXPECT_STREQ(root_item_id.id().c_str(), task_id.id().c_str());

        // Finally, check item::get_attachments

        some_task = service().get_task(task_id);
        auto attachments = some_task.get_attachments();
        ASSERT_EQ(1U, attachments.size());
        EXPECT_STREQ("This message", attachments[0].name().c_str());
    }

#ifdef EWS_USE_BOOST_LIBRARY
    TEST_F(AssetsFixture, ItemAttachmentFromXML)
    {
        using ews::internal::get_element_by_qname;

        std::vector<char> buf = read_contents(
            assets_dir() / "get_attachment_response_item.xml");
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = get_element_by_qname(doc,
                                        "ItemAttachment",
                                        ews::internal::uri<>::microsoft::types());
        ASSERT_TRUE(node);
        auto att = ews::attachment::from_xml_element(*node);

        EXPECT_EQ(ews::attachment::type::item, att.get_type());
        EXPECT_TRUE(att.id().valid());
        EXPECT_STREQ("This message", att.name().c_str());
        EXPECT_STREQ("", att.content_type().c_str());
        EXPECT_EQ(0U, att.content_size());

        std::string expected_xml;
        rapidxml::print(std::back_inserter(expected_xml),
                        *node,
                        rapidxml::print_no_indenting);
        const auto actual_xml = att.to_xml();
        EXPECT_STREQ(expected_xml.c_str(), actual_xml.c_str());
    }

    TEST_F(AssetsFixture, ToXML)
    {
        const auto path = assets_dir() / "ballmer_peak.png";
        auto attachment = ews::attachment::from_file(path.string(),
                                                     "image/png",
                                                     "Ballmer Peak");
        const auto xml = attachment.to_xml();
        EXPECT_FALSE(xml.empty());
        const char* prefix = "<t:FileAttachment>";
        EXPECT_TRUE(std::equal(prefix, prefix + std::strlen(prefix),
                    begin(xml)));
    }

    TEST_F(AssetsFixture, WriteContentToFileDoesNothingIfItemAttachment)
    {
        const auto target_path = cwd() / "output.bin";
        auto item = ews::item();
        auto item_attachment = ews::attachment::from_item(item, "Some name");
        const auto bytes_written =
            item_attachment.write_content_to_file(target_path.string());
        EXPECT_EQ(0U, bytes_written);
        EXPECT_FALSE(boost::filesystem::exists(target_path));
    }

    TEST_F(AssetsFixture, WriteContentToFile)
    {
        using uri = ews::internal::uri<>;
        using ews::internal::get_element_by_qname;
        using ews::internal::on_scope_exit;

        const auto target_path = cwd() / "output.png";
        std::vector<char> buf = read_contents(
                assets_dir() / "get_attachment_response.xml");
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = get_element_by_qname(doc,
                                         "FileAttachment",
                                         uri::microsoft::types());
        ASSERT_TRUE(node);
        auto attachment = ews::attachment::from_xml_element(*node);
        const auto bytes_written =
            attachment.write_content_to_file(target_path.string());
        on_scope_exit remove_file([&]
        {
            boost::filesystem::remove(target_path);
        });
        EXPECT_EQ(93525U, bytes_written);
        EXPECT_TRUE(boost::filesystem::exists(target_path));
    }

    TEST_F(AssetsFixture, WriteContentToFileThrowsOnEmptyFileName)
    {
        const auto path = assets_dir() / "ballmer_peak.png";
        auto attachment = ews::attachment::from_file(path.string(),
                                                     "image/png",
                                                     "Ballmer Peak");
        EXPECT_THROW(
        {
            attachment.write_content_to_file("");

        }, ews::exception);
    }

    TEST_F(AssetsFixture, WriteContentToFileExceptionMessage)
    {
        const auto path = assets_dir() / "ballmer_peak.png";
        auto attachment = ews::attachment::from_file(path.string(),
                                                     "image/png",
                                                     "Ballmer Peak");
        try
        {
            attachment.write_content_to_file("");
            FAIL() << "Expected exception to be raised";
        }
        catch (ews::exception& exc)
        {
            EXPECT_STREQ("Could not open file for writing: no file name given",
                         exc.what());
        }
    }

    TEST_F(AssetsFixture, CreateFromFile)
    {
        const auto path = assets_dir() / "ballmer_peak.png";
        auto file_attachment = ews::attachment::from_file(path.string(),
                                                          "image/png",
                                                          "Ballmer Peak");
        EXPECT_EQ(ews::attachment::type::file, file_attachment.get_type());
        EXPECT_FALSE(file_attachment.id().valid());
        EXPECT_STREQ("Ballmer Peak", file_attachment.name().c_str());
        EXPECT_STREQ("image/png", file_attachment.content_type().c_str());
        EXPECT_FALSE(file_attachment.content().empty());
        EXPECT_EQ(93525U, file_attachment.content_size());
    }

    TEST_F(AssetsFixture, CreateFromFileThrowsIfFileDoesNotExists)
    {
        const auto path = assets_dir() / "unlikely_to_exist.txt";
        EXPECT_THROW(
        {
            ews::attachment::from_file(path.string(), "image/png", "");

        }, ews::exception);
    }

    TEST_F(AssetsFixture, CreateFromFileExceptionMessage)
    {
        const auto path = assets_dir() / "unlikely_to_exist.txt";
        try
        {
            ews::attachment::from_file(path.string(), "image/png", "");
            FAIL() << "Expected exception to be raised";
        }
        catch (ews::exception& exc)
        {
            auto expected = "Could not open file for reading: " + path.string();
            EXPECT_STREQ(expected.c_str(), exc.what());
        }
    }

    TEST_F(AssetsFixture, FileAttachmentFromXML)
    {
        using ews::internal::get_element_by_qname;

        std::vector<char> buf = read_contents(
                assets_dir() / "get_attachment_response.xml");
        buf.push_back('\0');
        xml_document doc;
        doc.parse<0>(&buf[0]);
        auto node = get_element_by_qname(doc,
                                         "FileAttachment",
                                         ews::internal::uri<>::microsoft::types());
        ASSERT_TRUE(node);
        auto obj = ews::attachment::from_xml_element(*node);

        EXPECT_EQ(ews::attachment::type::file, obj.get_type());
        EXPECT_TRUE(obj.id().valid());
        EXPECT_STREQ("ballmer_peak.png", obj.name().c_str());
        EXPECT_STREQ("image/png", obj.content_type().c_str());
        EXPECT_EQ(0U, obj.content_size());
    }

    class FileAttachmentTest : public AttachmentTest
    {
    public:
        FileAttachmentTest()
            : assetsdir_(ews::test::global_data::instance().assets_dir)
        {
        }

        void SetUp()
        {
            AttachmentTest::SetUp();

            olddir_ = boost::filesystem::current_path();
            workingdir_ = boost::filesystem::unique_path(
                boost::filesystem::temp_directory_path() /
                "%%%%-%%%%-%%%%-%%%%");
            ASSERT_TRUE(boost::filesystem::create_directory(workingdir_))
                << "Unable to create temporary working directory";
            boost::filesystem::current_path(workingdir_);
        }

        void TearDown()
        {
            EXPECT_TRUE(boost::filesystem::is_empty(workingdir_))
                << "Temporary directory not empty on TearDown";
            boost::filesystem::current_path(olddir_);
            boost::filesystem::remove_all(workingdir_);

            AttachmentTest::TearDown();
        }

        const boost::filesystem::path& assets_dir() const
        {
            return assetsdir_;
        }

    private:
        boost::filesystem::path assetsdir_;
        boost::filesystem::path olddir_;
        boost::filesystem::path workingdir_;
    };

    TEST_F(FileAttachmentTest, CreateAndDeleteFileAttachmentOnServer)
    {
        auto& msg = test_message();
        const auto path = assets_dir() / "ballmer_peak.png";

        auto file_attachment = ews::attachment::from_file(path.string(),
                                                          "image/png",
                                                          "Ballmer Peak");

        // Attach image to email message
        auto attachment_id = service().create_attachment(msg.get_item_id(),
                                                         file_attachment);
        ASSERT_TRUE(attachment_id.valid());

        // Make sure two additional attributes of <AttachmentId> are set; only
        // <CreateAttachment> call returns them
        EXPECT_FALSE(attachment_id.root_item_id().id().empty());
        EXPECT_FALSE(attachment_id.root_item_id().change_key().empty());

        file_attachment = service().get_attachment(attachment_id);

        // Test if properties are as expected
        EXPECT_EQ(ews::attachment::type::file, file_attachment.get_type());
        EXPECT_TRUE(file_attachment.id().valid());
        EXPECT_STREQ("Ballmer Peak", file_attachment.name().c_str());
        EXPECT_STREQ("image/png", file_attachment.content_type().c_str());
        EXPECT_FALSE(file_attachment.content().empty());

        ASSERT_NO_THROW(
        {
            service().delete_attachment(file_attachment.id());
        });

        // Check if it is still in store
        EXPECT_THROW(
        {
            service().get_attachment(attachment_id);

        }, ews::exchange_error);
    }
#endif // EWS_USE_BOOST_LIBRARY
}

// vim:et ts=4 sw=4 noic cc=80
