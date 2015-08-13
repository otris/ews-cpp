#include "fixtures.hpp"

#include <vector>
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

#ifdef EWS_USE_BOOST_LIBRARY
    TEST_F(FileAttachmentTest, CreateAndDeleteAttachmentOnServer)
    {
        auto& msg = test_message();
        const auto path = assets_dir() / "ballmer_peak.png";

        auto file_attachment = ews::attachment::from_file(path.string(),
                                                          "image/png",
                                                          "Ballmer Peak");

        // Attach image to email message
        auto attachment_id = service().create_attachment(msg, file_attachment);
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
        EXPECT_EQ(93525U, file_attachment.content_size());

        ASSERT_NO_THROW(
        {
            service().delete_attachment(std::move(file_attachment));
        });

        // Test sink argument
        EXPECT_FALSE(file_attachment.id().valid());

        // Check if it is still in store
        EXPECT_THROW(
        {
            service().get_attachment(attachment_id);

        }, ews::exchange_error);
    }

    TEST_F(FileAttachmentTest, ToXML)
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

    TEST_F(FileAttachmentTest, WriteContentToFileDoesNothingIfItemAttachment)
    {
        const auto target_path = cwd() / "output.bin";
        auto item = ews::item();
        auto item_attachment = ews::attachment::from_item(item, "Some name");
        const auto bytes_written =
            item_attachment.write_content_to_file(target_path.string());
        EXPECT_EQ(0U, bytes_written);
        EXPECT_FALSE(boost::filesystem::exists(target_path));
    }

    TEST_F(FileAttachmentTest, WriteContentToFile)
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

    TEST_F(FileAttachmentTest, WriteContentToFileThrowsOnEmptyFileName)
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

    TEST_F(FileAttachmentTest, WriteContentToFileExceptionMessage)
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

    TEST_F(FileAttachmentTest, CreateFromFile)
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

    TEST_F(FileAttachmentTest, CreateFromFileThrowsIfFileDoesNotExists)
    {
        const auto path = assets_dir() / "unlikely_to_exist.txt";
        EXPECT_THROW(
        {
            ews::attachment::from_file(path.string(), "image/png", "");

        }, ews::exception);
    }

    TEST_F(FileAttachmentTest, CreateFromFileExceptionMessage)
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

    TEST_F(FileAttachmentTest, FileAttachmentFromXML)
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
#endif // EWS_USE_BOOST_LIBRARY
}
