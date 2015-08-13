#pragma once

#ifdef EWS_HAS_NOEXCEPT_SPECIFIER
# define EWS_NOEXCEPT noexcept
#else
# define EWS_NOEXCEPT
#endif

// Forward declarations
namespace ews
{
    class attachment;
    class attachment_id;
    template <typename T> class basic_service;
    class body;
    class contact;
    class date_time;
    class distinguished_folder_id;
    class email_address;
    class exception;
    class exchange_error;
    class folder_id;
    class http_error;
    class indexed_property_path;
    class is_equal_to;
    class item;
    class item_id;
    class message;
    class mime_content;
    class parse_error;
    class property;
    class property_path;
    class restriction;
    class schema_validation_error;
    class soap_fault;
    class task;

    bool operator==(const date_time&, const date_time&);
    bool operator==(const property_path&, const property_path&);
    void set_up() EWS_NOEXCEPT;
    void tear_down() EWS_NOEXCEPT;
}
