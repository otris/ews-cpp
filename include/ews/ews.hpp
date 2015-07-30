#pragma once

#include <exception>
#include <stdexcept>
#include <string>
#include <vector>
#include <unordered_map>
#include <sstream>
#include <fstream>
#include <ios>
#include <algorithm>
#include <iterator>
#include <functional>
#include <utility>
#include <memory>
#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
# include <type_traits>
#endif
#include <cstddef>
#include <cstring>
#include <cctype>

#include <curl/curl.h>

#include "rapidxml/rapidxml.hpp"
#include "rapidxml/rapidxml_print.hpp"

// Print more detailed error messages, HTTP request/response etc to stderr
#ifdef EWS_ENABLE_VERBOSE
# include <iostream>
# include <ostream>
#endif

// Macro for verifying expressions at run-time. Calls assert() with 'expr'.
// Allows turning assertions off, even if -DNDEBUG wasn't given at
// compile-time.  This macro does nothing unless EWS_ENABLE_ASSERTS was defined
// during compilation.
#define EWS_ASSERT(expr) ((void)0)
#ifndef NDEBUG
# ifdef EWS_ENABLE_ASSERTS
#  include <cassert>
#  undef EWS_ASSERT
#  define EWS_ASSERT(expr) assert(expr)
# endif
#endif // !NDEBUG

#ifdef EWS_HAS_NOEXCEPT_SPECIFIER
# define EWS_NOEXCEPT noexcept
#else
# define EWS_NOEXCEPT
#endif

namespace ews
{
    namespace internal
    {
        // Scope guard helper.
        class on_scope_exit final
        {
        public:
            template <typename Function>
            on_scope_exit(Function destructor_function) try
                : func_(std::move(destructor_function))
            {}
            catch (std::exception&)
            {
                destructor_function();
            }

            on_scope_exit() = delete;
            on_scope_exit(const on_scope_exit&) = delete;
            on_scope_exit& operator=(const on_scope_exit&) = delete;

            ~on_scope_exit()
            {
                if (func_)
                {
                    try
                    {
                        func_();
                    }
                    catch (std::exception&)
                    {
                        // Swallow, abort(3) if suitable
                        EWS_ASSERT(false);
                    }
                }
            }

            void release() { func_ = nullptr; }

        private:
            std::function<void(void)> func_;
        };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
        static_assert(!std::is_copy_constructible<on_scope_exit>::value, "");
        static_assert(!std::is_copy_assignable<on_scope_exit>::value, "");
        static_assert(!std::is_default_constructible<on_scope_exit>::value, "");
#endif

        // Helper functor; calculate hash of 'enum class'
        struct enum_class_hash
        {
            template <typename T> std::size_t operator()(T val) const
            {
                return static_cast<std::size_t>(val);
            }
        };

        // Helper class; makes sub-classes not copy assignable. Prevents MSVC++
        // from trying (and miserably failing) at generating those functions
        struct no_assign
        {
            no_assign& operator=(const no_assign&) = delete;
        };

        namespace base64
        {
            // Following code (everything in base64 namespace) is a slightly
            // modified version of the original implementation from
            // René Nyffenegger available at
            //
            //     http://www.adp-gmbh.ch/cpp/common/base64.html
            //
            // Copyright (C) 2004-2008 René Nyffenegger
            //
            // This source code is provided 'as-is', without any express or
            // implied warranty. In no event will the author be held liable for
            // any damages arising from the use of this software.
            //
            // Permission is granted to anyone to use this software for any
            // purpose, including commercial applications, and to alter it and
            // redistribute it freely, subject to the following restrictions:
            //
            // 1. The origin of this source code must not be misrepresented;
            //    you must not claim that you wrote the original source code.
            //    If you use this source code in a product, an acknowledgment
            //    in the product documentation would be appreciated but is not
            //    required.
            //
            // 2. Altered source versions must be plainly marked as such, and
            //    must not be misrepresented as being the original source code.
            //
            // 3. This notice may not be removed or altered from any source
            //    distribution.
            //
            // René Nyffenegger rene.nyffenegger@adp-gmbh.ch

            inline const std::string& valid_chars()
            {
                static const auto chars =
                    std::string("ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                                "abcdefghijklmnopqrstuvwxyz"
                                "0123456789+/");
                return chars;
            }

            inline bool is_base64(unsigned char c) EWS_NOEXCEPT
            {
                return std::isalnum(c) || (c == '+') || (c == '/');
            }

            inline std::string encode(const std::vector<unsigned char>& buf)
            {
                const auto& base64_chars = valid_chars();
                int i = 0;
                int j = 0;
                unsigned char char_array_3[3];
                unsigned char char_array_4[4];
                auto buflen = buf.size();
                auto bufit = begin(buf);
                std::string ret;

                while (buflen--)
                {
                    char_array_3[i++] = *(bufit++);
                    if (i == 3)
                    {
                        char_array_4[0] =  (char_array_3[0] & 0xfc) >> 2;
                        char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
                        char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
                        char_array_4[3] =   char_array_3[2] & 0x3f;

                        for(i = 0; (i < 4); i++)
                        {
                            ret += base64_chars[char_array_4[i]];
                        }
                        i = 0;
                    }
                }

                if (i)
                {
                    for(j = i; j < 3; j++)
                    {
                        char_array_3[j] = '\0';
                    }

                    char_array_4[0] =  (char_array_3[0] & 0xfc) >> 2;
                    char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
                    char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
                    char_array_4[3] =   char_array_3[2] & 0x3f;

                    for (j = 0; (j < i + 1); j++)
                    {
                        ret += base64_chars[char_array_4[j]];
                    }

                    while((i++ < 3))
                    {
                        ret += '=';
                    }
                }

                return ret;
            }

            inline
            std::vector<unsigned char> decode(const std::string& encoded_string)
            {
                const auto& base64_chars = valid_chars();
                int in_len = encoded_string.size();
                int i = 0;
                int j = 0;
                int in = 0;
                unsigned char char_array_4[4];
                unsigned char char_array_3[3];
                std::vector<unsigned char> ret;

                while (   in_len--
                       && (encoded_string[in] != '=')
                       && is_base64(encoded_string[in]))
                {
                    char_array_4[i++] = encoded_string[in]; in++;
                    if (i == 4)
                    {
                      for (i = 0; i < 4; i++)
                      {
                          char_array_4[i] = base64_chars.find(char_array_4[i]);
                      }

                      char_array_3[0] =  (char_array_4[0] << 2)        + ((char_array_4[1] & 0x30) >> 4);
                      char_array_3[1] = ((char_array_4[1] & 0xf) << 4) + ((char_array_4[2] & 0x3c) >> 2);
                      char_array_3[2] = ((char_array_4[2] & 0x3) << 6) +   char_array_4[3];

                      for (i = 0; (i < 3); i++)
                      {
                          ret.push_back(char_array_3[i]);
                      }
                      i = 0;
                    }
                }

                if (i)
                {
                    for (j = i; j < 4; j++)
                    {
                        char_array_4[j] = 0;
                    }

                    for (j = 0; j < 4; j++)
                    {
                        char_array_4[j] = base64_chars.find(char_array_4[j]);
                    }

                    char_array_3[0] =  (char_array_4[0] << 2)        + ((char_array_4[1] & 0x30) >> 4);
                    char_array_3[1] = ((char_array_4[1] & 0xf) << 4) + ((char_array_4[2] & 0x3c) >> 2);
                    char_array_3[2] = ((char_array_4[2] & 0x3) << 6) +   char_array_4[3];

                    for (j = 0; (j < i - 1); j++)
                    {
                        ret.push_back(char_array_3[j]);
                    }
                }

                return ret;
            }

        } // base64

        // Forward declarations
        class http_request;
    }

    // Forward declarations
    class item_id;

    enum class response_class { error, success, warning };

    // Base-class for all exceptions thrown by this library
    class exception : public std::runtime_error
    {
    public:
        explicit exception(const std::string& what)
            : std::runtime_error(what)
        {
        }
        explicit exception(const char* what) : std::runtime_error(what) {}
    };

    // Raised when a response from a server could not be parsed
    class parse_error final : public exception
    {
    public:
        explicit parse_error(const std::string& what)
            : exception(what)
        {
        }
        explicit parse_error(const char* what) : exception(what) {}
    };

    enum class response_code
    {
        no_error,

        // Calling account does not have the rights to perform the action
        // requested.
        error_access_denied,

        // Returned when the account in question has been disabled.
        error_account_disabled,

        // The address space (Domain Name System [DNS] domain name) record for
        // cross forest availability could not be found in the Microsoft Active
        // Directory
        error_address_space_not_found,

        // Operation failed due to issues talking with the Active Directory.
        error_ad_operation,

        // You should never encounter this response code, which occurs only as a
        // result of an issue in Exchange Web Services.
        error_ad_session_filter,

        // The Active Directory is temporarily unavailable. Try your request
        // again later.
        error_ad_unavailable,

        // Indicates that Exchange Web Services tried to determine the URL of a
        // cross forest Client Access Server (CAS) by using the AutoDiscover
        // service, but the call to AutoDiscover failed.
        error_auto_discover_failed,

        // The AffectedTaskOccurrences enumeration value is missing. It is
        // required when deleting a task so that Exchange Web Services knows
        // whether you want to delete a single task or all occurrences of a
        // repeating task.
        error_affected_task_occurrences_required,

        // You encounter this error only if the size of your attachment exceeds
        // Int32. MaxValue (in bytes). Of course, you probably have bigger
        // problems if that is the case.
        error_attachment_size_limit_exceeded,

        // The availability configuration information for the local Active
        // Directory forest is missing from the Active Directory.
        error_availability_config_not_found,

        // Indicates that the previous item in the request failed in such a way
        // that Exchange Web Services stopped processing the remaining items in
        // the request. All remaining items are marked with
        // ErrorBatchProcessingStopped.
        error_batch_processing_stopped,

        // You are not allowed to move or copy calendar item occurrences.
        error_calendar_cannot_move_or_copy_occurrence,

        // If the update in question would send out a meeting update, but the
        // meeting is in the organizer's deleted items folder, then you
        // encounter this error.
        error_calendar_cannot_update_deleted_item,

        // When a RecurringMasterId is examined, the OccurrenceId attribute is
        // examined to see if it refers to a valid occurrence. If it doesn't,
        // then this response code is returned.
        error_calendar_cannot_use_id_for_occurrence_id,

        // When an OccurrenceId is examined, the RecurringMasterId attribute is
        // examined to see if it refers to a valid recurring master. If it
        // doesn't, then this response code is returned.
        error_calendar_cannot_use_id_for_recurring_master_id,

        // Returned if the suggested duration of a calendar item exceeds five
        // years.
        error_calendar_duration_is_too_long,

        // The end date must be greater than the start date. Otherwise, it
        // isn't worth having the meeting.
        error_calendar_end_date_is_earlier_than_start_date,

        // You can apply calendar paging only on a CalendarFolder.
        error_calendar_folder_is_invalid_for_calendar_view,

        // You should never encounter this response code.
        error_calendar_invalid_attribute_value,

        // When defining a time change pattern, the values Day, WeekDay and
        // WeekendDay are invalid.
        error_calendar_invalid_day_for_time_change_pattern,

        // When defining a weekly recurring pattern, the values Day, Weekday,
        // and WeekendDay are invalid.
        error_calendar_invalid_day_for_weekly_recurrence,

        // Indicates that the state of the calendar item recurrence blob in the
        // Exchange Store is invalid.
        error_calendar_invalid_property_state,

        // You should never encounter this response code.
        error_calendar_invalid_property_value,

        // You should never encounter this response code. It indicates that the
        // internal structure of the objects representing the recurrence is
        // invalid.
        error_calendar_invalid_recurrence,

        // Occurs when an invalid time zone is encountered.
        error_calendar_invalid_time_zone,

        // Accepting a meeting request by using delegate access is not supported
        // in RTM.
        error_calendar_is_delegated_for_accept,

        // Declining a meeting request by using delegate access is not supported
        // in RTM.
        error_calendar_is_delegated_for_decline,

        // Removing an item from the calendar by using delegate access is not
        // supported in RTM.
        error_calendar_is_delegated_for_remove,

        // Tentatively accepting a meeting request by using delegate access is
        // not supported in RTM.
        error_calendar_is_delegated_for_tentative,

        // Only the meeting organizer can cancel the meeting, no matter how much
        // you are dreading it.
        error_calendar_is_not_organizer,

        // The organizer cannot accept the meeting. Only attendees can accept a
        // meeting request.
        error_calendar_is_organizer_for_accept,

        // The organizer cannot decline the meeting. Only attendees can decline
        // a meeting request.
        error_calendar_is_organizer_for_decline,

        // The organizer cannot remove the meeting from the calendar by using
        // RemoveItem. The organizer can do this only by cancelling the meeting
        // request. Only attendees can remove a calendar item.
        error_calendar_is_organizer_for_remove,

        // The organizer cannot tentatively accept the meeting request. Only
        // attendees can tentatively accept a meeting request.
        error_calendar_is_organizer_for_tentative,

        // Occurs when the occurrence index specified in the OccurenceId does
        // not correspond to a valid occurrence. For instance, if your
        // recurrence pattern defines a set of three meeting occurrences, and
        // you try to access the fifth occurrence, well, there is no fifth
        // occurrence. So instead, you get this response code.
        error_calendar_occurrence_index_is_out_of_recurrence_range,

        // Occurs when the occurrence index specified in the OccurrenceId
        // corresponds to a deleted occurrence.
        error_calendar_occurrence_is_deleted_from_recurrence,

        // Occurs when a recurrence pattern is defined that contains values for
        // month, day, week, and so on that are out of range. For example,
        // specifying the fifteenth week of the month is both silly and an
        // error.
        error_calendar_out_of_range,

        // Calendar paging can span a maximum of two years.
        error_calendar_view_range_too_big,

        // Calendar items can be created only within calendar folders.
        error_cannot_create_calendar_item_in_non_calendar_folder,

        // Contacts can be created only within contact folders.
        error_cannot_create_contact_in_non_contacts_folder,

        // Tasks can be created only within Task folders.
        error_cannot_create_task_in_non_task_folder,

        // Occurs when Exchange Web Services fails to delete the item or folder
        // in question for some unexpected reason.
        error_cannot_delete_object,

        // This error indicates that you either tried to delete an occurrence of
        // a nonrecurring task or tried to delete the last occurrence of a
        // recurring task.
        error_cannot_delete_task_occurrence,

        // Exchange Web Services could not open the file attachment
        error_cannot_open_file_attachment,

        // The Id that was passed represents a folder rather than an item
        error_cannot_use_folder_id_for_item_id,

        // The id that was passed in represents an item rather than a folder
        error_cannot_user_item_id_for_folder_id,

        // You will never encounter this response code. Superseded by
        // ErrorChangeKeyRequiredForWriteOperations.
        error_change_key_required,

        // When performing certain update operations, you must provide a valid
        // change key.
        error_change_key_required_for_write_operations,

        // This response code is returned when Exchange Web Services is unable
        // to connect to the Mailbox in question.
        error_connection_failed,

        // Occurs when Exchange Web Services is unable to retrieve the MIME
        // content for the item in question (GetItem), or is unable to create
        // the item from the supplied MIME content (CreateItem).
        error_content_conversion_failed,

        // Indicates that the data in question is corrupt and cannot be acted
        // upon.
        error_corrupt_data,

        // Indicates that the caller does not have the right to create the item
        // in question.
        error_create_item_access_denied,

        // Indicates that one ore more of the managed folders passed to
        // CreateManagedFolder failed to be created. The only way to determine
        // which ones failed is to search for the folders and see which ones do
        // not exist.
        error_create_managed_folder_partial_completion,

        // The calling account does not have the proper rights to create the
        // subfolder in question.
        error_create_subfolder_access_denied,

        // You cannot move an item or folder from one Mailbox to another.
        error_cross_mailbox_move_copy,

        // Either the data that you were trying to set exceeded the maximum size
        // for the property, or the value is large enough to require streaming
        // and the property does not support streaming (that is, folder
        // properties).
        error_data_size_limit_exceeded,

        // An Active Directory operation failed.
        error_data_source_operation,

        // You cannot delete a distinguished folder
        error_delete_distinguished_folder,

        // You will never encounter this response code.
        error_delete_items_failed,

        // There are duplicate values in the folder names array that was passed
        // into CreateManagedFolder.
        error_duplicate_input_folder_names,

        // The Mailbox subelement of DistinguishedFolderId pointed to a
        // different Mailbox than the one you are currently operating on. For
        // example, you cannot create a search folder that exists in one Mailbox
        // but considers distinguished folders from another Mailbox in its
        // search criteria.
        error_email_address_mismatch,

        // Indicates that the subscription that was created with a particular
        // watermark is no longer valid.
        error_event_not_found,

        // Indicates that the subscription referenced by GetEvents has expired.
        error_expired_subscription,

        // The folder is corrupt and cannot be saved. This means that you set
        // some strange and invalid property on the folder, or possibly that the
        // folder was already in some strange state before you tried to set
        // values on it (UpdateFolder). In any case, this is not a common error.
        error_folder_corrupt,

        // Indicates that the folder id passed in does not correspond to a valid
        // folder, or in delegate access cases that the delegate does not have
        // permissions to access the folder.
        error_folder_not_found,

        // Indicates that the property that was requested could not be
        // retrieved. Note that this does not mean that it just wasn't there.
        // This means that the property was corrupt in some way such that
        // retrieving it failed. This is not a common error.
        error_folder_property_request_failed,

        // The folder could not be created or saved due to invalid state.
        error_folder_save,

        // The folder could not be created or saved due to invalid state.
        error_folder_save_failed,

        // The folder could not be created or updated due to invalid property
        // values. The properties which caused the problem are listed in the
        // MessageXml element..
        error_folder_save_property_error,

        // A folder with that name already exists.
        error_folder_exists,

        // Unable to retrieve Free/Busy information.This should not be common.
        error_free_busy_generation_failed,

        // You will never encounter this response code.
        error_get_server_security_descriptor_failed,

        // This response code is always returned within a SOAP fault. It
        // indicates that the calling account does not have the ms-Exch-EPI-May-
        // Impersonate right on either the user/contact they are try to
        // impersonate or the Mailbox database containing the user Mailbox.
        error_impersonate_user_denied,

        // This response code is always returned within a SOAP fault. It
        // indicates that the calling account does not have the ms-Exch-EPI-
        // Impersonation right on the CAS it is calling.
        error_impersonation_denied,

        // There was an unexpected error trying to perform Server to Server
        // authentication. This typically indicates that the service account
        // running the Exchange Web Services application pool is misconfigured,
        // that Exchange Web Services cannot talk to the Active Directory, or
        // that a trust between Active Directory forests is not properly
        // configured.
        error_impersonation_failed,

        // Each change description within an UpdateItem or UpdateFolder call
        // must list one and only one property to update.
        error_incorrect_update_property_count,

        // Your request contained too many attendees to resolve. The default
        // mailbox count limit is 100.
        error_individual_mailbox_limit_reached,

        // Indicates that the Mailbox server is overloaded You should try your
        // request again later.
        error_insufficient_resources,

        // This response code means that the Exchange Web Services team members
        // didn't anticipate all possible scenarios, and therefore Exchange
        // Web Services encountered a condition that it couldn't recover from.
        error_internal_server_error,

        // This response code is an interesting one. It essentially means that
        // the Exchange Web Services team members didn't anticipate all
        // possible scenarios, but the nature of the unexpected condition is
        // such that it is likely temporary and so you should try again later.
        error_internal_server_transient_error,

        // It is unlikely that you will encounter this response code. It means
        // that Exchange Web Services tried to figure out what level of access
        // the caller has on the Free/Busy information of another account, but
        // the access that was returned didn't make sense.
        error_invalid_access_level,

        // Indicates that the attachment was not found within the attachments
        // collection on the item in question. You might encounter this if you
        // have an attachment id, the attachment is deleted, and then you try to
        // call GetAttachment on the old attachment id.
        error_invalid_attachment_id,

        // Exchange Web Services supports only simple contains filters against
        // the attachment table. If you try to retrieve the search parameters on
        // an existing search folder that has a more complex attachment table
        // restriction (called a subfilter), then Exchange Web Services has no
        // way of rendering the XML for that filter, and it returns this
        // response code. Note that you can still call GetFolder on this folder,
        // just don't request the SearchParameters property.
        error_invalid_attachment_subfilter,

        // Exchange Web Services supports only simple contains filters against
        // the attachment table. If you try to retrieve the search parameters on
        // an existing search folder that has a more complex attachment table
        // restriction, then Exchange Web Services has no way of rendering the
        // XML for that filter. This specific case means that the attachment
        // subfilter is a contains (text) filter, but the subfilter is not
        // referring to the attachment display name.
        error_invalid_attachment_subfilter_text_filter,

        // You should not encounter this error, which has to do with a failure
        // to proxy an availability request to another CAS
        error_invalid_authorization_context,

        // Indicates that the passed in change key was invalid. Note that many
        // methods do not require a change key to be passed. However, if you do
        // provide one, it must be a valid, though not necessarily up-to-date,
        // change key.
        error_invalid_change_key,

        // Indicates that there was an internal error related to trying to
        // resolve the caller's identity. This is not a common error.
        error_invalid_client_security_context,

        // Occurs when you try to set the CompleteDate of a task to a date in
        // the past. When converted to the local time of the CAS, the
        // CompleteDate cannot be set to a value less than the local time on the
        // CAS.
        error_invalid_complete_date,

        // This response code can be returned with two different error messages:
        // Unable to send cross-forest request for mailbox {mailbox identifier}
        // because of invalid configuration. When UseServiceAccount is
        // false, user name cannot be null or empty. You should never encounter
        // this response code.
        error_invalid_cross_forest_credentials,

        // An ExchangeImpersonation header was passed in but it did not contain
        // a security identifier (SID), user principal name (UPN) or primary
        // SMTP address. You must supply one of these identifiers and of course,
        // they cannot be empty strings. Note that this response code is always
        // returned within a SOAP fault.
        error_invalid_exchange_impersonation_header_data,

        // The bitmask passed into the Excludes restriction was unparsable.
        error_invalid_excludes_restriction,

        // You should never encounter this response code.
        error_invalid_expression_type_for_sub_filter,

        // The combination of extended property values that were specified is
        // invalid or results in a bad extended property URI.
        error_invalid_extended_property,

        // The value offered for the extended property is inconsistent with the
        // type specified in the associated extended field URI. For example, if
        // the PropertyType on the extended field URI is set to String, but you
        // set the value of the extended property as an array of integers, you
        // will encounter this error.
        error_invalid_extended_property_value,

        // You should never encounter this response code
        error_invalid_folder_id,

        // This response code will occur if: Numerator > denominator Numerator <
        // 0 Denominator <= 0
        error_invalid_fractional_paging_parameters,

        // Returned if you call GetUserAvailability with a FreeBusyViewType of
        // None
        error_invalid_free_busy_view_type,

        // Indicates that the Id and/or change key is malformed
        error_invalid_id,

        // Occurs if you pass in an empty id (Id="")
        error_invalid_id_empty,

        // Indicates that the Id is malformed.
        error_invalid_id_malformed,

        // Here is an example of over-engineering. Once again, this indicates
        // that the structure of the id is malformed The moniker referred to in
        // the name of this response code is contained within the id and
        // indicates which Mailbox the id belongs to. Exchange Web Services
        // checks the length of this moniker and fails it if the byte count is
        // more than expected. The check is good, but there is no reason to have
        // a separate response code for this condition.
        error_invalid_id_moniker_too_long,

        // The AttachmentId specified does not refer to an item attachment.
        error_invalid_id_not_an_item_attachment_id,

        // You should never encounter this response code. If you do, it
        // indicates that a contact in your Mailbox is so corrupt that Exchange
        // Web Services really can't tell the difference between it and a
        // waffle maker.
        error_invalid_id_returned_by_resolve_names,

        // Treat this like ErrorInvalidIdMalformed. Indicates that the id was
        // malformed
        error_invalid_id_store_object_id_too_long,

        // Exchange Web Services allows for attachment hierarchies to be a
        // maximum of 255 levels deep. If the attachment hierarchy on an item
        // exceeds 255 levels, you will get this response code.
        error_invalid_id_too_many_attachment_levels,

        // You will never encounter this response code.
        error_invalid_id_xml,

        // Indicates that the offset was < 0.
        error_invalid_indexed_paging_parameters,

        // You will never encounter this response code. At one point in the
        // Exchange Web Services development cycle, you could update the
        // Internet message headers via UpdateItem. Because that is no longer
        // the case, this response code is unused.
        error_invalid_internet_header_child_nodes,

        // Indicates that you tried to create an item attachment by using an
        // unsupported item type. Supported item types for item attachments are
        // Item, Message, CalendarItem, Contact, and Task. For instance, if you
        // try to create a MeetingMessage attachment, you encounter this
        // response code. In fact, the schema indicates that MeetingMessage,
        // MeetingRequest, MeetingResponse, and MeetingCancellation are all
        // permissible. Don't believe it.
        error_invalid_item_for_operation_create_item_attachment,

        // Indicates that you tried to create an unsupported item. In addition
        // to response objects, Supported items include Item, Message,
        // CalendarItem, Task, and Contact. For example, you cannot use
        // CreateItem to create a DistributionList. In addition, certain types
        // of items are created as a side effect of doing another action.
        // Meeting messages, for instance, are created as a result of sending a
        // calendar item to attendees. You never explicitly create a meeting
        // message.
        error_invalid_item_for_operation_create_item,

        // This response code is returned if: You created an AcceptItem response
        // object and referenced an item type other than a meeting request or a
        // calendar item. You tried to accept a calendar item occurrence that is
        // in the deleted items folder.
        error_invalid_item_for_operation_accept_item,

        // You created a CancelItem response object and referenced an item type
        // other than a calendar item.
        error_invalid_item_for_operation_cancel_item,

        // This response code is returned if: You created a DeclineItem response
        // object referencing an item with a type other than a meeting request
        // or a calendar item. You tried to decline a calendar item occurrence
        // that is in the deleted items folder.
        error_invalid_item_for_operation_decline_item,

        // The ItemId passed to ExpandDL does not represent a distribution list.
        // For example, you cannot expand a Message.
        error_invalid_item_for_operation_expand_dl,

        // You created a RemoveItem response object reference an item with a
        // type other than a meeting cancellation.
        error_invalid_item_for_operation_remove_item,

        // You tried to send an item with a type that does not derive from
        // MessageItem. Only items whose ItemClass begins with "IPM.Note"
        // are sendable
        error_invalid_item_for_operation_send_item,

        // This response code is returned if: You created a
        // TentativelyAcceptItem response object referencing an item whose type
        // is not a meeting request or a calendar item. You tried to tentatively
        // accept a calendar item occurrence that is in the deleted items
        // folder.
        error_invalid_item_for_operation_tentative,

        // Indicates that the structure of the managed folder is corrupt and
        // cannot be rendered.
        error_invalid_managed_folder_property,

        // Indicates that the quota that is set on the managed folder is less
        // than zero, indicating a corrupt managed folder.
        error_invalid_managed_folder_quota,

        // Indicates that the size that is set on the managed folder is less
        // than zero, indicating a corrupt managed folder.
        error_invalid_managed_folder_size,

        // Indicates that the supplied merged free/busy internal value is
        // invalid. Default minimum is 5 minutes. Default maximum is 1440
        // minutes.
        error_invalid_merged_free_busy_interval,

        // Indicates that the name passed into ResolveNames was invalid. For
        // instance, a zero length string, a single space, a comma, and a dash
        // are all invalid names. Vakue? Yes, that is part of the message text,
        // although it should obviously be "value."
        error_invalid_name_for_name_resolution,

        // Indicates that there is a problem with the NetworkService account on
        // the CAS. This response code is quite rare and has been seen only in
        // the wild by the most vigilant of hunters.
        error_invalid_network_service_context,

        // You will never encounter this response code.
        error_invalid_oof_parameter,

        // Indicates that you specified a MaxRows value that is <= 0.
        error_invalid_paging_max_rows,

        // You tried to create a folder within a search folder.
        error_invalid_parent_folder,

        // You tried to set the percentage complete property to an invalid value
        // (must be between 0 and 100 inclusive).
        error_invalid_percent_complete_value,

        // The property that you are trying to append to does not support
        // appending. Currently, the only properties that support appending are:
        // * Recipient collections (ToRecipients, CcRecipients, BccRecipients)
        // * Attendee collections (RequiredAttendees, OptionalAttendees,
        //   Resources)
        // * Body
        // * ReplyTo
        error_invalid_property_append,

        // The property that you are trying to delete does not support deleting.
        // An example of this is trying to delete the ItemId of an item.
        error_invalid_property_delete,

        // You cannot pass in a flags property to an Exists filter. The flags
        // properties are IsDraft, IsSubmitted, IsUnmodified, IsResend,
        // IsFromMe, and IsRead. Use IsEqualTo instead. The reason that flag
        // don't make sense in an Exists filter is that each of these flags is
        // actually a bit within a single property. So, calling Exists() on one
        // of these flags is like asking if a given bit exists within a value,
        // which is different than asking if that value exists or if the bit is
        // set. Likely you really mean to see if the bit is set, and you should
        // use the IsEqualTo filter expression instead.
        error_invalid_property_for_exists,

        // Indicates that the property you are trying to manipulate does not
        // support whatever operation your are trying to perform on it.
        error_invalid_property_for_operation,

        // Indicates that you requested a property in the response shape, and
        // that property is not within the schema of the item type in question.
        // For example, requesting calendar:OptionalAttendees in the response
        // shape of GetItem when binding to a message would result in this
        // error.
        error_invalid_property_request,

        // The property you are trying to set is read-only.
        error_invalid_property_set,

        // You cannot update a Message that has already been sent.
        error_invalid_property_update_sent_message,

        // You cannot call GetEvents or Unsubscribe on a push subscription id.
        // To unsubscribe from a push subscription, you must respond to a push
        // request with an unsubscribe response, or simply disconnect your Web
        // service and wait for the push notifications to time out.
        error_invalid_pull_subscription_id,

        // The URL provided as a callback for the push subscription has a bad
        // format. The following conditions must be met for Exchange Web
        // Services to accept the URL:
        // * String length > 0 and < 2083
        // * Protocol is HTTP or HTTPS
        // * Must be parsable by the System.Uri.NET Framework class
        error_invalid_push_subscription_url,

        // You should never encounter this response code. If you do, the
        // recipient collection on your message or attendee collection on your
        // calendar item is invalid.
        error_invalid_recipients,

        // Indicates that the search folder in question has a recipient table
        // filter that Exchange Web Services cannot represent. The response code
        // is a little misleading—there is nothing invalid about the search
        // folder restriction. To get around this issue, call GetFolder without
        // requesting the search parameters.
        error_invalid_recipient_subfilter,

        // Indicates that the search folder in question has a recipient table
        // filter that Exchange Web Services cannot represent. The error code is
        // a little misleading—there is nothing invalid about the search
        // folder restriction.. To get around this, issue, call GetFolder
        // without requesting the search parameters.
        error_invalid_recipient_subfilter_comparison,

        // Indicates that the search folder in question has a recipient table
        // filter that Exchange Web Services cannot represent. The response code
        // is a little misleading—there is nothing invalid about the search
        // folder restriction To get around this,issue, call GetFolder without
        // requesting the search parameters.
        error_invalid_recipient_subfilter_order,

        // Can you guess our comments on this one? Indicates that the search
        // folder in question has a recipient table filter that Exchange Web
        // Services cannot represent. The error code is a little
        // misleading—there is nothing invalid about the search folder
        // restriction. To get around this issue, call GetFolder without
        // requesting the search parameters.
        error_invalid_recipient_subfilter_text_filter,

        // You can only reply to/forward a Message, CalendarItem, or their
        // descendants. If you are referencing a CalendarItem and you are the
        // organizer, you can only forward the item. If you are referencing a
        // draft message, you cannot reply to the item. For read receipt
        // suppression, you can reference only a Message or descendant.
        error_invalid_reference_item,

        // Indicates that the SOAP request has a SOAP Action header, but nothing
        // in the SOAP body. Note that the SOAP Action header is not required
        // because Exchange Web Services can determine the method to call from
        // the local name of the root element in the SOAP body. However, you
        // must supply content in the SOAP body or else there is really nothing
        // for Exchange Web Services to act on..
        error_invalid_request,

        // You will never encounter this response code.
        error_invalid_restriction,

        // Indicates that the RoutingType element that was passed for an
        // EmailAddressType is not a valid routing type. Typically, routing type
        // will be set to Simple Mail Transfer Protocol (SMTP).
        error_invalid_routing_type,

        // The specified duration end time must be greater than the start time
        // and must occur in the future.
        error_invalid_scheduled_oof_duration,

        // Indicates that the security descriptor on the calendar folder in the
        // Store is corrupt.
        error_invalid_security_descriptor,

        // The SaveItemToFolder attribute is false, but you included a
        // SavedItemFolderId.
        error_invalid_send_item_save_settings,

        // Because you never use token serialization in your application, you
        // should never encounter this response code. The response code occurs
        // if the token passed in the SOAP header is malformed, doesn't refer
        // to a valid account in the Active Directory, or is missing the primary
        // group SID.
        error_invalid_serialized_access_token,

        // ExchangeImpersonation element have an invalid structure.
        error_invalid_sid,

        // The passed in SMTP address is not parsable.
        error_invalid_smtp_address,

        // You will never encounter this response code.
        error_invalid_subfilter_type,

        // You will never encounter this response code.
        error_invalid_subfilter_type_not_attendee_type,

        // You will never encounter this response code.
        error_invalid_subfilter_type_not_recipient_type,

        // Indicates that the subscription is no longer valid. This could be due
        // to the CAS having been rebooted or because the subscription has
        // expired.
        error_invalid_subscription,

        // Indicates that the sync state data is corrupt. You need to resync
        // without the sync state. Make sure that if you are persisting sync
        // state somewhere, you are not accidentally altering the information.
        error_invalid_sync_state_data,

        // The specified time interval is invalid (schema type Duration). The
        // start time must be greater than or equal to the end time.
        error_invalid_time_interval,

        // The user OOF settings are invalid due to a missing internal or
        // external reply.
        error_invalid_user_oof_settings,

        // Indicates that the UPN passed in the ExchangeImpersonation SOAP
        // header did not map to a valid account.
        error_invalid_user_principal_name,

        // Indicates that the SID passed in the ExchangeImpersonation SOAP
        // header was either invalid or did not map to a valid account.
        error_invalid_user_sid,

        // You will never encounter this response code.
        error_invalid_user_sid_missing_upn,

        // Indicates that the comparison value is invalid for the property your
        // are comparing against. For instance, DateTimeCreated > "true"
        // would cause this response code to be returned. You also encounter
        // this response code if you specify an enumeration property in the
        // comparison, but the value you are comparing against is not a valid
        // value for that enumeration.
        error_invalid_value_for_property,

        // Indicates that the supplied watermark is corrupt.
        error_invalid_watermark,

        // Indicates that conflict resolution was unable to resolve changes for
        // the properties in question. This typically means that someone changed
        // the item in the Store, and you are dealing with a stale copy.
        // Retrieve the updated change key and try again.
        error_irresolvable_conflict,

        // Indicates that the state of the object is corrupt and cannot be
        // retrieved. When retrieving an item, typically only certain properties
        // will be in this state (i.e. Body, MimeContent). Try omitting those
        // properties and retrying the operation.
        error_item_corrupt,

        // Indicates that the item in question was not found, or potentially
        // that it really does exist, and you just don't have rights to access
        // it.
        error_item_not_found,

        // Exchange Web Services tried to retrieve a given property on an item
        // or folder but failed to do so. Note that this means that some value
        // was there, but Exchange Web Services was unable to retrieve it.
        error_item_property_request_failed,

        // Attempts to save the item/folder failed.
        error_item_save,

        // Attempts to save the item/folder failed because of invalid property
        // values. The response includes the offending property paths.
        error_item_save_property_error,

        // You will never encounter this response code.
        error_legacy_mailbox_free_busy_view_type_not_merged,

        // You will never encounter this response code.
        error_local_server_object_not_found,

        // Indicates that the availability service was unable to log on as
        // Network Service to proxy requests to the appropriate sites/forests.
        // This typically indicates a configuration error.
        error_logon_as_network_service_failed,

        // Indicates that the Mailbox information in the Active Directory is
        // misconfigured.
        error_mailbox_configuration,

        // Indicates that the MailboxData array in the request is empty. You
        // must supply at least one Mailbox identifier.
        error_mailbox_data_array_empty,

        // You can supply a maximum of 100 entries in the MailboxData array.
        error_mailbox_data_array_too_big,

        // Failed to connect to the Mailbox to get the calendar view
        // information.
        error_mailbox_logon_failed,

        // The Mailbox in question is currently being moved. Try your request
        // again once the move is complete.
        error_mailbox_move_in_progress,

        // The Mailbox database is offline, corrupt, shutting down, or involved
        // in a plethora of other conditions that make the Mailbox unavailable.
        error_mailbox_store_unavailable,

        // Could not map the MailboxData information to a valid Mailbox account.
        error_mail_recipient_not_found,

        // The managed folder that you are trying to create already exists in
        // your Mailbox.
        error_managed_folder_already_exists,

        // The folder name specified in the request does not map to a managed
        // folder definition in the Active Directory. You can create instances
        // of managed folders only for folders defined in the Active Directory.
        // Check the name and try again.
        error_managed_folder_not_found,

        // Managed folders typically exist within the root managed folders
        // folder. The root must be properly created and functional in order to
        // deal with managed folders through Exchange Web Services. Typically,
        // this configuration happens transparently when you start dealing with
        // managed folders.
        // This response code indicates that the managed folders root was
        // deleted from the Mailbox or that there is already a folder in the
        // same parent folder with the name of the managed folder root. This
        // response code also occurs if the attempt to create the root managed
        // folder fails.
        error_managed_folders_root_failure,

        // Indicates that the suggestions engine encountered a problem when it
        // was trying to generate the suggestions.
        error_meeting_suggestion_generation_failed,

        // When creating or updating an item that descends from MessageType, you
        // must supply the MessageDisposition attribute on the request to
        // indicate what operations should be performed on the message. Note
        // that this attribute is not required for any other type of item.
        error_message_disposition_required,

        // Indicates that the message you are trying to send exceeds the
        // allowable limits.
        error_message_size_exceeded,

        // For CreateItem, the MIME content was not valid iCalendar content For
        // GetItem, the MIME content could not be generated.
        error_mime_content_conversion_failed,

        // The MIME content to set is invalid.
        error_mime_content_invalid,

        // The MIME content in the request is not a valid Base64 string.
        error_mime_content_invalid_base64_string,

        // A required argument was missing from the request. The response
        // message text indicates which argument to check.
        error_missing_argument,

        // Indicates that you specified a distinguished folder id in the
        // request, but the account that made the request does not have a
        // Mailbox on the system. In that case, you must supply a Mailbox
        // subelement under DistinguishedFolderId.
        error_missing_email_address,

        // This is really the same failure as ErrorMissingEmailAddress except
        // that ErrorMissingEmailAddressForManagedFolder is returned from
        // CreateManagedFolder.
        error_missing_email_address_for_managed_folder,

        // Indicates that the attendee or recipient does not have the
        // EmailAddress element set. You must set this element when making
        // requests. The other two elements within EmailAddressType are optional
        // (name and routing type).
        error_missing_information_email_address,

        // When creating a response object such as ForwardItem, you must supply
        // a reference item id.
        error_missing_information_reference_item_id,

        // When creating an item attachment, you must include a child element
        // indicating the item that you want to create. This response code is
        // returned if you omit this element.
        error_missing_item_for_create_item_attachment,

        // The policy ids property (Property Tag: 0x6732, String) for the folder
        // in question is missing. You should consider this a corrupt folder.
        error_missing_managed_folder_id,

        // Indicates you tried to send an item with no recipients. Note that if
        // you call CreateItem with a message disposition that causes the
        // message to be sent, you get a different response code
        // (ErrorInvalidRecipients).
        error_missing_recipients,

        // Indicates that the move or copy operation failed. A move occurs in
        // CreateItem when you accept a meeting request that is in the Deleted
        // Items folder. In addition, if you decline a meeting request, cancel a
        // calendar item, or remove a meeting from your calendar, it gets moved
        // to the deleted items folder.
        error_move_copy_failed,

        // You cannot move a distinguished folder.
        error_move_distinguished_folder,

        // This is not actually an error. It should be categorized as a warning.
        // This response code indicates that the ambiguous name that you
        // specified matched more than one contact or distribution list.. This
        // is also the only "error" response code that includes response
        // data (the matched names).
        error_name_resolution_multiple_results,

        // Indicates that the effective caller does not have a Mailbox on the
        // system. Name resolution considers both the Active Directory as well
        // as the Store Mailbox.
        error_name_resolution_no_mailbox,

        // The ambiguous name did not match any contacts in either the Active
        // Directory or the Mailbox.
        error_name_resolution_no_results,

        // There was no calendar folder for the Mailbox in question.
        error_no_calendar,

        // You can set the FolderClass only when creating a generic folder. For
        // typed folders (i.e. CalendarFolder and TaskFolder, the folder class
        // is implied. Note that if you try to set the folder class to a
        // different folder type via UpdateFolder, you get
        // ErrorObjectTypeChanged—so don't even try it (we knew you were
        // going there...). Exchange Web Services should enable you to create a
        // more specialized— but consistent—folder class when creating a
        // strongly typed folder. To get around this issue, use a generic folder
        // type but set the folder class to the value you need. Exchange Web
        // Services then creates the correct strongly typed folder.
        error_no_folder_class_override,

        // Indicates that the caller does not have free/busy viewing rights on
        // the calendar folder in question.
        error_no_free_busy_access,

        // This response code is returned in two cases:
        error_non_existent_mailbox,

        // For requests that take an SMTP address, the address must be the
        // primary SMTP address representing the Mailbox. Non-primary SMTP
        // addresses are not permitted. The response includes the correct SMTP
        // address to use. Because Exchange Web Services returns the primary
        // SMTP address, it makes you wonder why Exchange Web Services didn't
        // just use the proxy address you passed in… Note that this
        // requirement may be removed in a future release.
        error_non_primary_smtp_address,

        // Messaging Application Programming Interface (MAPI) properties in the
        // custom range (0x8000 and greater) cannot be referenced by property
        // tags. You must use PropertySetId or DistinguishedPropertySetId along
        // with PropertyName or PropertyId.
        error_no_property_tag_for_custom_properties,

        // The operation could not complete due to insufficient memory.
        error_not_enough_memory,

        // For CreateItem, you cannot set the ItemClass so that it is
        // inconsistent with the strongly typed item (i.e. Message or Contact).
        // It must be consistent. For UpdateItem/Folder, you cannot change the
        // item or folder class such that the type of the item/folder will
        // change. You can change the item/folder class to a more derived
        // instance of the same type (for example, IPM.Note to IPM.Note.Foo).
        // Note that with CreateFolder, if you try to override the folder class
        // so that it is different than the strongly typed folder element, you
        // get an ErrorNoFolderClassOverride. Treat ErrorObjectTypeChanged and
        // ErrorNoFolderClassOverride in the same manner.
        error_object_type_changed,

        // Indicates that the time allotment for a given occurrence overlaps
        // with one of its neighbors.
        error_occurrence_crossing_boundary,

        // Indicates that the time allotment for a given occurrence is too long,
        // which causes the occurrence to overlap with its neighbor. This
        // response code also occurs if the length in minutes of a given
        // occurrence is larger than Int32.MaxValue.
        error_occurrence_time_span_too_big,

        // You will never encounter this response code.
        error_parent_folder_id_required,

        // The parent folder id that you specified does not exist.
        error_parent_folder_not_found,

        // You must change your password before you can access this Mailbox.
        // This occurs when a new account has been created, and the
        // administrator indicated that the user must change the password at
        // first logon. You cannot change a password through Exchange Web
        // Services. You must use a user application such as Outlook Web Access
        // (OWA) to change your password.
        error_password_change_required,

        // The password associated with the calling account has expired.. You
        // need to change your password. You cannot change a password through
        // Exchange Web Services. You must use a user application such as
        // Outlook Web Access to change your password.
        error_password_expired,

        // Update failed due to invalid property values. The response message
        // includes the offending property paths.
        error_property_update,

        // You will never encounter this response code.
        error_property_validation_failure,

        // You will likely never encounter this response code. This response
        // code indicates that the request that Exchange Web Services sent to
        // another CAS when trying to fulfill a GetUserAvailability request was
        // invalid. This response code likely indicates a configuration or
        // rights error, or someone trying unsuccessfully to mimic an
        // availability proxy request.
        error_proxy_request_not_allowed,

        // The recipient passed to GetUserAvailability is located on a legacy
        // Exchange server (prior to Exchange Server 2007). As such, Exchange
        // Web Services needed to contact the public folder server to retrieve
        // free/busy information for that recipient. Unfortunately, this call
        // failed, resulting in Exchange Web Services returning a response code
        // of ErrorPublicFolderRequestProcessingFailed.
        error_public_folder_request_processing_failed,

        // The recipient in question is located on a legacy Exchange server
        // (prior to Exchange -2007). As such, Exchange Web Services needed to
        // contact the public folder server to retrieve free/busy information
        // for that recipient. However, the organizational unit in question did
        // not have a public folder server associated with it.
        error_public_folder_server_not_found,

        // Restrictions can contain a maximum of 255 filter expressions. If you
        // try to bind to an existing search folder that exceeds this limit, you
        // encounter this response code.
        error_query_filter_too_long,

        // The Mailbox quota has been exceeded.
        error_quota_exceeded,

        // The process for reading events was aborted due to an internal
        // failure. You should recreate the subscription based on a last known
        // watermark.
        error_read_events_failed,

        // You cannot suppress a read receipt if the message sender did not
        // request a read receipt on the message.
        error_read_receipt_not_pending,

        // The end date for the recurrence was out of range (it is limited to
        // September 1, 4500).
        error_recurrence_end_date_too_big,

        // The recurrence has no occurrence instances in the specified range.
        error_recurrence_has_no_occurrence,

        // You will never encounter this response code.
        error_request_aborted,

        // During GetUserAvailability processing, the request was deemed larger
        // than it should be. You should not encounter this response code.
        error_request_stream_too_big,

        // Indicates that one or more of the required properties is missing
        // during a CreateAttachment call. The response XML indicates which
        // property path was not set.
        error_required_property_missing,

        // You will never encounter this response code. Just as a piece of
        // trivia, the Exchange Web Services design team used this response code
        // for debug builds to ensure that their responses were schema
        // compliant. If Exchange Web Services expects you to send schema-
        // compliant XML, then the least Exchange Web Services can do is be
        // compliant in return.
        error_response_schema_validation,

        // A restriction can have a maximum of 255 filter elements.
        error_restriction_too_long,

        // Exchange Web Services cannot evaluate the restriction you supplied.
        // The restriction might appear simple, but Exchange Web Services does
        // not agree with you.
        error_restriction_too_complex,

        // The number of calendar entries for a given recipient exceeds the
        // allowable limit (1000). Reduce the window and try again.
        error_result_set_too_big,

        // Indicates that the folder you want to save the item to does not
        // exist.
        error_saved_item_folder_not_found,

        // Exchange Web Services validates all incoming requests against the
        // schema files (types.xsd, messages.xsd). Any instance documents that
        // are not compliant are rejected, and this response code is returned.
        // Note that this response code is always returned within a SOAP fault.
        error_schema_validation,

        // Indicates that the search folder has been created, but the search
        // criteria was never set on the folder. This condition occurs only when
        // you access corrupt search folders that were created with another
        // application programming interface (API) or client. Exchange Web
        // Services does not enable you to create search folders with this
        // condition To fix a search folder that has not been initialized, call
        // UpdateFolder and set the SearchParameters to include the restriction
        // that should be on the folder.
        error_search_folder_not_initialized,

        // The caller does not have Send As rights for the account in question.
        error_send_as_denied,

        // When you are the organizer and are deleting a calendar item, you must
        // set the SendMeetingCancellations attribute on the DeleteItem request
        // to indicate whether meeting cancellations will be sent to the meeting
        // attendees. If you are using the proxy classes don't forget to set
        // the SendMeetingCancellationsSpecified property to true.
        error_send_meeting_cancellations_required,

        // When you are the organizer and are updating a calendar item, you must
        // set the SendMeetingInvitationsOrCancellations attribute on the
        // UpdateItem request. If you are using the proxy classes don't forget
        // to set the SendMeetingInvitationsOrCancellationsSpecified attribute
        // to true.
        error_send_meeting_invitations_or_cancellations_required,

        // When creating a calendar item, you must set the
        // SendMeetingInvitiations attribute on the CreateItem request. If you
        // are using the proxy classes don't forget to set the
        // SendMeetingInvitationsSpecified attribute to true.
        error_send_meeting_invitations_required,

        // After the organizer sends a meeting request, that request cannot be
        // updated. If the organizer wants to modify the meeting, you need to
        // modify the calendar item, not the meeting request.
        error_sent_meeting_request_update,

        // After the task initiator sends a task request, that request cannot be
        // updated. However, you should not encounter this response code because
        // Exchange Web Services does not support task assignment at this point.
        error_sent_task_request_update,

        // The server is busy, potentially due to virus scan operations. It is
        // unlikely that you will encounter this response code.
        error_server_busy,

        // You must supply an up-to-date change key when calling the applicable
        // methods. You either did not supply a change key, or the change key
        // you supplied is stale. Call GetItem to retrieve an updated change key
        // and then try your operation again.
        error_stale_object,

        // You tried to access a subscription by using an account that did not
        // create that subscription. Each subscription is tied to its creator.
        // It does not matter which rights one account has on the Mailbox in
        // question. Jane's subscriptions can only be accessed by Jane.
        error_subscription_access_denied,

        // You can cannot create a subscription if you are not the owner or do
        // not have owner access to the Mailbox in question.
        error_subscription_delegate_access_not_supported,

        // The specified subscription does not exist which could mean that the
        // subscription expired, the Exchange Web Services process was
        // restarted, or you passed in an invalid subscription. If you encounter
        // this response code, recreate the subscription by using the last
        // watermark that you have.
        error_subscription_not_found,

        // Indicates that the folder id you specified in your SyncFolderItems
        // request does not exist.
        error_sync_folder_not_found,

        // The time window specified is larger than the allowable limit (42 by
        // default).
        error_time_interval_too_big,

        // The specified destination folder does not exist
        error_to_folder_not_found,

        // The calling account does not have the ms-Exch-EPI-TokenSerialization
        // right on the CAS that is being called. Of course, because you are not
        // using token serialization in your application, you should never
        // encounter this response code. Right?
        error_token_serialization_denied,

        // You will never encounter this response code.
        error_unable_to_get_user_oof_settings,

        // You tried to set the Culture property to a value that is not parsable
        // by the System.Globalization.CultureInfo class.
        error_unsupported_culture,

        // MAPI property types Error, Null, Object and ObjectArray are
        // unsupported.
        error_unsupported_mapi_property_type,

        // You can retrieve or set MIME content only for a post, message, or
        // calendar item.
        error_unsupported_mime_conversion,

        // Indicates that the property path cannot be used within a restriction.
        error_unsupported_path_for_query,

        // Indicates that the property path cannot be use for sorting or
        // grouping operations.
        error_unsupported_path_for_sort_group,

        // You should never encounter this response code.
        error_unsupported_property_definition,

        // Exchange Web Services cannot render the existing search folder
        // restriction. This response code does not mean that anything is wrong
        // with the search folder restriction. You can still call FindItem on
        // the search folder to retrieve the items in the search folder; you
        // just can't get the actual restriction clause.
        error_unsupported_query_filter,

        // You supplied a recurrence pattern that is not supported for tasks.
        error_unsupported_recurrence,

        // You should never encounter this response code.
        error_unsupported_sub_filter,

        // You should never encounter this response code. It indicates that
        // Exchange Web Services found a property type in the Store that it
        // cannot generate XML for.
        error_unsupported_type_for_conversion,

        // The single property path listed in a change description must match
        // the single property that is being set within the actual Item/Folder
        // element.
        error_update_property_mismatch,

        // The Exchange Store detected a virus in the message you are trying to
        // deal with.
        error_virus_detected,

        // The Exchange Store detected a virus in the message and deleted it.
        error_virus_message_deleted,

        // You will never encounter this response code. This was left over from
        // the development cycle before the Exchange Web Services team had
        // implemented voice mail folder support. Yes, there was a time when all
        // of this was not implemented.
        error_voice_mail_not_implemented,

        // You will never encounter this response code. It originally meant that
        // you intended to send your Web request from Arizona, but it actually
        // came from Minnesota instead.*
        error_web_request_in_invalid_state,

        // Indicates that there was a failure when Exchange Web Services was
        // talking with unmanaged code. Of course, you cannot see the inner
        // exception because this is a SOAP response.
        error_win32_interop_error,

        // You will never encounter this response code.
        error_working_hours_save_failed,

        // You will never encounter this response code.
        error_working_hours_xml_malformed
    };

    // TODO: move to internal namespace
    inline response_code str_to_response_code(const char* str)
    {
        static const std::unordered_map<std::string, response_code> m{
            { "NoError", response_code::no_error },
            { "ErrorAccessDenied", response_code::error_access_denied },
            { "ErrorAccountDisabled", response_code::error_account_disabled },
            { "ErrorAddressSpaceNotFound", response_code::error_address_space_not_found },
            { "ErrorADOperation", response_code::error_ad_operation },
            { "ErrorADSessionFilter", response_code::error_ad_session_filter },
            { "ErrorADUnavailable", response_code::error_ad_unavailable },
            { "ErrorAutoDiscoverFailed", response_code::error_auto_discover_failed },
            { "ErrorAffectedTaskOccurrencesRequired", response_code::error_affected_task_occurrences_required },
            { "ErrorAttachmentSizeLimitExceeded", response_code::error_attachment_size_limit_exceeded },
            { "ErrorAvailabilityConfigNotFound", response_code::error_availability_config_not_found },
            { "ErrorBatchProcessingStopped", response_code::error_batch_processing_stopped },
            { "ErrorCalendarCannotMoveOrCopyOccurrence", response_code::error_calendar_cannot_move_or_copy_occurrence },
            { "ErrorCalendarCannotUpdateDeletedItem", response_code::error_calendar_cannot_update_deleted_item },
            { "ErrorCalendarCannotUseIdForOccurrenceId", response_code::error_calendar_cannot_use_id_for_occurrence_id },
            { "ErrorCalendarCannotUseIdForRecurringMasterId", response_code::error_calendar_cannot_use_id_for_recurring_master_id },
            { "ErrorCalendarDurationIsTooLong", response_code::error_calendar_duration_is_too_long },
            { "ErrorCalendarEndDateIsEarlierThanStartDate", response_code::error_calendar_end_date_is_earlier_than_start_date },
            { "ErrorCalendarFolderIsInvalidForCalendarView", response_code::error_calendar_folder_is_invalid_for_calendar_view },
            { "ErrorCalendarInvalidAttributeValue", response_code::error_calendar_invalid_attribute_value },
            { "ErrorCalendarInvalidDayForTimeChangePattern", response_code::error_calendar_invalid_day_for_time_change_pattern },
            { "ErrorCalendarInvalidDayForWeeklyRecurrence", response_code::error_calendar_invalid_day_for_weekly_recurrence },
            { "ErrorCalendarInvalidPropertyState", response_code::error_calendar_invalid_property_state },
            { "ErrorCalendarInvalidPropertyValue", response_code::error_calendar_invalid_property_value },
            { "ErrorCalendarInvalidRecurrence", response_code::error_calendar_invalid_recurrence },
            { "ErrorCalendarInvalidTimeZone", response_code::error_calendar_invalid_time_zone },
            { "ErrorCalendarIsDelegatedForAccept", response_code::error_calendar_is_delegated_for_accept },
            { "ErrorCalendarIsDelegatedForDecline", response_code::error_calendar_is_delegated_for_decline },
            { "ErrorCalendarIsDelegatedForRemove", response_code::error_calendar_is_delegated_for_remove },
            { "ErrorCalendarIsDelegatedForTentative", response_code::error_calendar_is_delegated_for_tentative },
            { "ErrorCalendarIsNotOrganizer", response_code::error_calendar_is_not_organizer },
            { "ErrorCalendarIsOrganizerForAccept", response_code::error_calendar_is_organizer_for_accept },
            { "ErrorCalendarIsOrganizerForDecline", response_code::error_calendar_is_organizer_for_decline },
            { "ErrorCalendarIsOrganizerForRemove", response_code::error_calendar_is_organizer_for_remove },
            { "ErrorCalendarIsOrganizerForTentative", response_code::error_calendar_is_organizer_for_tentative },
            { "ErrorCalendarOccurrenceIndexIsOutOfRecurrenceRange", response_code::error_calendar_occurrence_index_is_out_of_recurrence_range },
            { "ErrorCalendarOccurrenceIsDeletedFromRecurrence", response_code::error_calendar_occurrence_is_deleted_from_recurrence },
            { "ErrorCalendarOutOfRange", response_code::error_calendar_out_of_range },
            { "ErrorCalendarViewRangeTooBig", response_code::error_calendar_view_range_too_big },
            { "ErrorCannotCreateCalendarItemInNonCalendarFolder", response_code::error_cannot_create_calendar_item_in_non_calendar_folder },
            { "ErrorCannotCreateContactInNonContactsFolder", response_code::error_cannot_create_contact_in_non_contacts_folder },
            { "ErrorCannotCreateTaskInNonTaskFolder", response_code::error_cannot_create_task_in_non_task_folder },
            { "ErrorCannotDeleteObject", response_code::error_cannot_delete_object },
            { "ErrorCannotDeleteTaskOccurrence", response_code::error_cannot_delete_task_occurrence },
            { "ErrorCannotOpenFileAttachment", response_code::error_cannot_open_file_attachment },
            { "ErrorCannotUseFolderIdForItemId", response_code::error_cannot_use_folder_id_for_item_id },
            { "ErrorCannotUserItemIdForFolderId", response_code::error_cannot_user_item_id_for_folder_id },
            { "ErrorChangeKeyRequired", response_code::error_change_key_required },
            { "ErrorChangeKeyRequiredForWriteOperations", response_code::error_change_key_required_for_write_operations },
            { "ErrorConnectionFailed", response_code::error_connection_failed },
            { "ErrorContentConversionFailed", response_code::error_content_conversion_failed },
            { "ErrorCorruptData", response_code::error_corrupt_data },
            { "ErrorCreateItemAccessDenied", response_code::error_create_item_access_denied },
            { "ErrorCreateManagedFolderPartialCompletion", response_code::error_create_managed_folder_partial_completion },
            { "ErrorCreateSubfolderAccessDenied", response_code::error_create_subfolder_access_denied },
            { "ErrorCrossMailboxMoveCopy", response_code::error_cross_mailbox_move_copy },
            { "ErrorDataSizeLimitExceeded", response_code::error_data_size_limit_exceeded },
            { "ErrorDataSourceOperation", response_code::error_data_source_operation },
            { "ErrorDeleteDistinguishedFolder", response_code::error_delete_distinguished_folder },
            { "ErrorDeleteItemsFailed", response_code::error_delete_items_failed },
            { "ErrorDuplicateInputFolderNames", response_code::error_duplicate_input_folder_names },
            { "ErrorEmailAddressMismatch", response_code::error_email_address_mismatch },
            { "ErrorEventNotFound", response_code::error_event_not_found },
            { "ErrorExpiredSubscription", response_code::error_expired_subscription },
            { "ErrorFolderCorrupt", response_code::error_folder_corrupt },
            { "ErrorFolderNotFound", response_code::error_folder_not_found },
            { "ErrorFolderPropertyRequestFailed", response_code::error_folder_property_request_failed },
            { "ErrorFolderSave", response_code::error_folder_save },
            { "ErrorFolderSaveFailed", response_code::error_folder_save_failed },
            { "ErrorFolderSavePropertyError", response_code::error_folder_save_property_error },
            { "ErrorFolderExists", response_code::error_folder_exists },
            { "ErrorFreeBusyGenerationFailed", response_code::error_free_busy_generation_failed },
            { "ErrorGetServerSecurityDescriptorFailed", response_code::error_get_server_security_descriptor_failed },
            { "ErrorImpersonateUserDenied", response_code::error_impersonate_user_denied },
            { "ErrorImpersonationDenied", response_code::error_impersonation_denied },
            { "ErrorImpersonationFailed", response_code::error_impersonation_failed },
            { "ErrorIncorrectUpdatePropertyCount", response_code::error_incorrect_update_property_count },
            { "ErrorIndividualMailboxLimitReached", response_code::error_individual_mailbox_limit_reached },
            { "ErrorInsufficientResources", response_code::error_insufficient_resources },
            { "ErrorInternalServerError", response_code::error_internal_server_error },
            { "ErrorInternalServerTransientError", response_code::error_internal_server_transient_error },
            { "ErrorInvalidAccessLevel", response_code::error_invalid_access_level },
            { "ErrorInvalidAttachmentId", response_code::error_invalid_attachment_id },
            { "ErrorInvalidAttachmentSubfilter", response_code::error_invalid_attachment_subfilter },
            { "ErrorInvalidAttachmentSubfilterTextFilter", response_code::error_invalid_attachment_subfilter_text_filter },
            { "ErrorInvalidAuthorizationContext", response_code::error_invalid_authorization_context },
            { "ErrorInvalidChangeKey", response_code::error_invalid_change_key },
            { "ErrorInvalidClientSecurityContext", response_code::error_invalid_client_security_context },
            { "ErrorInvalidCompleteDate", response_code::error_invalid_complete_date },
            { "ErrorInvalidCrossForestCredentials", response_code::error_invalid_cross_forest_credentials },
            { "ErrorInvalidExchangeImpersonationHeaderData", response_code::error_invalid_exchange_impersonation_header_data },
            { "ErrorInvalidExcludesRestriction", response_code::error_invalid_excludes_restriction },
            { "ErrorInvalidExpressionTypeForSubFilter", response_code::error_invalid_expression_type_for_sub_filter },
            { "ErrorInvalidExtendedProperty", response_code::error_invalid_extended_property },
            { "ErrorInvalidExtendedPropertyValue", response_code::error_invalid_extended_property_value },
            { "ErrorInvalidFolderId", response_code::error_invalid_folder_id },
            { "ErrorInvalidFractionalPagingParameters", response_code::error_invalid_fractional_paging_parameters },
            { "ErrorInvalidFreeBusyViewType", response_code::error_invalid_free_busy_view_type },
            { "ErrorInvalidId", response_code::error_invalid_id },
            { "ErrorInvalidIdEmpty", response_code::error_invalid_id_empty },
            { "ErrorInvalidIdMalformed", response_code::error_invalid_id_malformed },
            { "ErrorInvalidIdMonikerTooLong", response_code::error_invalid_id_moniker_too_long },
            { "ErrorInvalidIdNotAnItemAttachmentId", response_code::error_invalid_id_not_an_item_attachment_id },
            { "ErrorInvalidIdReturnedByResolveNames", response_code::error_invalid_id_returned_by_resolve_names },
            { "ErrorInvalidIdStoreObjectIdTooLong", response_code::error_invalid_id_store_object_id_too_long },
            { "ErrorInvalidIdTooManyAttachmentLevels", response_code::error_invalid_id_too_many_attachment_levels },
            { "ErrorInvalidIdXml", response_code::error_invalid_id_xml },
            { "ErrorInvalidIndexedPagingParameters", response_code::error_invalid_indexed_paging_parameters },
            { "ErrorInvalidInternetHeaderChildNodes", response_code::error_invalid_internet_header_child_nodes },
            { "ErrorInvalidItemForOperationCreateItemAttachment", response_code::error_invalid_item_for_operation_create_item_attachment },
            { "ErrorInvalidItemForOperationCreateItem", response_code::error_invalid_item_for_operation_create_item },
            { "ErrorInvalidItemForOperationAcceptItem", response_code::error_invalid_item_for_operation_accept_item },
            { "ErrorInvalidItemForOperationCancelItem", response_code::error_invalid_item_for_operation_cancel_item },
            { "ErrorInvalidItemForOperationDeclineItem", response_code::error_invalid_item_for_operation_decline_item },
            { "ErrorInvalidItemForOperationExpandDL", response_code::error_invalid_item_for_operation_expand_dl },
            { "ErrorInvalidItemForOperationRemoveItem", response_code::error_invalid_item_for_operation_remove_item },
            { "ErrorInvalidItemForOperationSendItem", response_code::error_invalid_item_for_operation_send_item },
            { "ErrorInvalidItemForOperationTentative", response_code::error_invalid_item_for_operation_tentative },
            { "ErrorInvalidManagedFolderProperty", response_code::error_invalid_managed_folder_property },
            { "ErrorInvalidManagedFolderQuota", response_code::error_invalid_managed_folder_quota },
            { "ErrorInvalidManagedFolderSize", response_code::error_invalid_managed_folder_size },
            { "ErrorInvalidMergedFreeBusyInterval", response_code::error_invalid_merged_free_busy_interval },
            { "ErrorInvalidNameForNameResolution", response_code::error_invalid_name_for_name_resolution },
            { "ErrorInvalidNetworkServiceContext", response_code::error_invalid_network_service_context },
            { "ErrorInvalidOofParameter", response_code::error_invalid_oof_parameter },
            { "ErrorInvalidPagingMaxRows", response_code::error_invalid_paging_max_rows },
            { "ErrorInvalidParentFolder", response_code::error_invalid_parent_folder },
            { "ErrorInvalidPercentCompleteValue", response_code::error_invalid_percent_complete_value },
            { "ErrorInvalidPropertyAppend", response_code::error_invalid_property_append },
            { "ErrorInvalidPropertyDelete", response_code::error_invalid_property_delete },
            { "ErrorInvalidPropertyForExists", response_code::error_invalid_property_for_exists },
            { "ErrorInvalidPropertyForOperation", response_code::error_invalid_property_for_operation },
            { "ErrorInvalidPropertyRequest", response_code::error_invalid_property_request },
            { "ErrorInvalidPropertySet", response_code::error_invalid_property_set },
            { "ErrorInvalidPropertyUpdateSentMessage", response_code::error_invalid_property_update_sent_message },
            { "ErrorInvalidPullSubscriptionId", response_code::error_invalid_pull_subscription_id },
            { "ErrorInvalidPushSubscriptionUrl", response_code::error_invalid_push_subscription_url },
            { "ErrorInvalidRecipients", response_code::error_invalid_recipients },
            { "ErrorInvalidRecipientSubfilter", response_code::error_invalid_recipient_subfilter },
            { "ErrorInvalidRecipientSubfilterComparison", response_code::error_invalid_recipient_subfilter_comparison },
            { "ErrorInvalidRecipientSubfilterOrder", response_code::error_invalid_recipient_subfilter_order },
            { "ErrorInvalidRecipientSubfilterTextFilter", response_code::error_invalid_recipient_subfilter_text_filter },
            { "ErrorInvalidReferenceItem", response_code::error_invalid_reference_item },
            { "ErrorInvalidRequest", response_code::error_invalid_request },
            { "ErrorInvalidRestriction", response_code::error_invalid_restriction },
            { "ErrorInvalidRoutingType", response_code::error_invalid_routing_type },
            { "ErrorInvalidScheduledOofDuration", response_code::error_invalid_scheduled_oof_duration },
            { "ErrorInvalidSecurityDescriptor", response_code::error_invalid_security_descriptor },
            { "ErrorInvalidSendItemSaveSettings", response_code::error_invalid_send_item_save_settings },
            { "ErrorInvalidSerializedAccessToken", response_code::error_invalid_serialized_access_token },
            { "ErrorInvalidSid", response_code::error_invalid_sid },
            { "ErrorInvalidSmtpAddress", response_code::error_invalid_smtp_address },
            { "ErrorInvalidSubfilterType", response_code::error_invalid_subfilter_type },
            { "ErrorInvalidSubfilterTypeNotAttendeeType", response_code::error_invalid_subfilter_type_not_attendee_type },
            { "ErrorInvalidSubfilterTypeNotRecipientType", response_code::error_invalid_subfilter_type_not_recipient_type },
            { "ErrorInvalidSubscription", response_code::error_invalid_subscription },
            { "ErrorInvalidSyncStateData", response_code::error_invalid_sync_state_data },
            { "ErrorInvalidTimeInterval", response_code::error_invalid_time_interval },
            { "ErrorInvalidUserOofSettings", response_code::error_invalid_user_oof_settings },
            { "ErrorInvalidUserPrincipalName", response_code::error_invalid_user_principal_name },
            { "ErrorInvalidUserSid", response_code::error_invalid_user_sid },
            { "ErrorInvalidUserSidMissingUPN", response_code::error_invalid_user_sid_missing_upn },
            { "ErrorInvalidValueForProperty", response_code::error_invalid_value_for_property },
            { "ErrorInvalidWatermark", response_code::error_invalid_watermark },
            { "ErrorIrresolvableConflict", response_code::error_irresolvable_conflict },
            { "ErrorItemCorrupt", response_code::error_item_corrupt },
            { "ErrorItemNotFound", response_code::error_item_not_found },
            { "ErrorItemPropertyRequestFailed", response_code::error_item_property_request_failed },
            { "ErrorItemSave", response_code::error_item_save },
            { "ErrorItemSavePropertyError", response_code::error_item_save_property_error },
            { "ErrorLegacyMailboxFreeBusyViewTypeNotMerged", response_code::error_legacy_mailbox_free_busy_view_type_not_merged },
            { "ErrorLocalServerObjectNotFound", response_code::error_local_server_object_not_found },
            { "ErrorLogonAsNetworkServiceFailed", response_code::error_logon_as_network_service_failed },
            { "ErrorMailboxConfiguration", response_code::error_mailbox_configuration },
            { "ErrorMailboxDataArrayEmpty", response_code::error_mailbox_data_array_empty },
            { "ErrorMailboxDataArrayTooBig", response_code::error_mailbox_data_array_too_big },
            { "ErrorMailboxLogonFailed", response_code::error_mailbox_logon_failed },
            { "ErrorMailboxMoveInProgress", response_code::error_mailbox_move_in_progress },
            { "ErrorMailboxStoreUnavailable", response_code::error_mailbox_store_unavailable },
            { "ErrorMailRecipientNotFound", response_code::error_mail_recipient_not_found },
            { "ErrorManagedFolderAlreadyExists", response_code::error_managed_folder_already_exists },
            { "ErrorManagedFolderNotFound", response_code::error_managed_folder_not_found },
            { "ErrorManagedFoldersRootFailure", response_code::error_managed_folders_root_failure },
            { "ErrorMeetingSuggestionGenerationFailed", response_code::error_meeting_suggestion_generation_failed },
            { "ErrorMessageDispositionRequired", response_code::error_message_disposition_required },
            { "ErrorMessageSizeExceeded", response_code::error_message_size_exceeded },
            { "ErrorMimeContentConversionFailed", response_code::error_mime_content_conversion_failed },
            { "ErrorMimeContentInvalid", response_code::error_mime_content_invalid },
            { "ErrorMimeContentInvalidBase64String", response_code::error_mime_content_invalid_base64_string },
            { "ErrorMissingArgument", response_code::error_missing_argument },
            { "ErrorMissingEmailAddress", response_code::error_missing_email_address },
            { "ErrorMissingEmailAddressForManagedFolder", response_code::error_missing_email_address_for_managed_folder },
            { "ErrorMissingInformationEmailAddress", response_code::error_missing_information_email_address },
            { "ErrorMissingInformationReferenceItemId", response_code::error_missing_information_reference_item_id },
            { "ErrorMissingItemForCreateItemAttachment", response_code::error_missing_item_for_create_item_attachment },
            { "ErrorMissingManagedFolderId", response_code::error_missing_managed_folder_id },
            { "ErrorMissingRecipients", response_code::error_missing_recipients },
            { "ErrorMoveCopyFailed", response_code::error_move_copy_failed },
            { "ErrorMoveDistinguishedFolder", response_code::error_move_distinguished_folder },
            { "ErrorNameResolutionMultipleResults", response_code::error_name_resolution_multiple_results },
            { "ErrorNameResolutionNoMailbox", response_code::error_name_resolution_no_mailbox },
            { "ErrorNameResolutionNoResults", response_code::error_name_resolution_no_results },
            { "ErrorNoCalendar", response_code::error_no_calendar },
            { "ErrorNoFolderClassOverride", response_code::error_no_folder_class_override },
            { "ErrorNoFreeBusyAccess", response_code::error_no_free_busy_access },
            { "ErrorNonExistentMailbox", response_code::error_non_existent_mailbox },
            { "ErrorNonPrimarySmtpAddress", response_code::error_non_primary_smtp_address },
            { "ErrorNoPropertyTagForCustomProperties", response_code::error_no_property_tag_for_custom_properties },
            { "ErrorNotEnoughMemory", response_code::error_not_enough_memory },
            { "ErrorObjectTypeChanged", response_code::error_object_type_changed },
            { "ErrorOccurrenceCrossingBoundary", response_code::error_occurrence_crossing_boundary },
            { "ErrorOccurrenceTimeSpanTooBig", response_code::error_occurrence_time_span_too_big },
            { "ErrorParentFolderIdRequired", response_code::error_parent_folder_id_required },
            { "ErrorParentFolderNotFound", response_code::error_parent_folder_not_found },
            { "ErrorPasswordChangeRequired", response_code::error_password_change_required },
            { "ErrorPasswordExpired", response_code::error_password_expired },
            { "ErrorPropertyUpdate", response_code::error_property_update },
            { "ErrorPropertyValidationFailure", response_code::error_property_validation_failure },
            { "ErrorProxyRequestNotAllowed", response_code::error_proxy_request_not_allowed },
            { "ErrorPublicFolderRequestProcessingFailed", response_code::error_public_folder_request_processing_failed },
            { "ErrorPublicFolderServerNotFound", response_code::error_public_folder_server_not_found },
            { "ErrorQueryFilterTooLong", response_code::error_query_filter_too_long },
            { "ErrorQuotaExceeded", response_code::error_quota_exceeded },
            { "ErrorReadEventsFailed", response_code::error_read_events_failed },
            { "ErrorReadReceiptNotPending", response_code::error_read_receipt_not_pending },
            { "ErrorRecurrenceEndDateTooBig", response_code::error_recurrence_end_date_too_big },
            { "ErrorRecurrenceHasNoOccurrence", response_code::error_recurrence_has_no_occurrence },
            { "ErrorRequestAborted", response_code::error_request_aborted },
            { "ErrorRequestStreamTooBig", response_code::error_request_stream_too_big },
            { "ErrorRequiredPropertyMissing", response_code::error_required_property_missing },
            { "ErrorResponseSchemaValidation", response_code::error_response_schema_validation },
            { "ErrorRestrictionTooLong", response_code::error_restriction_too_long },
            { "ErrorRestrictionTooComplex", response_code::error_restriction_too_complex },
            { "ErrorResultSetTooBig", response_code::error_result_set_too_big },
            { "ErrorSavedItemFolderNotFound", response_code::error_saved_item_folder_not_found },
            { "ErrorSchemaValidation", response_code::error_schema_validation },
            { "ErrorSearchFolderNotInitialized", response_code::error_search_folder_not_initialized },
            { "ErrorSendAsDenied", response_code::error_send_as_denied },
            { "ErrorSendMeetingCancellationsRequired", response_code::error_send_meeting_cancellations_required },
            { "ErrorSendMeetingInvitationsOrCancellationsRequired", response_code::error_send_meeting_invitations_or_cancellations_required },
            { "ErrorSendMeetingInvitationsRequired", response_code::error_send_meeting_invitations_required },
            { "ErrorSentMeetingRequestUpdate", response_code::error_sent_meeting_request_update },
            { "ErrorSentTaskRequestUpdate", response_code::error_sent_task_request_update },
            { "ErrorServerBusy", response_code::error_server_busy },
            { "ErrorStaleObject", response_code::error_stale_object },
            { "ErrorSubscriptionAccessDenied", response_code::error_subscription_access_denied },
            { "ErrorSubscriptionDelegateAccessNotSupported", response_code::error_subscription_delegate_access_not_supported },
            { "ErrorSubscriptionNotFound", response_code::error_subscription_not_found },
            { "ErrorSyncFolderNotFound", response_code::error_sync_folder_not_found },
            { "ErrorTimeIntervalTooBig", response_code::error_time_interval_too_big },
            { "ErrorToFolderNotFound", response_code::error_to_folder_not_found },
            { "ErrorTokenSerializationDenied", response_code::error_token_serialization_denied },
            { "ErrorUnableToGetUserOofSettings", response_code::error_unable_to_get_user_oof_settings },
            { "ErrorUnsupportedCulture", response_code::error_unsupported_culture },
            { "ErrorUnsupportedMapiPropertyType", response_code::error_unsupported_mapi_property_type },
            { "ErrorUnsupportedMimeConversion", response_code::error_unsupported_mime_conversion },
            { "ErrorUnsupportedPathForQuery", response_code::error_unsupported_path_for_query },
            { "ErrorUnsupportedPathForSortGroup", response_code::error_unsupported_path_for_sort_group },
            { "ErrorUnsupportedPropertyDefinition", response_code::error_unsupported_property_definition },
            { "ErrorUnsupportedQueryFilter", response_code::error_unsupported_query_filter },
            { "ErrorUnsupportedRecurrence", response_code::error_unsupported_recurrence },
            { "ErrorUnsupportedSubFilter", response_code::error_unsupported_sub_filter },
            { "ErrorUnsupportedTypeForConversion", response_code::error_unsupported_type_for_conversion },
            { "ErrorUpdatePropertyMismatch", response_code::error_update_property_mismatch },
            { "ErrorVirusDetected", response_code::error_virus_detected },
            { "ErrorVirusMessageDeleted", response_code::error_virus_message_deleted },
            { "ErrorVoiceMailNotImplemented", response_code::error_voice_mail_not_implemented },
            { "ErrorWebRequestInInvalidState", response_code::error_web_request_in_invalid_state },
            { "ErrorWin32InteropError", response_code::error_win32_interop_error },
            { "ErrorWorkingHoursSaveFailed", response_code::error_working_hours_save_failed },
            { "ErrorWorkingHoursXmlMalformed", response_code::error_working_hours_xml_malformed }
        };
        auto it = m.find(str);
        if (it == m.end())
        {
            throw exception(std::string("Unrecognized response code: ") + str);
        }
        return it->second;
    }

    // TODO: move to internal namespace
    inline const std::string& response_code_to_str(response_code code)
    {
        static const std::unordered_map<response_code, std::string,
                     internal::enum_class_hash> m{
            { response_code::no_error, "NoError" },
            { response_code::error_access_denied, "ErrorAccessDenied" },
            { response_code::error_account_disabled, "ErrorAccountDisabled" },
            { response_code::error_address_space_not_found, "ErrorAddressSpaceNotFound" },
            { response_code::error_ad_operation, "ErrorADOperation" },
            { response_code::error_ad_session_filter, "ErrorADSessionFilter" },
            { response_code::error_ad_unavailable, "ErrorADUnavailable" },
            { response_code::error_auto_discover_failed, "ErrorAutoDiscoverFailed" },
            { response_code::error_affected_task_occurrences_required, "ErrorAffectedTaskOccurrencesRequired" },
            { response_code::error_attachment_size_limit_exceeded, "ErrorAttachmentSizeLimitExceeded" },
            { response_code::error_availability_config_not_found, "ErrorAvailabilityConfigNotFound" },
            { response_code::error_batch_processing_stopped, "ErrorBatchProcessingStopped" },
            { response_code::error_calendar_cannot_move_or_copy_occurrence, "ErrorCalendarCannotMoveOrCopyOccurrence" },
            { response_code::error_calendar_cannot_update_deleted_item, "ErrorCalendarCannotUpdateDeletedItem" },
            { response_code::error_calendar_cannot_use_id_for_occurrence_id, "ErrorCalendarCannotUseIdForOccurrenceId" },
            { response_code::error_calendar_cannot_use_id_for_recurring_master_id, "ErrorCalendarCannotUseIdForRecurringMasterId" },
            { response_code::error_calendar_duration_is_too_long, "ErrorCalendarDurationIsTooLong" },
            { response_code::error_calendar_end_date_is_earlier_than_start_date, "ErrorCalendarEndDateIsEarlierThanStartDate" },
            { response_code::error_calendar_folder_is_invalid_for_calendar_view, "ErrorCalendarFolderIsInvalidForCalendarView" },
            { response_code::error_calendar_invalid_attribute_value, "ErrorCalendarInvalidAttributeValue" },
            { response_code::error_calendar_invalid_day_for_time_change_pattern, "ErrorCalendarInvalidDayForTimeChangePattern" },
            { response_code::error_calendar_invalid_day_for_weekly_recurrence, "ErrorCalendarInvalidDayForWeeklyRecurrence" },
            { response_code::error_calendar_invalid_property_state, "ErrorCalendarInvalidPropertyState" },
            { response_code::error_calendar_invalid_property_value, "ErrorCalendarInvalidPropertyValue" },
            { response_code::error_calendar_invalid_recurrence, "ErrorCalendarInvalidRecurrence" },
            { response_code::error_calendar_invalid_time_zone, "ErrorCalendarInvalidTimeZone" },
            { response_code::error_calendar_is_delegated_for_accept, "ErrorCalendarIsDelegatedForAccept" },
            { response_code::error_calendar_is_delegated_for_decline, "ErrorCalendarIsDelegatedForDecline" },
            { response_code::error_calendar_is_delegated_for_remove, "ErrorCalendarIsDelegatedForRemove" },
            { response_code::error_calendar_is_delegated_for_tentative, "ErrorCalendarIsDelegatedForTentative" },
            { response_code::error_calendar_is_not_organizer, "ErrorCalendarIsNotOrganizer" },
            { response_code::error_calendar_is_organizer_for_accept, "ErrorCalendarIsOrganizerForAccept" },
            { response_code::error_calendar_is_organizer_for_decline, "ErrorCalendarIsOrganizerForDecline" },
            { response_code::error_calendar_is_organizer_for_remove, "ErrorCalendarIsOrganizerForRemove" },
            { response_code::error_calendar_is_organizer_for_tentative, "ErrorCalendarIsOrganizerForTentative" },
            { response_code::error_calendar_occurrence_index_is_out_of_recurrence_range, "ErrorCalendarOccurrenceIndexIsOutOfRecurrenceRange" },
            { response_code::error_calendar_occurrence_is_deleted_from_recurrence, "ErrorCalendarOccurrenceIsDeletedFromRecurrence" },
            { response_code::error_calendar_out_of_range, "ErrorCalendarOutOfRange" },
            { response_code::error_calendar_view_range_too_big, "ErrorCalendarViewRangeTooBig" },
            { response_code::error_cannot_create_calendar_item_in_non_calendar_folder, "ErrorCannotCreateCalendarItemInNonCalendarFolder" },
            { response_code::error_cannot_create_contact_in_non_contacts_folder, "ErrorCannotCreateContactInNonContactsFolder" },
            { response_code::error_cannot_create_task_in_non_task_folder, "ErrorCannotCreateTaskInNonTaskFolder" },
            { response_code::error_cannot_delete_object, "ErrorCannotDeleteObject" },
            { response_code::error_cannot_delete_task_occurrence, "ErrorCannotDeleteTaskOccurrence" },
            { response_code::error_cannot_open_file_attachment, "ErrorCannotOpenFileAttachment" },
            { response_code::error_cannot_use_folder_id_for_item_id, "ErrorCannotUseFolderIdForItemId" },
            { response_code::error_cannot_user_item_id_for_folder_id, "ErrorCannotUserItemIdForFolderId" },
            { response_code::error_change_key_required, "ErrorChangeKeyRequired" },
            { response_code::error_change_key_required_for_write_operations, "ErrorChangeKeyRequiredForWriteOperations" },
            { response_code::error_connection_failed, "ErrorConnectionFailed" },
            { response_code::error_content_conversion_failed, "ErrorContentConversionFailed" },
            { response_code::error_corrupt_data, "ErrorCorruptData" },
            { response_code::error_create_item_access_denied, "ErrorCreateItemAccessDenied" },
            { response_code::error_create_managed_folder_partial_completion, "ErrorCreateManagedFolderPartialCompletion" },
            { response_code::error_create_subfolder_access_denied, "ErrorCreateSubfolderAccessDenied" },
            { response_code::error_cross_mailbox_move_copy, "ErrorCrossMailboxMoveCopy" },
            { response_code::error_data_size_limit_exceeded, "ErrorDataSizeLimitExceeded" },
            { response_code::error_data_source_operation, "ErrorDataSourceOperation" },
            { response_code::error_delete_distinguished_folder, "ErrorDeleteDistinguishedFolder" },
            { response_code::error_delete_items_failed, "ErrorDeleteItemsFailed" },
            { response_code::error_duplicate_input_folder_names, "ErrorDuplicateInputFolderNames" },
            { response_code::error_email_address_mismatch, "ErrorEmailAddressMismatch" },
            { response_code::error_event_not_found, "ErrorEventNotFound" },
            { response_code::error_expired_subscription, "ErrorExpiredSubscription" },
            { response_code::error_folder_corrupt, "ErrorFolderCorrupt" },
            { response_code::error_folder_not_found, "ErrorFolderNotFound" },
            { response_code::error_folder_property_request_failed, "ErrorFolderPropertyRequestFailed" },
            { response_code::error_folder_save, "ErrorFolderSave" },
            { response_code::error_folder_save_failed, "ErrorFolderSaveFailed" },
            { response_code::error_folder_save_property_error, "ErrorFolderSavePropertyError" },
            { response_code::error_folder_exists, "ErrorFolderExists" },
            { response_code::error_free_busy_generation_failed, "ErrorFreeBusyGenerationFailed" },
            { response_code::error_get_server_security_descriptor_failed, "ErrorGetServerSecurityDescriptorFailed" },
            { response_code::error_impersonate_user_denied, "ErrorImpersonateUserDenied" },
            { response_code::error_impersonation_denied, "ErrorImpersonationDenied" },
            { response_code::error_impersonation_failed, "ErrorImpersonationFailed" },
            { response_code::error_incorrect_update_property_count, "ErrorIncorrectUpdatePropertyCount" },
            { response_code::error_individual_mailbox_limit_reached, "ErrorIndividualMailboxLimitReached" },
            { response_code::error_insufficient_resources, "ErrorInsufficientResources" },
            { response_code::error_internal_server_error, "ErrorInternalServerError" },
            { response_code::error_internal_server_transient_error, "ErrorInternalServerTransientError" },
            { response_code::error_invalid_access_level, "ErrorInvalidAccessLevel" },
            { response_code::error_invalid_attachment_id, "ErrorInvalidAttachmentId" },
            { response_code::error_invalid_attachment_subfilter, "ErrorInvalidAttachmentSubfilter" },
            { response_code::error_invalid_attachment_subfilter_text_filter, "ErrorInvalidAttachmentSubfilterTextFilter" },
            { response_code::error_invalid_authorization_context, "ErrorInvalidAuthorizationContext" },
            { response_code::error_invalid_change_key, "ErrorInvalidChangeKey" },
            { response_code::error_invalid_client_security_context, "ErrorInvalidClientSecurityContext" },
            { response_code::error_invalid_complete_date, "ErrorInvalidCompleteDate" },
            { response_code::error_invalid_cross_forest_credentials, "ErrorInvalidCrossForestCredentials" },
            { response_code::error_invalid_exchange_impersonation_header_data, "ErrorInvalidExchangeImpersonationHeaderData" },
            { response_code::error_invalid_excludes_restriction, "ErrorInvalidExcludesRestriction" },
            { response_code::error_invalid_expression_type_for_sub_filter, "ErrorInvalidExpressionTypeForSubFilter" },
            { response_code::error_invalid_extended_property, "ErrorInvalidExtendedProperty" },
            { response_code::error_invalid_extended_property_value, "ErrorInvalidExtendedPropertyValue" },
            { response_code::error_invalid_folder_id, "ErrorInvalidFolderId" },
            { response_code::error_invalid_fractional_paging_parameters, "ErrorInvalidFractionalPagingParameters" },
            { response_code::error_invalid_free_busy_view_type, "ErrorInvalidFreeBusyViewType" },
            { response_code::error_invalid_id, "ErrorInvalidId" },
            { response_code::error_invalid_id_empty, "ErrorInvalidIdEmpty" },
            { response_code::error_invalid_id_malformed, "ErrorInvalidIdMalformed" },
            { response_code::error_invalid_id_moniker_too_long, "ErrorInvalidIdMonikerTooLong" },
            { response_code::error_invalid_id_not_an_item_attachment_id, "ErrorInvalidIdNotAnItemAttachmentId" },
            { response_code::error_invalid_id_returned_by_resolve_names, "ErrorInvalidIdReturnedByResolveNames" },
            { response_code::error_invalid_id_store_object_id_too_long, "ErrorInvalidIdStoreObjectIdTooLong" },
            { response_code::error_invalid_id_too_many_attachment_levels, "ErrorInvalidIdTooManyAttachmentLevels" },
            { response_code::error_invalid_id_xml, "ErrorInvalidIdXml" },
            { response_code::error_invalid_indexed_paging_parameters, "ErrorInvalidIndexedPagingParameters" },
            { response_code::error_invalid_internet_header_child_nodes, "ErrorInvalidInternetHeaderChildNodes" },
            { response_code::error_invalid_item_for_operation_create_item_attachment, "ErrorInvalidItemForOperationCreateItemAttachment" },
            { response_code::error_invalid_item_for_operation_create_item, "ErrorInvalidItemForOperationCreateItem" },
            { response_code::error_invalid_item_for_operation_accept_item, "ErrorInvalidItemForOperationAcceptItem" },
            { response_code::error_invalid_item_for_operation_cancel_item, "ErrorInvalidItemForOperationCancelItem" },
            { response_code::error_invalid_item_for_operation_decline_item, "ErrorInvalidItemForOperationDeclineItem" },
            { response_code::error_invalid_item_for_operation_expand_dl, "ErrorInvalidItemForOperationExpandDL" },
            { response_code::error_invalid_item_for_operation_remove_item, "ErrorInvalidItemForOperationRemoveItem" },
            { response_code::error_invalid_item_for_operation_send_item, "ErrorInvalidItemForOperationSendItem" },
            { response_code::error_invalid_item_for_operation_tentative, "ErrorInvalidItemForOperationTentative" },
            { response_code::error_invalid_managed_folder_property, "ErrorInvalidManagedFolderProperty" },
            { response_code::error_invalid_managed_folder_quota, "ErrorInvalidManagedFolderQuota" },
            { response_code::error_invalid_managed_folder_size, "ErrorInvalidManagedFolderSize" },
            { response_code::error_invalid_merged_free_busy_interval, "ErrorInvalidMergedFreeBusyInterval" },
            { response_code::error_invalid_name_for_name_resolution, "ErrorInvalidNameForNameResolution" },
            { response_code::error_invalid_network_service_context, "ErrorInvalidNetworkServiceContext" },
            { response_code::error_invalid_oof_parameter, "ErrorInvalidOofParameter" },
            { response_code::error_invalid_paging_max_rows, "ErrorInvalidPagingMaxRows" },
            { response_code::error_invalid_parent_folder, "ErrorInvalidParentFolder" },
            { response_code::error_invalid_percent_complete_value, "ErrorInvalidPercentCompleteValue" },
            { response_code::error_invalid_property_append, "ErrorInvalidPropertyAppend" },
            { response_code::error_invalid_property_delete, "ErrorInvalidPropertyDelete" },
            { response_code::error_invalid_property_for_exists, "ErrorInvalidPropertyForExists" },
            { response_code::error_invalid_property_for_operation, "ErrorInvalidPropertyForOperation" },
            { response_code::error_invalid_property_request, "ErrorInvalidPropertyRequest" },
            { response_code::error_invalid_property_set, "ErrorInvalidPropertySet" },
            { response_code::error_invalid_property_update_sent_message, "ErrorInvalidPropertyUpdateSentMessage" },
            { response_code::error_invalid_pull_subscription_id, "ErrorInvalidPullSubscriptionId" },
            { response_code::error_invalid_push_subscription_url, "ErrorInvalidPushSubscriptionUrl" },
            { response_code::error_invalid_recipients, "ErrorInvalidRecipients" },
            { response_code::error_invalid_recipient_subfilter, "ErrorInvalidRecipientSubfilter" },
            { response_code::error_invalid_recipient_subfilter_comparison, "ErrorInvalidRecipientSubfilterComparison" },
            { response_code::error_invalid_recipient_subfilter_order, "ErrorInvalidRecipientSubfilterOrder" },
            { response_code::error_invalid_recipient_subfilter_text_filter, "ErrorInvalidRecipientSubfilterTextFilter" },
            { response_code::error_invalid_reference_item, "ErrorInvalidReferenceItem" },
            { response_code::error_invalid_request, "ErrorInvalidRequest" },
            { response_code::error_invalid_restriction, "ErrorInvalidRestriction" },
            { response_code::error_invalid_routing_type, "ErrorInvalidRoutingType" },
            { response_code::error_invalid_scheduled_oof_duration, "ErrorInvalidScheduledOofDuration" },
            { response_code::error_invalid_security_descriptor, "ErrorInvalidSecurityDescriptor" },
            { response_code::error_invalid_send_item_save_settings, "ErrorInvalidSendItemSaveSettings" },
            { response_code::error_invalid_serialized_access_token, "ErrorInvalidSerializedAccessToken" },
            { response_code::error_invalid_sid, "ErrorInvalidSid" },
            { response_code::error_invalid_smtp_address, "ErrorInvalidSmtpAddress" },
            { response_code::error_invalid_subfilter_type, "ErrorInvalidSubfilterType" },
            { response_code::error_invalid_subfilter_type_not_attendee_type, "ErrorInvalidSubfilterTypeNotAttendeeType" },
            { response_code::error_invalid_subfilter_type_not_recipient_type, "ErrorInvalidSubfilterTypeNotRecipientType" },
            { response_code::error_invalid_subscription, "ErrorInvalidSubscription" },
            { response_code::error_invalid_sync_state_data, "ErrorInvalidSyncStateData" },
            { response_code::error_invalid_time_interval, "ErrorInvalidTimeInterval" },
            { response_code::error_invalid_user_oof_settings, "ErrorInvalidUserOofSettings" },
            { response_code::error_invalid_user_principal_name, "ErrorInvalidUserPrincipalName" },
            { response_code::error_invalid_user_sid, "ErrorInvalidUserSid" },
            { response_code::error_invalid_user_sid_missing_upn, "ErrorInvalidUserSidMissingUPN" },
            { response_code::error_invalid_value_for_property, "ErrorInvalidValueForProperty" },
            { response_code::error_invalid_watermark, "ErrorInvalidWatermark" },
            { response_code::error_irresolvable_conflict, "ErrorIrresolvableConflict" },
            { response_code::error_item_corrupt, "ErrorItemCorrupt" },
            { response_code::error_item_not_found, "ErrorItemNotFound" },
            { response_code::error_item_property_request_failed, "ErrorItemPropertyRequestFailed" },
            { response_code::error_item_save, "ErrorItemSave" },
            { response_code::error_item_save_property_error, "ErrorItemSavePropertyError" },
            { response_code::error_legacy_mailbox_free_busy_view_type_not_merged, "ErrorLegacyMailboxFreeBusyViewTypeNotMerged" },
            { response_code::error_local_server_object_not_found, "ErrorLocalServerObjectNotFound" },
            { response_code::error_logon_as_network_service_failed, "ErrorLogonAsNetworkServiceFailed" },
            { response_code::error_mailbox_configuration, "ErrorMailboxConfiguration" },
            { response_code::error_mailbox_data_array_empty, "ErrorMailboxDataArrayEmpty" },
            { response_code::error_mailbox_data_array_too_big, "ErrorMailboxDataArrayTooBig" },
            { response_code::error_mailbox_logon_failed, "ErrorMailboxLogonFailed" },
            { response_code::error_mailbox_move_in_progress, "ErrorMailboxMoveInProgress" },
            { response_code::error_mailbox_store_unavailable, "ErrorMailboxStoreUnavailable" },
            { response_code::error_mail_recipient_not_found, "ErrorMailRecipientNotFound" },
            { response_code::error_managed_folder_already_exists, "ErrorManagedFolderAlreadyExists" },
            { response_code::error_managed_folder_not_found, "ErrorManagedFolderNotFound" },
            { response_code::error_managed_folders_root_failure, "ErrorManagedFoldersRootFailure" },
            { response_code::error_meeting_suggestion_generation_failed, "ErrorMeetingSuggestionGenerationFailed" },
            { response_code::error_message_disposition_required, "ErrorMessageDispositionRequired" },
            { response_code::error_message_size_exceeded, "ErrorMessageSizeExceeded" },
            { response_code::error_mime_content_conversion_failed, "ErrorMimeContentConversionFailed" },
            { response_code::error_mime_content_invalid, "ErrorMimeContentInvalid" },
            { response_code::error_mime_content_invalid_base64_string, "ErrorMimeContentInvalidBase64String" },
            { response_code::error_missing_argument, "ErrorMissingArgument" },
            { response_code::error_missing_email_address, "ErrorMissingEmailAddress" },
            { response_code::error_missing_email_address_for_managed_folder, "ErrorMissingEmailAddressForManagedFolder" },
            { response_code::error_missing_information_email_address, "ErrorMissingInformationEmailAddress" },
            { response_code::error_missing_information_reference_item_id, "ErrorMissingInformationReferenceItemId" },
            { response_code::error_missing_item_for_create_item_attachment, "ErrorMissingItemForCreateItemAttachment" },
            { response_code::error_missing_managed_folder_id, "ErrorMissingManagedFolderId" },
            { response_code::error_missing_recipients, "ErrorMissingRecipients" },
            { response_code::error_move_copy_failed, "ErrorMoveCopyFailed" },
            { response_code::error_move_distinguished_folder, "ErrorMoveDistinguishedFolder" },
            { response_code::error_name_resolution_multiple_results, "ErrorNameResolutionMultipleResults" },
            { response_code::error_name_resolution_no_mailbox, "ErrorNameResolutionNoMailbox" },
            { response_code::error_name_resolution_no_results, "ErrorNameResolutionNoResults" },
            { response_code::error_no_calendar, "ErrorNoCalendar" },
            { response_code::error_no_folder_class_override, "ErrorNoFolderClassOverride" },
            { response_code::error_no_free_busy_access, "ErrorNoFreeBusyAccess" },
            { response_code::error_non_existent_mailbox, "ErrorNonExistentMailbox" },
            { response_code::error_non_primary_smtp_address, "ErrorNonPrimarySmtpAddress" },
            { response_code::error_no_property_tag_for_custom_properties, "ErrorNoPropertyTagForCustomProperties" },
            { response_code::error_not_enough_memory, "ErrorNotEnoughMemory" },
            { response_code::error_object_type_changed, "ErrorObjectTypeChanged" },
            { response_code::error_occurrence_crossing_boundary, "ErrorOccurrenceCrossingBoundary" },
            { response_code::error_occurrence_time_span_too_big, "ErrorOccurrenceTimeSpanTooBig" },
            { response_code::error_parent_folder_id_required, "ErrorParentFolderIdRequired" },
            { response_code::error_parent_folder_not_found, "ErrorParentFolderNotFound" },
            { response_code::error_password_change_required, "ErrorPasswordChangeRequired" },
            { response_code::error_password_expired, "ErrorPasswordExpired" },
            { response_code::error_property_update, "ErrorPropertyUpdate" },
            { response_code::error_property_validation_failure, "ErrorPropertyValidationFailure" },
            { response_code::error_proxy_request_not_allowed, "ErrorProxyRequestNotAllowed" },
            { response_code::error_public_folder_request_processing_failed, "ErrorPublicFolderRequestProcessingFailed" },
            { response_code::error_public_folder_server_not_found, "ErrorPublicFolderServerNotFound" },
            { response_code::error_query_filter_too_long, "ErrorQueryFilterTooLong" },
            { response_code::error_quota_exceeded, "ErrorQuotaExceeded" },
            { response_code::error_read_events_failed, "ErrorReadEventsFailed" },
            { response_code::error_read_receipt_not_pending, "ErrorReadReceiptNotPending" },
            { response_code::error_recurrence_end_date_too_big, "ErrorRecurrenceEndDateTooBig" },
            { response_code::error_recurrence_has_no_occurrence, "ErrorRecurrenceHasNoOccurrence" },
            { response_code::error_request_aborted, "ErrorRequestAborted" },
            { response_code::error_request_stream_too_big, "ErrorRequestStreamTooBig" },
            { response_code::error_required_property_missing, "ErrorRequiredPropertyMissing" },
            { response_code::error_response_schema_validation, "ErrorResponseSchemaValidation" },
            { response_code::error_restriction_too_long, "ErrorRestrictionTooLong" },
            { response_code::error_restriction_too_complex, "ErrorRestrictionTooComplex" },
            { response_code::error_result_set_too_big, "ErrorResultSetTooBig" },
            { response_code::error_saved_item_folder_not_found, "ErrorSavedItemFolderNotFound" },
            { response_code::error_schema_validation, "ErrorSchemaValidation" },
            { response_code::error_search_folder_not_initialized, "ErrorSearchFolderNotInitialized" },
            { response_code::error_send_as_denied, "ErrorSendAsDenied" },
            { response_code::error_send_meeting_cancellations_required, "ErrorSendMeetingCancellationsRequired" },
            { response_code::error_send_meeting_invitations_or_cancellations_required, "ErrorSendMeetingInvitationsOrCancellationsRequired" },
            { response_code::error_send_meeting_invitations_required, "ErrorSendMeetingInvitationsRequired" },
            { response_code::error_sent_meeting_request_update, "ErrorSentMeetingRequestUpdate" },
            { response_code::error_sent_task_request_update, "ErrorSentTaskRequestUpdate" },
            { response_code::error_server_busy, "ErrorServerBusy" },
            { response_code::error_stale_object, "ErrorStaleObject" },
            { response_code::error_subscription_access_denied, "ErrorSubscriptionAccessDenied" },
            { response_code::error_subscription_delegate_access_not_supported, "ErrorSubscriptionDelegateAccessNotSupported" },
            { response_code::error_subscription_not_found, "ErrorSubscriptionNotFound" },
            { response_code::error_sync_folder_not_found, "ErrorSyncFolderNotFound" },
            { response_code::error_time_interval_too_big, "ErrorTimeIntervalTooBig" },
            { response_code::error_to_folder_not_found, "ErrorToFolderNotFound" },
            { response_code::error_token_serialization_denied, "ErrorTokenSerializationDenied" },
            { response_code::error_unable_to_get_user_oof_settings, "ErrorUnableToGetUserOofSettings" },
            { response_code::error_unsupported_culture, "ErrorUnsupportedCulture" },
            { response_code::error_unsupported_mapi_property_type, "ErrorUnsupportedMapiPropertyType" },
            { response_code::error_unsupported_mime_conversion, "ErrorUnsupportedMimeConversion" },
            { response_code::error_unsupported_path_for_query, "ErrorUnsupportedPathForQuery" },
            { response_code::error_unsupported_path_for_sort_group, "ErrorUnsupportedPathForSortGroup" },
            { response_code::error_unsupported_property_definition, "ErrorUnsupportedPropertyDefinition" },
            { response_code::error_unsupported_query_filter, "ErrorUnsupportedQueryFilter" },
            { response_code::error_unsupported_recurrence, "ErrorUnsupportedRecurrence" },
            { response_code::error_unsupported_sub_filter, "ErrorUnsupportedSubFilter" },
            { response_code::error_unsupported_type_for_conversion, "ErrorUnsupportedTypeForConversion" },
            { response_code::error_update_property_mismatch, "ErrorUpdatePropertyMismatch" },
            { response_code::error_virus_detected, "ErrorVirusDetected" },
            { response_code::error_virus_message_deleted, "ErrorVirusMessageDeleted" },
            { response_code::error_voice_mail_not_implemented, "ErrorVoiceMailNotImplemented" },
            { response_code::error_web_request_in_invalid_state, "ErrorWebRequestInInvalidState" },
            { response_code::error_win32_interop_error, "ErrorWin32InteropError" },
            { response_code::error_working_hours_save_failed, "ErrorWorkingHoursSaveFailed" },
            { response_code::error_working_hours_xml_malformed, "ErrorWorkingHoursXmlMalformed" }
        };
        auto it = m.find(code);
        if (it == m.end())
        {
            throw exception("Unrecognized response code");
        }
        return it->second;
    }

    enum class server_version
    {
        // Target the schema files for the initial release version of
        // Exchange 2007
        exchange_2007,

        // Target the schema files for Exchange 2007 Service Pack 1 (SP1),
        // Exchange 2007 Service Pack 2 (SP2), and
        // Exchange 2007 Service Pack 3 (SP3)
        exchange_2007_sp1,

        // Target the schema files for Exchange 2010
        exchange_2010,

        // Target the schema files for Exchange 2010 Service Pack 1 (SP1)
        exchange_2010_sp1,

        // Target the schema files for Exchange 2010 Service Pack 2 (SP2)
        // and Exchange 2010 Service Pack 3 (SP3)
        exchange_2010_sp2,

        // Target the schema files for Exchange 2013
        exchange_2013,

        // Target the schema files for Exchange 2013 Service Pack 1 (SP1)
        exchange_2013_sp1
    };

    namespace internal
    {
        inline std::string server_version_to_str(server_version vers)
        {
            switch (vers)
            {
                case server_version::exchange_2007:
                    return "Exchange2007";
                case server_version::exchange_2007_sp1:
                    return "Exchange2007_SP1";
                case server_version::exchange_2010:
                    return "Exchange2010";
                case server_version::exchange_2010_sp1:
                    return "Exchange2010_SP1";
                case server_version::exchange_2010_sp2:
                    return "Exchange2010_SP2";
                case server_version::exchange_2013:
                    return "Exchange2013";
                case server_version::exchange_2013_sp1:
                    return "Exchange2013_SP1";
                default:
                    throw exception("Bad enum value");
            }
        }

        inline server_version str_to_server_version(const std::string& str)
        {
            if (str == "Exchange2007")
            {
                return server_version::exchange_2007;
            }
            else if (str == "Exchange2007_SP1")
            {
                return server_version::exchange_2007_sp1;
            }
            else if (str == "Exchange2010")
            {
                return server_version::exchange_2010;
            }
            else if (str == "Exchange2010_SP1")
            {
                return server_version::exchange_2010_sp1;
            }
            else if (str == "Exchange2010_SP2")
            {
                return server_version::exchange_2010_sp2;
            }
            else if (str == "Exchange2013")
            {
                return server_version::exchange_2013;
            }
            else if (str == "Exchange2013_SP1")
            {
                return server_version::exchange_2013_sp1;
            }
            else
            {
                throw exception("Unrecognized <RequestServerVersion>");
            }
        }
    }

    enum class base_shape { id_only, default_shape, all_properties };

    // TODO: move to internal namespace
    inline std::string base_shape_str(base_shape shape)
    {
        switch (shape)
        {
            case base_shape::id_only:        return "IdOnly";
            case base_shape::default_shape:  return "Default";
            case base_shape::all_properties: return "AllProperties";
            default: throw exception("Bad enum value");
        }
    }

    // Side note: we do not provide SoftDelete because that does not make much
    // sense from an EWS perspective
    enum class delete_type { hard_delete, move_to_deleted_items };

    // TODO: move to internal namespace
    inline std::string delete_type_str(delete_type d)
    {
        switch (d)
        {
            case delete_type::hard_delete:
                return "HardDelete";
            case delete_type::move_to_deleted_items:
                return "MoveToDeletedItems";
            default:
                throw exception("Bad enum value");
        }
    }

    enum class affected_task_occurrences
    {
        all_occurrences,
        specified_occurrence_only
    };

    // TODO: move to internal namespace
    inline std::string
    affected_task_occurrences_str(affected_task_occurrences o)
    {
        switch (o)
        {
            case affected_task_occurrences::all_occurrences:
                return "AllOccurrences";
            case affected_task_occurrences::specified_occurrence_only:
                return "SpecifiedOccurrenceOnly";
            default:
                throw exception("Bad enum value");
        }
    }

    enum class conflict_resolution
    {
        // If there is a conflict, the update operation fails and an error is
        // returned. The call to update_item never overwrites data that has
        // changed underneath you!
        never_overwrite,

        // The update operation automatically resolves any conflict (if it can,
        // otherwise the request fails)
        auto_resolve,

        // If there is a conflict, the update operation will overwrite
        // information. Ignores changes that occurred underneath you; last
        // writer wins!
        always_overwrite
    };

    // TODO: move to internal namespace
    inline std::string conflict_resolution_str(conflict_resolution val)
    {
        switch (val)
        {
            case conflict_resolution::never_overwrite:
                return "NeverOverwrite";
            case conflict_resolution::auto_resolve:
                return "AutoResolve";
            case conflict_resolution::always_overwrite:
                return "AlwaysOverwrite";
            default:
                throw exception("Bad enum value");
        }
    }

    // <CreateItem> and <UpdateItem> methods use this attribute. Only applicable
    // to e-mail messages
    enum class message_disposition
    {
        // Save the message in a specified folder or in the Drafts folder if
        // none is given
        save_only,

        // Send the message and do not save a copy in the sender's mailbox
        send_only,

        // Send the message and save a copy in a specified folder or in the
        // mailbox owner's Sent Items folder if none is given
        send_and_save_copy
    };

    // TODO: move to internal namespace
    inline std::string message_disposition_str(message_disposition val)
    {
        switch (val)
        {
            case message_disposition::save_only:
                return "SaveOnly";
            case message_disposition::send_only:
                return "SendOnly";
            case message_disposition::send_and_save_copy:
                return "SendAndSaveCopy";
            default:
                throw exception("Bad enum value");
        }
    }

    // Exception thrown when a request was not successful
    class exchange_error final : public exception
    {
    public:
        explicit exchange_error(response_code code)
            : exception(response_code_to_str(code)), code_(code)
        {
        }

        response_code code() const EWS_NOEXCEPT { return code_; }

    private:
        response_code code_;
    };

    // A SOAP fault occurred due to a bad request
    class soap_fault : public exception
    {
    public:
        explicit soap_fault(const std::string& what) : exception(what) {}
        explicit soap_fault(const char* what) : exception(what) {}
    };

    // A SOAP fault that is raised when we sent invalid XML.
    //
    // This is an internal error and indicates a bug in this library, thus
    // should never happen.
    //
    // Note: because this exception is due to a SOAP fault (sometimes recognized
    // before any server-side XML parsing finished) any included failure message
    // is likely not localized according to any MailboxCulture SOAP header.
    class schema_validation_error final : public soap_fault
    {
    public:
        schema_validation_error(unsigned long line_number,
                                unsigned long line_position,
                                std::string violation)
            : soap_fault("The request failed schema validation"),
              violation_(std::move(violation)),
              line_pos_(line_position),
              line_no_(line_number)
        {
        }

        // Line number in request string where the error was found
        unsigned long line_number() const EWS_NOEXCEPT { return line_no_; }

        // Column number in request string where the error was found
        unsigned long line_position() const EWS_NOEXCEPT { return line_pos_; }

        // A more detailed explanation of what went wrong
        const std::string& violation() const EWS_NOEXCEPT { return violation_; }

    private:
        std::string violation_;
        unsigned long line_pos_;
        unsigned long line_no_;
    };

    namespace internal
    {
        // Exception for libcurl related runtime errors
        class curl_error final : public exception
        {
        public:
            explicit curl_error(const std::string& what)
                : exception(what)
            {
            }
            explicit curl_error(const char* what) : exception(what) {}
        };

        // Helper function; constructs an exception with a meaningful error
        // message from the given result code for the most recent cURL API call.
        //
        // msg: A string that prepends the actual cURL error message.
        // rescode: The result code of a failed cURL operation.
        inline curl_error make_curl_error(const std::string& msg,
                                          CURLcode rescode)
        {
            std::string reason{ curl_easy_strerror(rescode) };
#ifdef NDEBUG
            (void)msg;
            return curl_error(reason);
#else
            return curl_error(msg + ": \'" + reason + "\'");
#endif
        }

        // RAII helper class for CURL* handles.
        class curl_ptr final
        {
        public:
            curl_ptr() : handle_{ curl_easy_init() }
            {
                if (!handle_)
                {
                    throw curl_error{ "Could not start libcurl session" };
                }
            }

            ~curl_ptr() { curl_easy_cleanup(handle_); }

            // Could use curl_easy_duphandle for copying
            curl_ptr(const curl_ptr&) = delete;
            curl_ptr& operator=(const curl_ptr&) = delete;

            CURL* get() const EWS_NOEXCEPT { return handle_; }

        private:
            CURL* handle_;
        };

        // RAII wrapper class around cURLs slist construct.
        class curl_string_list final
        {
        public:
            curl_string_list() EWS_NOEXCEPT : slist_{ nullptr } {}

            ~curl_string_list() { curl_slist_free_all(slist_); }

            void append(const char* str) EWS_NOEXCEPT
            {
                slist_ = curl_slist_append(slist_, str);
            }

            curl_slist* get() const EWS_NOEXCEPT { return slist_; }

        private:
            curl_slist* slist_;
        };

        // String constants
        // TODO: sure this can't be done easier within a header file?
        // We need better handling of static strings (URIs, XML node names,
        // values and stuff, see usage of rapidxml::internal::compare)
        template <typename Ch = char> struct uri
        {
            struct microsoft
            {
                enum { errors_size = 58U };
                static const Ch* errors()
                {
                    static const Ch* val =
                        "http://schemas.microsoft.com/exchange/services/2006/errors";
                    return val;
                }
                enum { types_size = 57U };
                static const Ch* types()
                {
                    static const Ch* const val =
                        "http://schemas.microsoft.com/exchange/services/2006/types";
                    return val;
                }
                enum { messages_size = 60U };
                static const Ch* messages()
                {
                    static const Ch* const val =
                        "http://schemas.microsoft.com/exchange/services/2006/messages";
                    return val;
                }
            };
            struct soapxml
            {
                enum { envelope_size = 41U };
                static const Ch* envelope()
                {
                    static const Ch* const val =
                        "http://schemas.xmlsoap.org/soap/envelope/";
                    return val;
                }
            };
        };

        // This ought to be a DOM wrapper; usually around a web response
        //
        // This class basically wraps rapidxml::xml_document because the parsed
        // data must persist for the lifetime of the rapidxml::xml_document.
        class http_response final
        {
        public:
            http_response(long code, std::vector<char>&& data)
                : data_(std::move(data)),
                  doc_(new rapidxml::xml_document<char>()),
                  code_(code),
                  parsed_(false)
            {
                EWS_ASSERT(!data_.empty());
            }

            ~http_response() = default;

            http_response(const http_response&) = delete;
            http_response& operator=(const http_response&) = delete;

            http_response(http_response&& other)
                : data_(std::move(other.data_)),
                  doc_(std::move(other.doc_)),
                  code_(std::move(other.code_)),
                  parsed_(std::move(other.parsed_))
            {
                other.code_ = 0U;
            }

            http_response& operator=(http_response&& rhs)
            {
                if (&rhs != this)
                {
                    data_ = std::move(rhs.data_);
                    doc_ = std::move(rhs.doc_);
                    code_ = std::move(rhs.code_);
                    parsed_ = std::move(rhs.parsed_);
                }
                return *this;
            }

            // Returns the SOAP payload in this response.
            //
            // Parses the payload (if it hasn't already) and returns it as
            // xml_document.
            //
            // Note: we are using a mutable temporary buffer internally because
            // we are using RapidXml in destructive mode (the parser modifies
            // source text during the parsing process). Hence, we need to make
            // sure that parsing is done only once!
            const rapidxml::xml_document<char>& payload()
            {
                // FIXME: data race, synchronization missing, read of parsed_ and call parse()
                // TODO: parsed_ boolean not necessary, use if (doc_) { /* ... */ }
                if (!parsed_)
                {
                    on_scope_exit set_parsed([&]
                    {
                        parsed_ = true;
                    });
                    parse();
                }
                return *doc_;
            }

            // Returns the response code of the HTTP request.
            long code() const EWS_NOEXCEPT
            {
                return code_;
            }

            // Returns whether the response is a SOAP fault.
            //
            // This means the server responded with status code 500 and
            // indicates that the entire request failed (not just a normal EWS
            // error). This can happen e.g. when the request we sent was not
            // schema compliant.
            bool is_soap_fault() const EWS_NOEXCEPT
            {
                return code() == 500U;
            }

            // Returns whether the HTTP response code is 200 (OK).
            bool ok() const EWS_NOEXCEPT
            {
                return code() == 200U;
            }

        private:
            // Here we handle the server's response. We load the SOAP payload
            // from response into the xml_document.
            void parse()
            {
                try
                {
                    static const int flags = 0;
                    doc_->parse<flags>(&data_[0]);
                }
                catch (rapidxml::parse_error& e)
                {
                    // Swallow and erase type
                    throw parse_error(e.what());
                }
            }

            std::vector<char> data_;
            std::unique_ptr<rapidxml::xml_document<char>> doc_;
            long code_;
            bool parsed_;
        };

        // Traverse elements, depth first, beginning with given node.
        //
        // Applies given function to every element during traversal, stopping as
        // soon as that function returns true.
        template <typename Function>
        inline void traverse_elements(const rapidxml::xml_node<>& node,
                                      Function func) EWS_NOEXCEPT
        {
            for (auto child = node.first_node(); child != nullptr;
                 child = child->next_sibling())
            {
                traverse_elements(*child, func);
                if (child->type() == rapidxml::node_element)
                {
                    if (func(*child))
                    {
                        return;
                    }
                }
            }
        }

        // Select element by qualified name, nullptr if there is no such element
        // TODO: what if namespace_uri is empty, can we do that cheaper?
        inline rapidxml::xml_node<>*
        get_element_by_qname(const rapidxml::xml_node<>& node,
                             const char* local_name,
                             const char* namespace_uri) EWS_NOEXCEPT
        {
            EWS_ASSERT(local_name && namespace_uri);

            rapidxml::xml_node<>* element = nullptr;
            const auto local_name_len = std::strlen(local_name);
            const auto namespace_uri_len = std::strlen(namespace_uri);
            traverse_elements(node, [&](rapidxml::xml_node<>& elem) -> bool
            {
                using rapidxml::internal::compare;

                if (compare(elem.namespace_uri(),
                            elem.namespace_uri_size(),
                            namespace_uri,
                            namespace_uri_len)
                    &&
                    compare(elem.local_name(),
                            elem.local_name_size(),
                            local_name,
                            local_name_len))
                {
                    element = std::addressof(elem);
                    return true;
                }
                return false;
            });
            return element;
        }

        // Does nothing if given response is not a SOAP fault
        inline void raise_exception_if_soap_fault(http_response& response)
        {
            using rapidxml::internal::compare;

            if (!response.is_soap_fault())
            {
                return;
            }
            const auto& doc = response.payload();
            auto elem = get_element_by_qname(doc,
                                             "ResponseCode",
                                             uri<>::microsoft::errors());
            if (!elem)
            {
                throw soap_fault(
                    "The request failed for unknown reason (no XML in response)"
                );
                // TODO: what about getting information from HTTP headers
            }

            if (compare(elem->value(),
                        elem->value_size(),
                        "ErrorSchemaValidation",
                        21))
            {
                // Get some more helpful details
                elem = get_element_by_qname(doc,
                                            "LineNumber",
                                            uri<>::microsoft::types());
                EWS_ASSERT(elem &&
                        "Expected <LineNumber> element in response");
                const auto line_number =
                    std::stoul(std::string(elem->value(), elem->value_size()));

                elem = get_element_by_qname(doc,
                                            "LinePosition",
                                            uri<>::microsoft::types());
                EWS_ASSERT(elem &&
                        "Expected <LinePosition> element in response");
                const auto line_position =
                    std::stoul(std::string(elem->value(), elem->value_size()));

                elem = get_element_by_qname(doc,
                                            "Violation",
                                            uri<>::microsoft::types());
                EWS_ASSERT(elem &&
                        "Expected <Violation> element in response");
                throw schema_validation_error(
                        line_number,
                        line_position,
                        std::string(elem->value(), elem->value_size()));
            }
            else
            {
                elem = get_element_by_qname(doc, "faultstring", "");
                EWS_ASSERT(elem &&
                        "Expected <faultstring> element in response");
                throw soap_fault(elem->value());
            }
        }

        class credentials
        {
        public:
            virtual ~credentials() = default;
            virtual void certify(http_request*) const = 0;
        };

        class ntlm_credentials final : public credentials
        {
        public:
            ntlm_credentials(std::string username, std::string password,
                             std::string domain)
                : username_(std::move(username)),
                  password_(std::move(password)),
                  domain_(std::move(domain))
            {
            }

        private:
            void certify(http_request*) const override; // implemented below

            std::string username_;
            std::string password_;
            std::string domain_;
        };

        class http_request final
        {
        public:
            enum class method
            {
                POST
            };

            // Create a new HTTP request to the given URL.
            explicit http_request(const std::string& url)
            {
                set_option(CURLOPT_URL, url.c_str());
            }

            // Set the HTTP method (only POST supported).
            void set_method(method)
            {
                // Method can only be a regular POST in our use case
                set_option(CURLOPT_POST, 1);
            }

            // Set this HTTP request's content type.
            void set_content_type(const std::string& content_type)
            {
                const std::string str = "Content-Type: " + content_type;
                headers_.append(str.c_str());
            }

            // Set credentials for authentication.
            void set_credentials(const credentials& creds)
            {
                creds.certify(this);
            }

            // Small wrapper around curl_easy_setopt(3).
            //
            // Converts return codes to exceptions of type curl_error and allows
            // objects other than http_web_requests to set transfer options
            // without direct access to the internal CURL handle.
            template <typename... Args>
            void set_option(CURLoption option, Args... args)
            {
                auto retcode = curl_easy_setopt(
                    handle_.get(), option, std::forward<Args>(args)...);
                switch (retcode)
                {
                case CURLE_OK:
                    return;

                case CURLE_FAILED_INIT:
                {
                    throw make_curl_error(
                        "curl_easy_setopt: unsupported option", retcode);
                }

                default:
                {
                    throw make_curl_error(
                        "curl_easy_setopt: failed setting option", retcode);
                }
                };
            }

            // Perform the HTTP request and returns the response. This function
            // blocks until the complete response is received or a timeout is
            // reached. Throws curl_error if operation could not be completed.
            //
            // request: The complete request string; you must make sure that
            // the data is encoded the way you want the server to receive it.
            http_response send(const std::string& request)
            {
                auto callback =
                    [](char* ptr, std::size_t size, std::size_t nmemb,
                       void* userdata) -> std::size_t
                {
                    std::vector<char>* buf
                        = reinterpret_cast<std::vector<char>*>(userdata);
                    const auto realsize = size * nmemb;
                    try
                    {
                        buf->reserve(realsize + 1); // plus 0-terminus
                    }
                    catch (std::bad_alloc&)
                    {
                        // Out of memory, indicate error to libcurl
                        return 0U;
                    }
                    std::copy(ptr, ptr + realsize, std::back_inserter(*buf));
                    return realsize;
                };

                // Do not install (directly or indirectly) signal handlers nor
                // call any functions that cause signals to be sent to the
                // process
                // Note: SIGCHLD is raised anyway if we use CURLAUTH_NTLM_WB and
                // SIGPIPE is still possible, too
                set_option(CURLOPT_NOSIGNAL, 1L);

#ifdef EWS_ENABLE_VERBOSE
                // Print HTTP headers to stderr
                set_option(CURLOPT_VERBOSE, 1L);
# endif

                // Set complete request string for HTTP POST method; note: no
                // encoding here
                set_option(CURLOPT_POSTFIELDS, request.c_str());
                set_option(CURLOPT_POSTFIELDSIZE, request.length());

                // Finally, set HTTP headers. We do this as last action here
                // because we want to overwrite implicitly set header lines due
                // to the options set above with our own header lines
                set_option(CURLOPT_HTTPHEADER, headers_.get());

                std::vector<char> response_data;
                set_option(
                    CURLOPT_WRITEFUNCTION,
                    static_cast<std::size_t (*)(char*, std::size_t, std::size_t,
                                                void*)>(callback));
                set_option(CURLOPT_WRITEDATA,
                        std::addressof(response_data));

#ifndef NDEBUG
                // Turn-off verification of the server's authenticity
                set_option(CURLOPT_SSL_VERIFYPEER, 0);
#endif

                auto retcode = curl_easy_perform(handle_.get());
                if (retcode != 0)
                {
                    throw make_curl_error("curl_easy_perform", retcode);
                }
                long response_code = 0U;
                curl_easy_getinfo(handle_.get(),
                    CURLINFO_RESPONSE_CODE, &response_code);
                response_data.emplace_back('\0');
                return http_response(response_code, std::move(response_data));
            }

        private:
            curl_ptr handle_;
            curl_string_list headers_;
        };

        // Makes a raw SOAP request.
        //
        // url: The URL of the server to talk to.
        // username: The username of user.
        // password: The user's secret password, plain-text.
        // domain: The user's Windows domain.
        // soap_body: The contents of the SOAP body (minus the body element);
        // this is the actual EWS request.
        // soap_headers: Any SOAP headers to add.
        //
        // Returns the response.
        template <typename RequestHandler = http_request>
        inline http_response make_raw_soap_request(
            const std::string& url,
            const std::string& username,
            const std::string& password,
            const std::string& domain,
            const std::string& soap_body,
            const std::vector<std::string>& soap_headers)
        {
            RequestHandler request{url};
            request.set_method(RequestHandler::method::POST);
            request.set_content_type("text/xml; charset=utf-8");

            ntlm_credentials creds{username, password, domain};
            request.set_credentials(creds);

            std::stringstream request_stream;
            request_stream <<
R"(<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
>)";

            // Add SOAP headers if present
            if (!soap_headers.empty())
            {
                request_stream << "<soap:Header>\n";
                for (const auto& header : soap_headers)
                {
                    request_stream << header;
                }
                request_stream << "</soap:Header>\n";
            }

            request_stream << "<soap:Body>\n";
            // Add the passed request
            request_stream << soap_body;
            request_stream << "</soap:Body>\n";
            request_stream << "</soap:Envelope>\n";

#if EWS_ENABLE_VERBOSE
            std::cerr << request_stream.str() << std::endl;
#endif

            return request.send(request_stream.str());
        }

        // Parse response class and response code from given element.
        inline std::pair<response_class, response_code>
        parse_response_class_and_code(const rapidxml::xml_node<>& elem)
        {
            using uri = internal::uri<>;
            using rapidxml::internal::compare;

            auto cls = response_class::success;
            auto response_class_attr = elem.first_attribute("ResponseClass");
            if (compare(response_class_attr->value(),
                        response_class_attr->value_size(),
                        "Error",
                        5))
            {
                cls = response_class::error;
            }
            else if (compare(response_class_attr->value(),
                             response_class_attr->value_size(),
                             "Warning",
                             7))
            {
                cls = response_class::warning;
            }

            // One thing we can count on is that when the ResponseClass
            // attribute is set to Success, ResponseCode will be set to NoError.
            // So we only parse the <ResponseCode> element when we have a
            // warning or an error.

            auto code = response_code::no_error;
            if (cls != response_class::success)
            {
                auto response_code_elem =
                    elem.first_node_ns(uri::microsoft::messages(),
                                        "ResponseCode");
                EWS_ASSERT(response_code_elem
                        && "Expected <ResponseCode> element");
                code = str_to_response_code(response_code_elem->value());
            }

            return std::make_pair(cls, code);
        }

        // Iterate over <Items> array and execute given function for each node.
        //
        // elem: a response message element, e.g., CreateItemResponseMessage
        // func: A callable that is invoked for each item in the response
        // message's <Items> array. A const rapidxml::xml_node& is passed to
        // that callable.
        template <typename Func>
        inline void for_each_item(const rapidxml::xml_node<>& elem, Func func)
        {
            using uri = internal::uri<>;

            auto items_elem =
                elem.first_node_ns(uri::microsoft::messages(), "Items");
            EWS_ASSERT(items_elem && "Expected <Items> element");

            for (auto item_elem = items_elem->first_node(); item_elem;
                 item_elem = item_elem->next_sibling())
            {
                EWS_ASSERT(item_elem && "Expected an element, got nullptr");
                func(*item_elem);
            }
        }

        // Base-class for all response messages
        class response_message_base
        {
        public:
            response_message_base(response_class cls, response_code code)
                : cls_(cls),
                  code_(code)
            {
            }

            response_class get_response_class() const EWS_NOEXCEPT
            {
                return cls_;
            }

            bool success() const EWS_NOEXCEPT
            {
                return get_response_class() == response_class::success;
            }

            response_code get_response_code() const EWS_NOEXCEPT
            {
                return code_;
            }

        private:
            response_class cls_;
            response_code code_;
        };

        // Base-class for response messages that contain an <Items> array.
        //
        // The ItemType template parameter denotes the type of all items in the
        // returned array. The choice for a compile-time parameter has following
        // implications and restrictions:
        //
        // - Microsoft EWS allows for different types of items in the returned
        //   array. However, this implementation forces you to only issue
        //   requests that return only one type of item in a single response at
        //   a time.
        //
        // - You need to know the type of the item returned by a request
        //   up-front at compile time. Microsoft EWS would allow to deal with
        //   different types of items in a single response dynamically.
        template <typename ItemType>
        class response_message_with_items : public response_message_base
        {
        public:
            typedef ItemType item_type;

            response_message_with_items(response_class cls,
                                        response_code code,
                                        std::vector<item_type> items)
                : response_message_base(cls, code),
                  items_(std::move(items))
            {
            }

            const std::vector<item_type>& items() const EWS_NOEXCEPT
            {
                return items_;
            }

        private:
            std::vector<item_type> items_;
        };

        class create_item_response_message final
            : public response_message_with_items<item_id>
        {
        public:
            // implemented below
            static create_item_response_message parse(http_response&);

        private:
            create_item_response_message(response_class cls,
                                         response_code code,
                                         std::vector<item_id> items)
                : response_message_with_items<item_id>(cls,
                                                       code,
                                                       std::move(items))
            {
            }
        };

        class find_item_response_message final
            : public response_message_with_items<item_id>
        {
        public:
            // implemented below
            static find_item_response_message parse(http_response&);

        private:
            find_item_response_message(response_class cls,
                                       response_code code,
                                       std::vector<item_id> items)
                : response_message_with_items<item_id>(cls,
                  code,
                  std::move(items))
            {
            }
        };

        class update_item_response_message final
            : public response_message_with_items<item_id>
        {
        public:
            // implemented below
            static update_item_response_message parse(http_response&);

        private:
            update_item_response_message(response_class cls,
                                       response_code code,
                                       std::vector<item_id> items)
                : response_message_with_items<item_id>(cls,
                  code,
                  std::move(items))
            {
            }
        };

        template <typename ItemType>
        class get_item_response_message final
            : public response_message_with_items<ItemType>
        {
        public:
            // implemented below
            static get_item_response_message parse(http_response&);

        private:
            get_item_response_message(response_class cls,
                                      response_code code,
                                      std::vector<ItemType> items)
                : response_message_with_items<ItemType>(cls,
                                                        code,
                                                        std::move(items))
            {
            }
        };

        class delete_item_response_message final : public response_message_base
        {
        public:
            static delete_item_response_message parse(http_response& response)
            {
                using uri = internal::uri<>;
                using internal::get_element_by_qname;

                const auto& doc = response.payload();
                auto elem = get_element_by_qname(doc,
                                                 "DeleteItemResponseMessage",
                                                 uri::microsoft::messages());

#ifdef EWS_ENABLE_VERBOSE
            if (!elem)
            {
                std::cerr
                    << "Parsing DeleteItemResponseMessage failed, response code: "
                    << response.code() << ", payload:\n" << doc << std::endl;
            }
#endif

                EWS_ASSERT(elem &&
                        "Expected <DeleteItemResponseMessage>, got nullptr");
                const auto result = parse_response_class_and_code(*elem);
                return delete_item_response_message(result.first,
                                                    result.second);
            }

        private:
            delete_item_response_message(response_class cls, response_code code)
                : response_message_base(cls, code)
            {
            }
        };
    }

    // Function is not thread-safe; should be set-up when application is still
    // in single-threaded context. Calling this function more than once does no
    // harm.
    inline void set_up() { curl_global_init(CURL_GLOBAL_DEFAULT); }

    // Function is not thread-safe; you should call this function only when no
    // other thread is running (see libcurl(3) man-page or
    // http://curl.haxx.se/libcurl/c/libcurl.html)
    inline void tear_down() EWS_NOEXCEPT { curl_global_cleanup(); }

    // Contains the unique identifier and change key of an item in the Exchange
    // store.
    //
    // Instances of this class are somewhat immutable. You can default construct
    // an item_id in which case valid() will always return false. (Default
    // construction is needed because we need item and it's sub-classes to be
    // default constructible.) Only item_ids that come from an Exchange store
    // are considered to be valid.
    class item_id final
    {
    public:
        item_id() = default;

        explicit item_id(std::string id)
            : id_(std::move(id)),
              change_key_()
        {}

        item_id(std::string id, std::string change_key)
            : id_(std::move(id)),
              change_key_(std::move(change_key))
        {}

        const std::string& id() const EWS_NOEXCEPT { return id_; }

        const std::string& change_key() const EWS_NOEXCEPT
        {
            return change_key_;
        }

        bool valid() const EWS_NOEXCEPT { return !id_.empty(); }

        std::string to_xml(const char* xmlns=nullptr) const
        {
            if (xmlns)
            {
                return std::string("<") + xmlns + ":ItemId Id=\"" + id() +
                    "\" ChangeKey=\"" + change_key() + "\"/>";
            }
            else
            {
                return "<ItemId Id=\"" + id() +
                    "\" ChangeKey=\"" + change_key() + "\"/>";
            }
        }

        // Makes an item_id instance from an <ItemId> XML element
        static item_id from_xml_element(const rapidxml::xml_node<>& elem)
        {
            auto id_attr = elem.first_attribute("Id");
            EWS_ASSERT(id_attr && "Missing attribute Id in <ItemId>");
            auto id = std::string(id_attr->value(), id_attr->value_size());
            auto ckey_attr = elem.first_attribute("ChangeKey");
            EWS_ASSERT(ckey_attr && "Missing attribute ChangeKey in <ItemId>");
            auto ckey = std::string(ckey_attr->value(),
                    ckey_attr->value_size());
            return item_id(std::move(id), std::move(ckey));
        }

    private:
        // case-sensitive; therefore, comparisons between IDs must be
        // case-sensitive or binary
        std::string id_;

        // Identifies a specific version of an item.
        std::string change_key_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<item_id>::value, "");
    static_assert(std::is_copy_constructible<item_id>::value, "");
    static_assert(std::is_copy_assignable<item_id>::value, "");
    static_assert(std::is_move_constructible<item_id>::value, "");
    static_assert(std::is_move_assignable<item_id>::value, "");
#endif

    // Contains the unique identifier of an attachment.
    class attachment_id final
    {
    public:
        attachment_id() = default;

        explicit attachment_id(std::string id) : id_(std::move(id)) {}

        attachment_id(std::string id, item_id root_item_id)
            : id_(std::move(id)),
              root_item_id_(std::move(root_item_id))
        {}

        const std::string& id() const EWS_NOEXCEPT { return id_; }

        const item_id& root_item_id() const EWS_NOEXCEPT
        {
            return root_item_id_;
        }

        bool valid() const EWS_NOEXCEPT { return !id_.empty(); }

        std::string to_xml(const char* xmlns=nullptr) const
        {
            auto pref = std::string();
            if (xmlns)
            {
                pref = std::string(xmlns) + ":";
            }
            std::stringstream sstr;
            sstr << "<" << pref << "AttachmentId Id=\"" << id() << "\"";
            if (root_item_id().valid())
            {
                sstr << " RootItemId=\"" << root_item_id().id()
                     << "\" RootItemChangeKey=\"" << root_item_id().change_key()
                     << "\"";
            }
            sstr << "/>";
            return sstr.str();
        }

        // Makes an attachment_id instance from an <AttachmentId> element
        static attachment_id from_xml_element(const rapidxml::xml_node<>& elem)
        {
            auto id_attr = elem.first_attribute("Id");
            EWS_ASSERT(id_attr && "Missing attribute Id in <AttachmentId>");
            auto id = std::string(id_attr->value(), id_attr->value_size());
            auto root_item_id = std::string();
            auto root_item_ckey = std::string();

            auto root_item_id_attr = elem.first_attribute("RootItemId");
            if (root_item_id_attr)
            {
                root_item_id =
                    std::string(root_item_id_attr->value(),
                                root_item_id_attr->value_size());
                auto root_item_ckey_attr =
                    elem.first_attribute("RootItemChangeKey");
                EWS_ASSERT(root_item_ckey_attr
                        && "Expected attribute RootItemChangeKey");
                root_item_ckey = std::string(root_item_ckey_attr->value(),
                                             root_item_ckey_attr->value_size());
            }

            return root_item_id.empty() ?
                   attachment_id(std::move(id)) :
                   attachment_id(std::move(id),
                                 item_id(std::move(root_item_id),
                                         std::move(root_item_ckey)));
        }

    private:
        std::string id_;
        item_id root_item_id_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<attachment_id>::value, "");
    static_assert(std::is_copy_constructible<attachment_id>::value, "");
    static_assert(std::is_copy_assignable<attachment_id>::value, "");
    static_assert(std::is_move_constructible<attachment_id>::value, "");
    static_assert(std::is_move_assignable<attachment_id>::value, "");
#endif

    namespace internal
    {
        // A self-contained copy of a DOM sub-tree generally used to hold
        // properties of an item class
        //
        // We use rapidxml::print for copying the already parsed xml_document
        // into a new buffer and then re-parse this buffer again under the
        // assumption that all child elements are contained in
        //
        //     http://schemas.microsoft.com/exchange/services/2006/types
        //
        // XML namespace.
        //
        // A default constructed item_properties instance makes only sense when
        // an item class is default constructed. In that case the buffer (and
        // the DOM) is initially empty and elements are added directly to the
        // document's root node.
        class item_properties final
        {
        public:
            // Default constructible because item class (and it's descendants)
            // need to be
            item_properties()
                : rawdata_(),
                  doc_(new rapidxml::xml_document<char>)
            {}

            explicit item_properties(const rapidxml::xml_node<char>& origin,
                                 std::size_t size_hint=0U)
                : rawdata_(),
                  doc_(new rapidxml::xml_document<char>)
            {
                reparse(origin, size_hint);
            }

            // Needs to be copy- and copy-assignable because item classes are.
            // However xml_document isn't (and can't be without major rewrite
            // IMHO).  Hence, copying is quite costly as it involves serializing
            // DOM tree to string and re-parsing into new tree
            item_properties& operator=(const item_properties& rhs)
            {
                if (&rhs != this)
                {
                    // FIXME: not strong exception safe
                    rawdata_.clear();
                    doc_->clear();
                    reparse(*rhs.doc_, rhs.rawdata_.size());
                }
                return *this;
            }

            item_properties(const item_properties& other)
                : rawdata_(),
                  doc_(new rapidxml::xml_document<char>)
            {
                reparse(*other.doc_, other.rawdata_.size());
            }

            rapidxml::xml_node<char>& root() EWS_NOEXCEPT { return *doc_; }

            const rapidxml::xml_node<char>& root() const EWS_NOEXCEPT
            {
                return *doc_;
            }

            // Might return nullptr when there is no such element. Client code
            // needs to check returned pointer. Should never throw
            rapidxml::xml_node<char>*
            get_node(const char* node_name) const EWS_NOEXCEPT
            {
                return get_element_by_qname(*doc_,
                                            node_name,
                                            uri<>::microsoft::types());
            }

            rapidxml::xml_document<char>* document() const EWS_NOEXCEPT
            {
                return doc_.get();
            }

            std::string get_value_as_string(const char* node_name) const
            {
                const auto node = get_node(node_name);
                return node ?
                      std::string(node->value(), node->value_size())
                    : "";
            }

            void set_or_update(std::string node_name, std::string node_value)
            {
                using uri = internal::uri<>;
                using rapidxml::internal::compare;

                auto oldnode = get_element_by_qname(*doc_,
                                                    node_name.c_str(),
                                                    uri::microsoft::types());
                if (oldnode && compare(node_value.c_str(),
                                       node_value.length(),
                                       oldnode->value(),
                                       oldnode->value_size()))
                {
                    // Nothing to do
                    return;
                }

                auto node_qname = "t:" + node_name;
                auto ptr_to_qname = doc_->allocate_string(node_qname.c_str());
                auto ptr_to_value = doc_->allocate_string(node_value.c_str());

                // Strong exception-safety guarantee? Memory isn't leaked, OTOH
                // it is not freed until document (or the item that owns it) is
                // destructed either

                auto newnode = doc_->allocate_node(rapidxml::node_element);
                newnode->qname(ptr_to_qname,
                               node_qname.length(),
                               ptr_to_qname + 2);
                newnode->value(ptr_to_value);
                newnode->namespace_uri(uri::microsoft::types(),
                                       uri::microsoft::types_size);
                if (oldnode)
                {
                    doc_->first_node()->insert_node(oldnode, newnode);
                    doc_->first_node()->remove_node(oldnode);
                }
                else
                {
                    doc_->append_node(newnode);
                }
            }

            std::string to_string() const
            {
                std::string str;
                rapidxml::print(std::back_inserter(str),
                                *doc_,
                                rapidxml::print_no_indenting);
                return str;
            }

        private:
            std::vector<char> rawdata_;
            std::unique_ptr<rapidxml::xml_document<char>> doc_;

            // Custom namespace processor for parsing XML sub-tree. Makes
            // uri::microsoft::types the default namespace.
            struct custom_namespace_processor
                : public rapidxml::internal::xml_namespace_processor<char>
            {
                struct scope
                    : public rapidxml::internal::xml_namespace_processor<char>::scope
                {
                    explicit scope(xml_namespace_processor& processor)
                        : rapidxml::internal::xml_namespace_processor<char>::scope(processor)
                    {
                        static auto nsattr = make_namespace_attribute();
                        add_namespace_prefix(&nsattr);
                    }

                    static
                    rapidxml::xml_attribute<char> make_namespace_attribute()
                    {
                        static const char* const name = "t";
                        static const char* const value =
                            uri<>::microsoft::types();
                        rapidxml::xml_attribute<char> attr;
                        attr.name(name, 1);
                        attr.value(value, uri<>::microsoft::types_size);
                        return attr;
                    }
                };
            };

            inline void reparse(const rapidxml::xml_node<char>& source,
                                std::size_t size_hint)
            {
                rawdata_.reserve(size_hint);
                rapidxml::print(std::back_inserter(rawdata_),
                                source,
                                rapidxml::print_no_indenting);
                rawdata_.emplace_back('\0');

                // Note: we use xml_document::parse_ns here. This is because we
                // substitute the default namespace processor with a custom one.
                // Reason we do this: when we re-parse only a sub-tree of the
                // original XML document, we loose all enclosing namespace URIs.

                doc_->parse_ns<0, custom_namespace_processor>(&rawdata_[0]);
            }
        };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
        static_assert(std::is_default_constructible<item_properties>::value, "");
        static_assert(std::is_copy_constructible<item_properties>::value, "");
        static_assert(std::is_copy_assignable<item_properties>::value, "");
        static_assert(std::is_move_constructible<item_properties>::value, "");
        static_assert(std::is_move_assignable<item_properties>::value, "");
#endif

    }

    // Note About Dates in EWS
    //
    // Microsoft EWS uses date and date/time string representations as described
    // in http://www.w3.org/TR/xmlschema-2/, notably xs:dateTime (or
    // http://www.w3.org/2001/XMLSchema:dateTime) and xs:date (also known as
    // http://www.w3.org/2001/XMLSchema:date).
    //
    // For example, the lexical representation of xs:date is
    //
    //     '-'? yyyy '-' mm '-' dd zzzzzz?
    //
    // whereas the z represents the timezone. Two examples of date strings are:
    // 2000-01-16Z and 1981-07-02 (the Z means Zulu time which is the same as
    // UTC). xs:dateTime is formatted accordingly, just with a time component;
    // you get the idea.
    //
    // This library does not interpret, parse, or in any way touch date nor
    // date/time strings in any circumstance. This library provides two classes,
    // date and date_type. Both classes, date and date_time, act solely as thin
    // wrapper to make the signatures of public API functions more type-rich and
    // easier to understand. Both types are implicitly convertible from
    // std::string.
    //
    // If your date or date/time strings are not formatted properly, Microsoft
    // EWS will likely give you a SOAP fault which this library transports to
    // you as an exception of type ews::soap_fault.
    // TODO: we don't really need date class, date_time class is sufficient if
    // xs:date and xs:dateTime are interchangeable.

    // A date/time string wrapper class for xs:dateTime formatted strings.
    //
    // See Note About Dates in EWS above.
    class date_time final
    {
    public:
        date_time(std::string str) // intentionally not explicit
            : val_(std::move(str))
        {
        }
        const std::string& to_string() const EWS_NOEXCEPT { return val_; }

    private:
        friend bool operator==(const date_time&, const date_time&);
        std::string val_;
    };

    inline bool operator==(const date_time& lhs, const date_time& rhs)
    {
        return lhs.val_ == rhs.val_;
    }

    // A date string wrapper class for xs:date formatted strings.
    //
    // See Note About Dates in EWS above.
    class date final
    {
    public:
        date(std::string str) // intentionally not explicit
            : date_string_(std::move(str))
        {
        }

    private:
        std::string date_string_;
    };

    // Specifies the type of a <Body> element
    enum class body_type { best, plain_text, html };

    // TODO: move to internal namespace
    inline std::string body_type_str(body_type type)
    {
        switch (type)
        {
            case body_type::best: return "Best";
            case body_type::plain_text: return "Text";
            case body_type::html: return "HTML";
            default: throw exception("Bad enum value");
        }
    }

    // Represents the actual body content of a message.
    //
    // This can be of type Best, HTML, or plain-text. See EWS XML elements
    // documentation on MSDN.
    class body final
    {
    public:
        // Creates an empty body element; body_type is plain-text
        body()
            : content_(),
              type_(body_type::plain_text),
              is_truncated_(false)
        {
        }

        // Creates a new body element with given content and type
        explicit body(std::string content,
                      body_type type = body_type::plain_text)
            : content_(std::move(content)),
              type_(type),
              is_truncated_(false)
        {
        }

        body_type type() const EWS_NOEXCEPT { return type_; }

        void set_type(body_type type) { type_ = type; }

        bool is_truncated() const EWS_NOEXCEPT { return is_truncated_; }

        void set_truncated(bool truncated) { is_truncated_ = truncated; }

        const std::string& content() const EWS_NOEXCEPT { return content_; }

        void set_content(std::string content) { content_ = std::move(content); }

        std::string to_xml(const char* xmlns=nullptr) const
        {
            //FIXME: what about IsTruncated attribute?
            static const auto cdata_beg = std::string("<![CDATA[");
            static const auto cdata_end = std::string("]]>");

            auto pref = std::string();
            if (xmlns)
            {
                pref = std::string(xmlns) + ":";
            }
            std::stringstream sstr;
            sstr << "<" << pref << "Body BodyType=\"" << body_type_str(type());
            sstr << "\">";
            if (   type() == body_type::html
                && !(content_.compare(0, cdata_beg.length(), cdata_beg) == 0))
            {
                sstr << cdata_beg << content_ << cdata_end;
            }
            else
            {
                sstr << content_;
            }
            sstr << "</" << pref << "Body>";
            return sstr.str();
        }

    private:
        std::string content_;
        body_type type_;
        bool is_truncated_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<body>::value, "");
    static_assert(std::is_copy_constructible<body>::value, "");
    static_assert(std::is_copy_assignable<body>::value, "");
    static_assert(std::is_move_constructible<body>::value, "");
    static_assert(std::is_move_assignable<body>::value, "");
#endif

    // Represents an item's <MimeContent CharacterSet="" /> element.
    //
    // Contains the ASCII MIME stream of an object that is represented in
    // base64Binary format (as in RFC 2045).
    class mime_content final
    {
    public:
        mime_content() = default;

        // Copies len bytes from ptr into an internal buffer.
        mime_content(std::string charset,
                     const char* const ptr,
                     std::size_t len)
            : charset_(std::move(charset)),
              bytearray_(ptr, ptr + len)
        {}

        // Returns how the string is encoded, e.g., "UTF-8"
        const std::string& character_set() const EWS_NOEXCEPT
        {
            return charset_;
        }

        // Note: the pointer to the data is not 0-terminated
        const char* bytes() const EWS_NOEXCEPT { return bytearray_.data(); }

        std::size_t len_bytes() const EWS_NOEXCEPT { return bytearray_.size(); }

        // Returns true if no MIME content is available.  Note that a
        // <MimeContent> property is only included in a GetItem response when
        // explicitly requested using additional properties. This function lets
        // you test whether MIME content is available.
        bool none() const EWS_NOEXCEPT { return len_bytes() == 0U; }

    private:
        std::string charset_;
        std::vector<char> bytearray_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<mime_content>::value, "");
    static_assert(std::is_copy_constructible<mime_content>::value, "");
    static_assert(std::is_copy_assignable<mime_content>::value, "");
    static_assert(std::is_move_constructible<mime_content>::value, "");
    static_assert(std::is_move_assignable<mime_content>::value, "");
#endif

    // Represents a contact's email address
    // TODO: maybe rename to mailbox
    class email_address final
    {
    public:
        explicit email_address(item_id id)
            : id_(std::move(id)),
              value_(),
              name_(),
              routing_type_(),
              mailbox_type_()
        {}

        explicit email_address(std::string value,
                               std::string name = std::string(),
                               std::string routing_type = std::string(),
                               std::string mailbox_type = std::string())
            : id_(),
              value_(std::move(value)),
              name_(std::move(name)),
              routing_type_(std::move(routing_type)),
              mailbox_type_(std::move(mailbox_type))
        {}

        // TODO: rename
        const item_id& id() const EWS_NOEXCEPT { return id_; }

        // The address
        const std::string& value() const EWS_NOEXCEPT { return value_; }

        // Defines the name of the mailbox user; optional
        const std::string& name() const EWS_NOEXCEPT { return name_; }

        // Defines the routing that is used for the mailbox, attribute is
        // optional. Default is SMTP
        const std::string& routing_type() const EWS_NOEXCEPT
        {
            return routing_type_;
        }

        // Defines the mailbox type of a mailbox user; optional
        const std::string& mailbox_type() const EWS_NOEXCEPT
        {
            return mailbox_type_;
        }

        std::string to_xml(const char* xmlns=nullptr) const
        {
            auto pref = std::string();
            if (xmlns)
            {
                pref = std::string(xmlns) + ":";
            }
            std::stringstream sstr;
            sstr << "<" << pref << "Mailbox>";
            if (id().valid())
            {
                sstr << id().to_xml(xmlns);
            }
            else
            {
                sstr << "<" << pref << "EmailAddress>" << value()
                     <<"</" << pref << "EmailAddress>";

                if (!name().empty())
                {
                    sstr << "<" << pref << "Name>" << name()
                         << "</" << pref << "Name>";
                }

                if (!routing_type().empty())
                {
                    sstr << "<" << pref << "RoutingType>" << routing_type()
                         << "</" << pref << "RoutingType>";
                }

                if (!mailbox_type().empty())
                {
                    sstr << "<" << pref << "MailboxType>" << mailbox_type()
                         << "</" << pref << "MailboxType>";
                }
            }
            sstr << "</" << pref << "Mailbox>";
            return sstr.str();
        }

    private:
        item_id id_;
        std::string value_;
        std::string name_;
        std::string routing_type_;
        std::string mailbox_type_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(!std::is_default_constructible<email_address>::value, "");
    static_assert(std::is_copy_constructible<email_address>::value, "");
    static_assert(std::is_copy_assignable<email_address>::value, "");
    static_assert(std::is_move_constructible<email_address>::value, "");
    static_assert(std::is_move_assignable<email_address>::value, "");
#endif


    // Represents a generic item in the Exchange store.
    //
    // Basically:
    //
    //      item
    //      ├── appointment
    //      ├── contact
    //      ├── message
    //      └── task
    class item
    {
    public:
        item() = default;

        explicit item(item_id id) : item_id_(std::move(id)), properties_() {}

        item(item_id&& id, internal::item_properties&& properties)
            : item_id_(std::move(id)),
              properties_(std::move(properties))
        {}

        const item_id& get_item_id() const EWS_NOEXCEPT { return item_id_; }

        // Base64-encoded contents of the MIME stream for an item
        mime_content get_mime_content() const
        {
            const auto node = properties().get_node("MimeContent");
            if (!node)
            {
                return mime_content();
            }
            auto charset = node->first_attribute("CharacterSet");
            EWS_ASSERT(charset &&
                    "Expected <MimeContent> to have CharacterSet attribute");
            return mime_content(charset->value(),
                                node->value(),
                                node->value_size());
        }

        // Unique identifier for the folder that contains an item. This is a
        // read-only property
        // TODO: get_parent_folder_id

        // PR_MESSAGE_CLASS MAPI property (the message class) for an item
        // TODO: get_item_class

        // Sets this item's subject. Limited to 255 characters.
        void set_subject(const std::string& subject)
        {
            properties().set_or_update("Subject", subject);
        }

        // Returns this item's subject
        std::string get_subject() const
        {
            return properties().get_value_as_string("Subject");
        }

        // Enumeration indicating the sensitive nature of an item; valid
        // values are Normal, Personal, Private, and Confidential
        // TODO: get_sensitivity

        // Set the body content of an item
        void set_body(const body& b)
        {
            using uri = internal::uri<>;

            auto doc = properties().document();

            auto body_node = properties().get_node("Body");
            if (body_node)
            {
                doc->remove_node(body_node);
            }

            auto ptr_to_qname = doc->allocate_string("t:Body");
            body_node = doc->allocate_node(rapidxml::node_element);
            body_node->qname(ptr_to_qname,
                             std::strlen("t:Body"),
                             ptr_to_qname + 2);
            body_node->namespace_uri(uri::microsoft::types(),
                                     uri::microsoft::types_size);

            auto ptr_to_value =
                doc->allocate_string(b.content().c_str());
            body_node->value(ptr_to_value, b.content().length());

            auto ptr_to_key = doc->allocate_string("BodyType");
            ptr_to_value = doc->allocate_string(
                    body_type_str(b.type()).c_str());
            body_node->append_attribute(
                    doc->allocate_attribute(ptr_to_key, ptr_to_value));

            bool truncated = b.is_truncated();
            if (truncated)
            {
                ptr_to_key = doc->allocate_string("IsTruncated");
                ptr_to_value = doc->allocate_string(
                        truncated ? "true" : "false");
                body_node->append_attribute(
                        doc->allocate_attribute(ptr_to_key, ptr_to_value));
            }

            doc->append_node(body_node);
        }

        // Returns the body contents of an item
        body get_body() const
        {
            using rapidxml::internal::compare;

            body b;
            auto body_node = properties().get_node("Body");
            if (body_node)
            {
                for (auto attr = body_node->first_attribute(); attr;
                     attr = attr->next_attribute())
                {
                    if (compare(attr->name(), attr->name_size(),
                                "BodyType", std::strlen("BodyType")))
                    {
                        if (compare(attr->value(), attr->value_size(),
                                    "HTML", std::strlen("HTML")))
                        {
                            b.set_type(body_type::html);
                        }
                        else if (compare(attr->value(), attr->value_size(),
                                         "Text", std::strlen("Text")))
                        {
                            b.set_type(body_type::plain_text);
                        }
                        else if (compare(attr->value(), attr->value_size(),
                                         "Best", std::strlen("Best")))
                        {
                            b.set_type(body_type::best);
                        }
                        else
                        {
                            EWS_ASSERT(false
                                && "Unexpected attribute value for BodyType");
                        }
                    }
                    else if (compare(attr->name(), attr->name_size(),
                                     "IsTruncated", std::strlen("IsTruncated")))
                    {
                        const auto val =
                            compare(attr->value(), attr->value_size(),
                                    "true", std::strlen("true"));
                        b.set_truncated(val);
                    }
                    else
                    {
                        EWS_ASSERT(false
                                && "Unexpected attribute in <Body> element");
                    }
                }

                auto content = std::string(body_node->value(),
                                           body_node->value_size());
                b.set_content(std::move(content));
            }
            return b;
        }

        // Metadata about the attachments of an item
        // TODO: get_attachments

        // Date/time an item was received
        // TODO: get_date_time_received

        // Size in bytes of an item. This is a read-only property
        // TODO: get_size

        // Categories associated with an item
        // TODO: get_categories

        // Enumeration indicating the importance of an item; valid values
        // are Low, Normal, and High
        // TODO: get_importance

        // Taken from PR_IN_REPLY_TO_ID MAPI property
        // TODO: get_in_reply_to

        // True if an item has been submitted for delivery
        // TODO: get_is_submitted

        // True if an item is a draft
        // TODO: is_draft

        // True if an item is from you
        // TODO: is_from_me

        // True if an item a re-send
        // TODO: is_resend

        // True if an item is unmodified
        // TODO: is_unmodified

        // Collection of Internet message headers associated with an item
        // TODO: get_internet_message_headers

        // Date/time an item was sent
        // TODO: get_date_time_sent

        // Date/time an item was created
        // TODO: get_date_time_created

        // Applicable actions for an item
        // (NonEmptyArrayOfResponseObjectsType)
        // TODO: get_response_objects

        // Set due date of an item; used for reminders
        void set_reminder_due_by(const date_time& due_by)
        {
            properties().set_or_update("ReminderDueBy", due_by.to_string());
        }

        // Returns the due date of an item; used for reminders
        date_time get_reminder_due_by() const
        {
            return date_time(properties().get_value_as_string("ReminderDueBy"));
        }

        // Set a reminder on an item
        void set_reminder_enabled(bool enabled)
        {
            properties().set_or_update("ReminderIsSet",
                                       enabled ? "true" : "false");
        }

        // True if a reminder has been set on an item
        bool is_reminder_enabled() const
        {
            return properties().get_value_as_string("ReminderIsSet") == "true";
        }

        // Number of minutes before the due date that a reminder should be
        // shown to the user
        // TODO: get_reminder_minutes_before_start

        // Concatenated string of the display names of the Cc recipients of
        // an item; each recipient is separated by a semicolon
        // TODO: get_display_cc

        // Concatenated string of the display names of the To recipients of
        // an item; each recipient is separated by a semicolon
        // TODO: get_display_to

        // True if an item has non-hidden attachments. This is a read-only
        // property
        bool has_attachments() const
        {
            return properties().get_value_as_string("HasAttachments") == "true";
        }

        // List of zero or more extended properties that are requested for
        // an item
        // TODO: get_extended_property

        // Culture name associated with the body of an item
        // TODO: get_culture

        // Following properties are beyond 2007 scope:
        //   <EffectiveRights/>
        //   <LastModifiedName/>
        //   <LastModifiedTime/>
        //   <IsAssociated/>
        //   <WebClientReadFormQueryString/>
        //   <WebClientEditFormQueryString/>
        //   <ConversationId/>
        //   <UniqueBody/>

    protected:
        internal::item_properties& properties() EWS_NOEXCEPT
        {
            return properties_;
        }

        const internal::item_properties& properties() const EWS_NOEXCEPT
        {
            return properties_;
        }

    private:
        item_id item_id_;
        internal::item_properties properties_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<item>::value, "");
    static_assert(std::is_copy_constructible<item>::value, "");
    static_assert(std::is_copy_assignable<item>::value, "");
    static_assert(std::is_move_constructible<item>::value, "");
    static_assert(std::is_move_assignable<item>::value, "");
#endif

    // Represents a concrete task in the Exchange store.
    class task final : public item
    {
    public:
        task() = default;

        explicit task(item_id id) : item(std::move(id)) {}

        task(item_id&& id, internal::item_properties&& properties)
            : item(std::move(id), std::move(properties))
        {}

        // Represents the actual amount of work expended on the task. Measured
        // in minutes
        // TODO: get_actual_work

        // Time the task was assigned to the current owner
        // TODO: get_assigned_time

        // Billing information associated with this task
        // TODO: get_billing_information

        // How many times this task has been acted upon (sent, accepted,
        // etc.). This is simply a way to resolve conflicts when the
        // delegator sends multiple updates. Also known as TaskVersion
        // TODO: get_change_count

        // A list of company names associated with this task
        // TODO: get_companies

        // Time the task was completed
        // TODO: get_complete_date

        // Contact names associated with this task
        // TODO: get_contacts

        // Enumeration value indicating whether the delegated task was
        // accepted or not
        // TODO: get_delegation_state

        // Display name of the user that delegated the task
        // TODO: get_delegator

        // Sets the date that the task is due
        void set_due_date(const date_time& due_date)
        {
            properties().set_or_update("DueDate", due_date.to_string());
        }

        // Returns the date that the task is due
        date_time get_due_date() const
        {
            return date_time(properties().get_value_as_string("DueDate"));
        }

        // TODO: is_assignment_editable, possible values 0-5, 2007 dialect?

        // True if the task is marked as complete. This is a read-only property
        // See also set_percent_complete
        bool is_complete() const
        {
            return properties().get_value_as_string("IsComplete") == "true";
        }

        // True if the task is recurring
        // TODO: is_recurring

        // True if the task is a team task
        // TODO: is_team_task

        // Mileage associated with the task, potentially used for reimbursement
        // purposes
        // TODO: get_mileage

        // The name of the user who owns the task. This is a read-only property
        // TODO: Not in AllProperties shape in EWS 2013, investigate
#if 0
        std::string get_owner() const
        {
            return properties().get_value_as_string("Owner");
        }
#endif

        // The percentage of the task that has been completed. Valid values are
        // 0-100
        // TODO: get_percent_complete

        // Used for recurring tasks
        // TODO: get_recurrence

        // Set the date that work on the task should start
        void set_start_date(const date_time& start_date)
        {
            properties().set_or_update("StartDate", start_date.to_string());
        }

        // Returns the date that work on the task should start
        date_time get_start_date() const
        {
            return date_time(properties().get_value_as_string("StartDate"));
        }

        // The status of the task
        // TODO: get_status

        // A localized string version of the status. Useful for display
        // purposes
        // TODO: get_status_description

        // The total amount of work for this task
        // TODO: get_total_work

        // Every property below is 2012 or 2013 dialect

        // TODO: add remaining properties

        // Makes a task instance from a <Task> XML element
        static task from_xml_element(const rapidxml::xml_node<>& elem)
        {
            using uri = internal::uri<>;
            auto id_node = elem.first_node_ns(uri::microsoft::types(),
                                              "ItemId");
            EWS_ASSERT(id_node && "Expected <ItemId>");
            return task(item_id::from_xml_element(*id_node),
                        internal::item_properties(elem));
        }

    private:
        template <typename U> friend class basic_service;
        std::string create_item_request_string() const
        {
            std::stringstream sstr;
            sstr <<
              "<CreateItem " \
                  "xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " \
                  "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" >" \
              "<Items>" \
              "<t:Task>";
            sstr << properties().to_string() << "\n";
            sstr << "</t:Task>" \
              "</Items>" \
              "</CreateItem>";
            return sstr.str();
        }
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<task>::value, "");
    static_assert(std::is_copy_constructible<task>::value, "");
    static_assert(std::is_copy_assignable<task>::value, "");
    static_assert(std::is_move_constructible<task>::value, "");
    static_assert(std::is_move_assignable<task>::value, "");
#endif

    // A contact item in the Exchange store.
    class contact final : public item
    {
    public:
        contact() = default;

        explicit contact(item_id id) : item(id) {}

        contact(item_id&& id, internal::item_properties&& properties)
            : item(std::move(id), std::move(properties))
        {}

        // How the name should be filed for display/sorting purposes
        // TODO: file_as

        // How the various parts of a contact's information interact to form
        // the FileAs property value
        // TODO: file_as_mapping

        // The name to display for a contact
        // TODO: get_display_name

        // Sets the name by which a person is known to `given_name`; often
        // referred to as a person's first name
        void set_given_name(const std::string& given_name)
        {
            properties().set_or_update("GivenName", given_name);
        }

        // Returns the person's first name
        std::string get_given_name() const
        {
            return properties().get_value_as_string("GivenName");
        }

        // Initials for the contact
        // TODO: get_initials

        // The middle name for the contact
        // TODO: get_middle_name

        // Another name by which the contact is known
        // TODO: get_nickname

        // A combination of several name fields in one convenient place
        // (read-only)
        // TODO: get_complete_name

        // The company that the contact is affiliated with
        // TODO: get_company_name

        // A collection of e-mail addresses for the contact
        std::vector<email_address> get_email_addresses() const
        {
            const auto addresses = properties().get_node("EmailAddresses");
            if (!addresses)
            {
                return std::vector<email_address>();
            }
            std::vector<email_address> result;
            for (auto entry = addresses->first_node(); entry;
                 entry = entry->next_sibling())
            {
                auto name_attr = entry->first_attribute("Name");
                auto routing_type_attr = entry->first_attribute("RoutingType");
                auto mailbox_type_attr = entry->first_attribute("MailboxType");
                result.emplace_back(
                    email_address(
                        entry->value(),
                        name_attr ? name_attr->value() : "",
                        routing_type_attr ? routing_type_attr->value() : "",
                        mailbox_type_attr ? mailbox_type_attr->value() : ""));
            }
            return result;
        }

        std::string get_email_address_1() const
        {
            return get_email_address_by_key("EmailAddress1");
        }

        void set_email_address_1(email_address address)
        {
            set_email_address_by_key("EmailAddress1", std::move(address));
        }

        std::string get_email_address_2() const
        {
            return get_email_address_by_key("EmailAddress2");
        }

        void set_email_address_2(email_address address)
        {
            set_email_address_by_key("EmailAddress2", std::move(address));
        }

        std::string get_email_address_3() const
        {
            return get_email_address_by_key("EmailAddress3");
        }

        void set_email_address_3(email_address address)
        {
            set_email_address_by_key("EmailAddress3", std::move(address));
        }

        // A collection of mailing addresses for the contact
        // TODO: get_physical_addresses

        // A collection of phone numbers for the contact
        // TODO: get_phone_numbers

        // The name of the contact's assistant
        // TODO: get_assistant_name

        // The contact's birthday
        // TODO: get_birthday

        // Web page for the contact's business; typically a URL
        // TODO: get_business_homepage

        // A collection of children's names associated with the contact
        // TODO: get_children

        // A collection of companies a contact is associated with
        // TODO: get_companies

        // Indicates whether this is a directory or a store contact
        // (read-only)
        // TODO: get_contact_source

        // The department name that the contact is in
        // TODO: get_department

        // Sr, Jr, I, II, III, and so on
        // TODO: get_generation

        // A collection of instant messaging addresses for the contact
        // TODO: get_im_addresses

        // Sets this contact's job title.
        void set_job_title(const std::string& title)
        {
            properties().set_or_update("JobTitle", title);
        }

        // Returns the job title for the contact
        std::string get_job_title() const
        {
            return properties().get_value_as_string("JobTitle");
        }

        // The name of the contact's manager
        // TODO: get_manager

        // The distance that the contact resides from some reference point
        // TODO: get_mileage

        // Location of the contact's office
        // TODO: get_office_location

        // The physical addresses in the PhysicalAddresses collection that
        // represents the mailing address for the contact
        // TODO: get_postal_address_index

        // Occupation or discipline of the contact
        // TODO: get_profession

        // Set name of the contact's significant other
        void set_spouse_name(const std::string& spouse_name)
        {
            properties().set_or_update("SpouseName", spouse_name);
        }

        // Get name of the contact's significant other
        std::string get_spouse_name() const
        {
            return properties().get_value_as_string("SpouseName");
        }

        // Sets the family name of the contact; usually considered the last name
        void set_surname(const std::string& surname)
        {
            properties().set_or_update("Surname", surname);
        }

        // Returns the family name of the contact; usually considered the last
        // name
        std::string get_surname() const
        {
            return properties().get_value_as_string("Surname");
        }

        // Date that the contact was married
        // TODO: get_wedding_anniversary

        // Everything below is beyond EWS 2007 subset

        // has_picture
        // phonetic_full_name
        // phonetic_first_name
        // phonetic_last_name
        // alias
        // notes
        // photo
        // user_smime_certificate
        // msexchange_certificate
        // directory_id
        // manager_mailbox
        // direct_reports

        // Makes a contact instance from a <Contact> XML element
        static contact from_xml_element(const rapidxml::xml_node<>& elem)
        {
            using uri = internal::uri<>;
            auto id_node = elem.first_node_ns(uri::microsoft::types(),
                                              "ItemId");
            EWS_ASSERT(id_node && "Expected <ItemId>");
            return contact(item_id::from_xml_element(*id_node),
                           internal::item_properties(elem));
        }

    private:
        template <typename U> friend class basic_service;
        std::string create_item_request_string() const
        {
            std::stringstream sstr;
            sstr <<
              "<CreateItem " \
                  "xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " \
                  "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" >" \
              "<Items>" \
              "<t:Contact>";
            sstr << properties().to_string() << "\n";
            sstr << "</t:Contact>" \
              "</Items>" \
              "</CreateItem>";
            return sstr.str();
        }

        // Helper function for get_email_address_{1,2,3}
        std::string get_email_address_by_key(const char* key) const
        {
            using rapidxml::internal::compare;

            // <Entry Key="" Name="" RoutingType="" MailboxType="" />
            const auto addresses = properties().get_node("EmailAddresses");
            if (!addresses)
            {
                return "";
            }
            for (auto entry = addresses->first_node(); entry;
                 entry = entry->next_sibling())
            {
                for (auto attr = entry->first_attribute(); attr;
                     attr = attr->next_attribute())
                {
                    if (   compare(attr->name(), attr->name_size(),
                                   "Key", 3)
                        && compare(attr->value(), attr->value_size(),
                                   key, std::strlen(key)))
                    {
                        return std::string(entry->value(), entry->value_size());
                    }
                }
            }
            // None with such key
            return "";
        }

        // Helper function for set_email_address_{1,2,3}
        void set_email_address_by_key(const char* key, email_address&& mail)
        {
            using uri = internal::uri<>;
            using rapidxml::internal::compare;

            auto doc = properties().document();
            auto addresses = properties().get_node("EmailAddresses");
            if (addresses)
            {
                // Check if there is alread any entry for given key

                bool exists = false;
                auto entry = addresses->first_node();
                for (; entry && !exists; entry = entry->next_sibling())
                {
                    for (auto attr = entry->first_attribute(); attr;
                         attr = attr->next_attribute())
                    {
                        if (   compare(attr->name(), attr->name_size(),
                                       "Key", 3)
                            && compare(attr->value(), attr->value_size(),
                                       key, std::strlen(key)))
                        {
                            exists = true;
                            break;
                        }
                    }
                }
                if (exists)
                {
                    addresses->remove_node(entry);
                }
            }
            else
            {
                // Need to construct <EmailAddresses> node first

                auto ptr_to_qname = doc->allocate_string("t:EmailAddresses");
                addresses = doc->allocate_node(rapidxml::node_element);
                addresses->qname(ptr_to_qname,
                                 std::strlen("t:EmailAddresses"),
                                 ptr_to_qname + 2);
                addresses->namespace_uri(uri::microsoft::types(),
                                         uri::microsoft::types_size);
                doc->append_node(addresses);
            }

            // <Entry Key="" Name="" RoutingType="" MailboxType="" />
            auto new_entry_qname = doc->allocate_string("t:Entry");
            auto new_entry_value = doc->allocate_string(mail.value().c_str());
            auto new_entry = doc->allocate_node(rapidxml::node_element);
            new_entry->qname(new_entry_qname,
                             std::strlen("t:Entry"),
                             new_entry_qname + 2);
            new_entry->namespace_uri(uri::microsoft::types(),
                                     uri::microsoft::types_size);
            new_entry->value(new_entry_value);
            auto ptr_to_key = doc->allocate_string("Key");
            auto ptr_to_value = doc->allocate_string(key);
            new_entry->append_attribute(doc->allocate_attribute(ptr_to_key,
                                                                ptr_to_value));
            if (!mail.name().empty())
            {
                ptr_to_key = doc->allocate_string("Name");
                ptr_to_value = doc->allocate_string(mail.name().c_str());
                new_entry->append_attribute(
                        doc->allocate_attribute(ptr_to_key, ptr_to_value));
            }
            if (!mail.routing_type().empty())
            {
                ptr_to_key = doc->allocate_string("RoutingType");
                ptr_to_value = doc->allocate_string(
                        mail.routing_type().c_str());
                new_entry->append_attribute(
                        doc->allocate_attribute(ptr_to_key, ptr_to_value));
            }
            if (!mail.mailbox_type().empty())
            {
                ptr_to_key = doc->allocate_string("MailboxType");
                ptr_to_value = doc->allocate_string(
                        mail.mailbox_type().c_str());
                new_entry->append_attribute(
                        doc->allocate_attribute(ptr_to_key, ptr_to_value));
            }
            addresses->append_node(new_entry);
        }
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<contact>::value, "");
    static_assert(std::is_copy_constructible<contact>::value, "");
    static_assert(std::is_copy_assignable<contact>::value, "");
    static_assert(std::is_move_constructible<contact>::value, "");
    static_assert(std::is_move_assignable<contact>::value, "");
#endif

    // A message item in the Exchange store.
    class message final : public item
    {
    public:
        message() = default;

        explicit message(item_id id) : item(id) {}

        message(item_id&& id, internal::item_properties&& properties)
            : item(std::move(id), std::move(properties))
        {}

        // <Sender/>

        void set_to_recipients(const std::vector<email_address>& recipients)
        {
            using uri = internal::uri<>;

            auto doc = properties().document();

            auto to_recipients_node = properties().get_node("ToRecipients");
            if (to_recipients_node)
            {
                doc->remove_node(to_recipients_node);
            }

            auto ptr_to_qname = doc->allocate_string("t:ToRecipients");
            to_recipients_node = doc->allocate_node(rapidxml::node_element);
            to_recipients_node->qname(ptr_to_qname,
                                      std::strlen("t:ToRecipients"),
                                      ptr_to_qname + 2);
            to_recipients_node->namespace_uri(uri::microsoft::types(),
                                              uri::microsoft::types_size);

            for (const auto& recipient : recipients)
            {
                ptr_to_qname = doc->allocate_string("t:Mailbox");
                auto mailbox_node = doc->allocate_node(rapidxml::node_element);
                mailbox_node->qname(ptr_to_qname,
                                    std::strlen("t:Mailbox"),
                                    ptr_to_qname + 2);
                mailbox_node->namespace_uri(uri::microsoft::types(),
                                            uri::microsoft::types_size);

                if (!recipient.id().valid())
                {
                    EWS_ASSERT(!recipient.value().empty()
                  && "Neither item_id nor value set in email_address instance");

                    ptr_to_qname = doc->allocate_string("t:EmailAddress");
                    auto ptr_to_value = doc->allocate_string(
                            recipient.value().c_str());
                    auto node = doc->allocate_node(rapidxml::node_element);
                    node->qname(ptr_to_qname,
                                std::strlen("t:EmailAddress"),
                                ptr_to_qname + 2);
                    node->value(ptr_to_value);
                    node->namespace_uri(uri::microsoft::types(),
                                        uri::microsoft::types_size);
                    mailbox_node->append_node(node);

                    if (!recipient.name().empty())
                    {
                        ptr_to_qname = doc->allocate_string("t:Name");
                        ptr_to_value = doc->allocate_string(
                                recipient.name().c_str());
                        node = doc->allocate_node(rapidxml::node_element);
                        node->qname(ptr_to_qname,
                                    std::strlen("t:Name"),
                                    ptr_to_qname + 2);
                        node->value(ptr_to_value);
                        node->namespace_uri(uri::microsoft::types(),
                                            uri::microsoft::types_size);
                        mailbox_node->append_node(node);
                    }

                    if (!recipient.routing_type().empty())
                    {
                        ptr_to_qname = doc->allocate_string("t:RoutingType");
                        ptr_to_value = doc->allocate_string(
                                recipient.routing_type().c_str());
                        node = doc->allocate_node(rapidxml::node_element);
                        node->qname(ptr_to_qname,
                                    std::strlen("t:RoutingType"),
                                    ptr_to_qname + 2);
                        node->value(ptr_to_value);
                        node->namespace_uri(uri::microsoft::types(),
                                            uri::microsoft::types_size);
                        mailbox_node->append_node(node);
                    }

                    if (!recipient.mailbox_type().empty())
                    {
                        ptr_to_qname = doc->allocate_string("t:MailboxType");
                        ptr_to_value = doc->allocate_string(
                                recipient.mailbox_type().c_str());
                        node = doc->allocate_node(rapidxml::node_element);
                        node->qname(ptr_to_qname,
                                    std::strlen("t:MailboxType"),
                                    ptr_to_qname + 2);
                        node->value(ptr_to_value);
                        node->namespace_uri(uri::microsoft::types(),
                                            uri::microsoft::types_size);
                        mailbox_node->append_node(node);
                    }
                }
                else
                {
                    ptr_to_qname = doc->allocate_string("t:ItemId");
                    auto item_id_node =
                        doc->allocate_node(rapidxml::node_element);
                    item_id_node->qname(ptr_to_qname,
                                        std::strlen("t:ItemId"),
                                        ptr_to_qname + 2);
                    item_id_node->namespace_uri(uri::microsoft::types(),
                                                uri::microsoft::types_size);

                    auto ptr_to_key = doc->allocate_string("Id");
                    auto ptr_to_value = doc->allocate_string(
                            recipient.id().id().c_str()); // Uh, thats smelly
                    item_id_node->append_attribute(
                            doc->allocate_attribute(ptr_to_key, ptr_to_value));

                    ptr_to_key = doc->allocate_string("ChangeKey");
                    ptr_to_value = doc->allocate_string(
                            recipient.id().change_key().c_str());
                    item_id_node->append_attribute(
                            doc->allocate_attribute(ptr_to_key, ptr_to_value));

                    mailbox_node->append_node(item_id_node);
                }

                to_recipients_node->append_node(mailbox_node);
            }

            doc->append_node(to_recipients_node);
        }

        std::vector<email_address> get_to_recipients()
        {
            using rapidxml::internal::compare;

            const auto recipients = properties().get_node("ToRecipients");
            if (!recipients)
            {
                return std::vector<email_address>();
            }
            std::vector<email_address> result;
            for (auto mailbox = recipients->first_node(); mailbox;
                 mailbox = mailbox->next_sibling())
            {
                // <EmailAddress> child element is required except when dealing
                // with a private distribution list or a contact from a user's
                // contacts folder, in which case the <ItemId> child element is
                // used instead

                std::string name;
                std::string address;
                std::string routing_type;
                std::string mailbox_type;
                item_id id;

                for (auto node = mailbox->first_node(); node;
                     node = node->next_sibling())
                {
                    if (compare(node->local_name(),
                                node->local_name_size(),
                                "Name",
                                std::strlen("Name")))
                    {
                        name = std::string(node->value(), node->value_size());
                    }
                    else if (compare(node->local_name(),
                                     node->local_name_size(),
                                     "EmailAddress",
                                     std::strlen("EmailAddress")))
                    {
                        address =
                            std::string(node->value(), node->value_size());
                    }
                    else if (compare(node->local_name(),
                                     node->local_name_size(),
                                     "RoutingType",
                                     std::strlen("RoutingType")))
                    {
                        routing_type =
                            std::string(node->value(), node->value_size());
                    }
                    else if (compare(node->local_name(),
                                     node->local_name_size(),
                                     "MailboxType",
                                     std::strlen("MailboxType")))
                    {
                        mailbox_type =
                            std::string(node->value(), node->value_size());
                    }
                    else if (compare(node->local_name(),
                                     node->local_name_size(),
                                     "ItemId",
                                     std::strlen("ItemId")))
                    {
                        id = item_id::from_xml_element(*node);
                    }
                    else
                    {
                        throw exception(
                                "Unexpected child element in <Mailbox>");
                    }
                }

                if (!id.valid())
                {
                    EWS_ASSERT(!address.empty()
                            && "<EmailAddress> element value can't be empty");

                    result.emplace_back(
                        email_address(
                            std::move(address),
                            std::move(name),
                            std::move(routing_type),
                            std::move(mailbox_type)));
                }
                else
                {
                    result.emplace_back(
                        email_address(
                            std::move(id)));
                }
            }
            return result;
        }

        // <CcRecipients/>
        // <BccRecipients/>
        // <IsReadReceiptRequested/>
        // <IsDeliveryReceiptRequested/>
        // <ConversationIndex/>
        // <ConversationTopic/>
        // <From/>
        // <InternetMessageId/>
        // <IsRead/>
        // <IsResponseRequested/>
        // <References/>
        // <ReplyTo/>

        // Makes a message instance from a <Message> XML element
        static message from_xml_element(const rapidxml::xml_node<>& elem)
        {
            using uri = internal::uri<>;
            auto id_node = elem.first_node_ns(uri::microsoft::types(),
                                              "ItemId");
            EWS_ASSERT(id_node && "Expected <ItemId>");
            return message(item_id::from_xml_element(*id_node),
                           internal::item_properties(elem));
        }

    private:
        template <typename U> friend class basic_service;
        std::string
        create_item_request_string(ews::message_disposition disposition) const
        {
            std::stringstream sstr;
            sstr << "<m:CreateItem MessageDisposition=\""
                 << message_disposition_str(disposition) << "\">";
            sstr << "<m:Items>";
            sstr << "<t:Message>";
            sstr << properties().to_string() << "\n";
            sstr << "</t:Message>";
            sstr << "</m:Items>";
            sstr << "</m:CreateItem>";
            return sstr.str();
        }
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<message>::value, "");
    static_assert(std::is_copy_constructible<message>::value, "");
    static_assert(std::is_copy_assignable<message>::value, "");
    static_assert(std::is_move_constructible<message>::value, "");
    static_assert(std::is_move_assignable<message>::value, "");
#endif

    // Well known folder names enumeration. Usually rendered to XML as
    // <DistinguishedFolderId> element.
    enum class standard_folder
    {
        // The Calendar folder
        calendar,

        // The Contacts folder
        contacts,

        // The Deleted Items folder
        deleted_items,

        // The Drafts folder
        drafts,

        // The Inbox folder
        inbox,

        // The Journal folder
        journal,

        // The Notes folder
        notes,

        // The Outbox folder
        outbox,

        // The Sent Items folder
        sent_items,

        // The Tasks folder
        tasks,

        // The root of the message folder hierarchy
        msg_folder_root,

        // The root of the mailbox
        root,

        // The Junk E-mail folder
        junk_email,

        // The Search Folders folder, also known as the Finder folder
        search_folders,

        // The Voicemail folder
        voice_mail,

        // Following are folders containing recoverable items:

        // The root of the Recoverable Items folder hierarchy
        recoverable_items_root,

        // The root of the folder hierarchy of recoverable items that have been
        // soft-deleted from the Deleted Items folder
        recoverable_items_deletions,

        // The root of the Recoverable Items versions folder hierarchy in the
        // archive mailbox
        recoverable_items_versions,

        // The root of the folder hierarchy of recoverable items that have been
        // hard-deleted from the Deleted Items folder
        recoverable_items_purges,

        // The root of the folder hierarchy in the archive mailbox
        archive_root,

        // The root of the message folder hierarchy in the archive mailbox
        archive_msg_folder_root,

        // The Deleted Items folder in the archive mailbox
        archive_deleted_items,

        // Represents the archive Inbox folder. Caution: only versions of
        // Exchange starting with build number 15.00.0913.09 include this folder
        archive_inbox,

        // The root of the Recoverable Items folder hierarchy in the archive
        // mailbox
        archive_recoverable_items_root,

        // The root of the folder hierarchy of recoverable items that have been
        // soft-deleted from the Deleted Items folder of the archive mailbox
        archive_recoverable_items_deletions,

        // The root of the Recoverable Items versions folder hierarchy in the
        // archive mailbox
        archive_recoverable_items_versions,

        // The root of the hierarchy of recoverable items that have been
        // hard-deleted from the Deleted Items folder of the archive mailbox
        archive_recoverable_items_purges,

        // Following are folders that came with EWS 2013 and Exchange Online:

        // The Sync Issues folder
        sync_issues,

        // The Conflicts folder
        conflicts,

        // The Local Failures folder
        local_failures,

        // Represents the Server Failures folder
        server_failures,

        // The recipient cache folder
        recipient_cache,

        // The quick contacts folder
        quick_contacts,

        // The conversation history folder
        conversation_history,

        // Represents the admin audit logs folder
        admin_audit_logs,

        // The todo search folder
        todo_search,

        // Represents the My Contacts folder
        my_contacts,

        // Represents the directory folder
        directory,

        // Represents the IM contact list folder
        im_contact_list,

        // Represents the people connect folder
        people_connect,

        // Represents the Favorites folder
        favorites,
    };

    // Identifies a foler.
    //
    // Renders a <FolderId> element. Contains the identifier and change key of a
    // folder.
    class folder_id
    {
    public:
        folder_id() = delete;

        ~folder_id() = default;

        std::string to_xml(const char* xmlns=nullptr) const
        {
            return func_(xmlns);
        }

    protected:
        explicit folder_id(std::function<std::string (const char*)>&& func)
            : func_(std::move(func))
        {
        }

    private:
        std::function<std::string (const char*)> func_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(!std::is_default_constructible<folder_id>::value, "");
    static_assert(std::is_copy_constructible<folder_id>::value, "");
    static_assert(std::is_copy_assignable<folder_id>::value, "");
    static_assert(std::is_move_constructible<folder_id>::value, "");
    static_assert(std::is_move_assignable<folder_id>::value, "");
#endif

    // Renders a <DistinguishedFolderId> element. Implicitly convertible from
    // standard_folder.
    class distinguished_folder_id final : public folder_id
    {
    public:
        distinguished_folder_id() = delete;

        // Intentionally not explicit
        distinguished_folder_id(standard_folder folder)
            : folder_id([=](const char* xmlns) -> std::string
                    {
                        auto pref = std::string();
                        if (xmlns)
                        {
                            pref = std::string(xmlns) + ":";
                        }
                        std::stringstream sstr;
                        sstr << "<" << pref << "DistinguishedFolderId Id=\"";
                        sstr << well_known_name(folder);
                        sstr << "\" />";
                        return sstr.str();
                    })
        {
        }

#if 0
        // TODO: Constructor for EWS delegate access
        distinguished_folder_id(standard_folder, mailbox) {}
#endif

    private:
        static std::string well_known_name(standard_folder enumeration)
        {
            std::string name;
            switch (enumeration)
            {
                case standard_folder::calendar: name = "calendar"; break;
                case standard_folder::contacts: name = "contacts"; break;
                case standard_folder::deleted_items: name = "deleteditems"; break;
                case standard_folder::drafts: name = "drafts"; break;
                case standard_folder::inbox: name = "inbox"; break;
                case standard_folder::journal: name = "journal"; break;
                case standard_folder::notes: name = "notes"; break;
                case standard_folder::outbox: name = "outbox"; break;
                case standard_folder::sent_items: name = "sentitems"; break;
                case standard_folder::tasks: name = "tasks"; break;
                case standard_folder::msg_folder_root: name = "msgfolderroot"; break;
                case standard_folder::root: name = "root"; break;
                case standard_folder::junk_email: name = "junkemail"; break;
                case standard_folder::search_folders: name = "searchfolders"; break;
                case standard_folder::voice_mail: name = "voicemail"; break;
                case standard_folder::recoverable_items_root: name = "recoverableitemsroot"; break;
                case standard_folder::recoverable_items_deletions: name = "recoverableitemsdeletions"; break;
                case standard_folder::recoverable_items_versions: name = "recoverableitemsversions"; break;
                case standard_folder::recoverable_items_purges: name = "recoverableitemspurges"; break;
                case standard_folder::archive_root: name = "archiveroot"; break;
                case standard_folder::archive_msg_folder_root: name = "archivemsgfolderroot"; break;
                case standard_folder::archive_deleted_items: name = "archivedeleteditems"; break;
                case standard_folder::archive_inbox: name = "archiveinbox"; break;
                case standard_folder::archive_recoverable_items_root: name = "archiverecoverableitemsroot"; break;
                case standard_folder::archive_recoverable_items_deletions: name = "archiverecoverableitemsdeletions"; break;
                case standard_folder::archive_recoverable_items_versions: name = "archiverecoverableitemsversions"; break;
                case standard_folder::archive_recoverable_items_purges: name = "archiverecoverableitemspurges"; break;
                case standard_folder::sync_issues: name = "syncissues"; break;
                case standard_folder::conflicts: name = "conflicts"; break;
                case standard_folder::local_failures: name = "localfailures"; break;
                case standard_folder::server_failures: name = "serverfailures"; break;
                case standard_folder::recipient_cache: name = "recipientcache"; break;
                case standard_folder::quick_contacts: name = "quickcontacts"; break;
                case standard_folder::conversation_history: name = "conversationhistory"; break;
                case standard_folder::admin_audit_logs: name = "adminauditlogs"; break;
                case standard_folder::todo_search: name = "todosearch"; break;
                case standard_folder::my_contacts: name = "mycontacts"; break;
                case standard_folder::directory: name = "directory"; break;
                case standard_folder::im_contact_list: name = "imcontactlist"; break;
                case standard_folder::people_connect: name = "peopleconnect"; break;
                case standard_folder::favorites: name = "favorites"; break;
                default:
                    throw exception("Unrecognized folder name");
            };
            return name;
        }
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(!std::is_default_constructible<distinguished_folder_id>::value, "");
    static_assert(std::is_copy_constructible<distinguished_folder_id>::value, "");
    static_assert(std::is_copy_assignable<distinguished_folder_id>::value, "");
    static_assert(std::is_move_constructible<distinguished_folder_id>::value, "");
    static_assert(std::is_move_assignable<distinguished_folder_id>::value, "");
#endif

    class property_path
    {
    public:
        // Intentionally not explicit
        property_path(const char* uri) : uri_(std::string(uri)) {}

        // Returns the <FieldURI> element for this property. Identifies
        // frequently referenced properties by URI
        const std::string& field_uri() const EWS_NOEXCEPT { return uri_; }

        std::string property_name() const
        {
            const auto n = uri_.rfind(':');
            EWS_ASSERT(n != std::string::npos);
            return uri_.substr(n + 1);
        }

        std::string class_name() const
        {
            // TODO: we know at compile-time to which class a property belongs
            const auto n = uri_.find(':');
            EWS_ASSERT(n != std::string::npos);
            const auto substr = uri_.substr(0, n);
            if (substr == "folder")
            {
                return "Folder";
            }
            else if (substr == "item")
            {
                return "Item";
            }
            else if (substr == "message")
            {
                return "Message";
            }
            else if (substr == "meeting")
            {
                return "Meeting";
            }
            else if (substr == "meetingRequest")
            {
                return "MeetingRequest";
            }
            else if (substr == "calendar")
            {
                return "Calendar";
            }
            else if (substr == "task")
            {
                return "Task";
            }
            else if (substr == "contacts")
            {
                return "Contact";
            }
            else if (substr == "distributionlist")
            {
                return "DistributionList";
            }
            else if (substr == "postitem")
            {
                return "PostItem";
            }
            else if (substr == "conversation")
            {
                return "Conversation";
            }
            // Persona missing
            // else if (substr == "")
            // {
            //     return "";
            // }
            throw exception("Unknown property path");
        }

    private:
        std::string uri_;
    };

    inline bool operator==(const property_path& lhs, const property_path& rhs)
    {
        return lhs.field_uri() == rhs.field_uri();
    }

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(!std::is_default_constructible<property_path>::value, "");
    static_assert(std::is_copy_constructible<property_path>::value, "");
    static_assert(std::is_copy_assignable<property_path>::value, "");
    static_assert(std::is_move_constructible<property_path>::value, "");
    static_assert(std::is_move_assignable<property_path>::value, "");
#endif

    class indexed_property_path : public property_path
    {
    public:
        indexed_property_path(const char* uri, const char* index)
            : property_path(uri), index_(std::string(index))
        {}

        const std::string& field_index() const EWS_NOEXCEPT { return index_; }

    private:
        std::string index_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(!std::is_default_constructible<indexed_property_path>::value, "");
    static_assert(std::is_copy_constructible<indexed_property_path>::value, "");
    static_assert(std::is_copy_assignable<indexed_property_path>::value, "");
    static_assert(std::is_move_constructible<indexed_property_path>::value, "");
    static_assert(std::is_move_assignable<indexed_property_path>::value, "");
#endif

    struct folder_property_path final : public internal::no_assign
    {
        const property_path folder_id = "folder:FolderId";
        const property_path parent_folder_id = "folder:ParentFolderId";
        const property_path display_name = "folder:DisplayName";
        const property_path unread_count = "folder:UnreadCount";
        const property_path total_count = "folder:TotalCount";
        const property_path child_folder_count = "folder:ChildFolderCount";
        const property_path folder_class = "folder:FolderClass";
        const property_path search_parameters = "folder:SearchParameters";
        const property_path managed_folder_information = "folder:ManagedFolderInformation";
        const property_path permission_set = "folder:PermissionSet";
        const property_path effective_rights = "folder:EffectiveRights";
        const property_path sharing_effective_rights = "folder:SharingEffectiveRights";
    };

    struct item_property_path final : public internal::no_assign
    {
        const property_path item_id = "item:ItemId";
        const property_path parent_folder_id = "item:ParentFolderId";
        const property_path item_class = "item:ItemClass";
        const property_path mime_content = "item:MimeContent";
        const property_path attachment = "item:Attachments";
        const property_path subject = "item:Subject";
        const property_path date_time_received = "item:DateTimeReceived";
        const property_path size = "item:Size";
        const property_path categories = "item:Categories";
        const property_path has_attachments = "item:HasAttachments";
        const property_path importance = "item:Importance";
        const property_path in_reply_to = "item:InReplyTo";
        const property_path internet_message_headers = "item:InternetMessageHeaders";
        const property_path is_associated = "item:IsAssociated";
        const property_path is_draft = "item:IsDraft";
        const property_path is_from_me = "item:IsFromMe";
        const property_path is_resend = "item:IsResend";
        const property_path is_submitted = "item:IsSubmitted";
        const property_path is_unmodified = "item:IsUnmodified";
        const property_path date_time_sent = "item:DateTimeSent";
        const property_path date_time_created = "item:DateTimeCreated";
        const property_path body = "item:Body";
        const property_path response_objects = "item:ResponseObjects";
        const property_path sensitivity = "item:Sensitivity";
        const property_path reminder_due_by = "item:ReminderDueBy";
        const property_path reminder_is_set = "item:ReminderIsSet";
        const property_path reminder_next_time = "item:ReminderNextTime";
        const property_path reminder_minutes_before_start = "item:ReminderMinutesBeforeStart";
        const property_path display_to = "item:DisplayTo";
        const property_path display_cc = "item:DisplayCc";
        const property_path culture = "item:Culture";
        const property_path effective_rights = "item:EffectiveRights";
        const property_path last_modified_name = "item:LastModifiedName";
        const property_path last_modified_time = "item:LastModifiedTime";
        const property_path conversation_id = "item:ConversationId";
        const property_path unique_body = "item:UniqueBody";
        const property_path flag = "item:Flag";
        const property_path store_entry_id = "item:StoreEntryId";
        const property_path instance_key = "item:InstanceKey";
        const property_path normalized_body = "item:NormalizedBody";
        const property_path entity_extraction_result = "item:EntityExtractionResult";
        const property_path policy_tag = "item:PolicyTag";
        const property_path archive_tag = "item:ArchiveTag";
        const property_path retention_date = "item:RetentionDate";
        const property_path preview = "item:Preview";
        const property_path next_predicted_action = "item:NextPredictedAction";
        const property_path grouping_action = "item:GroupingAction";
        const property_path predicted_action_reasons = "item:PredictedActionReasons";
        const property_path is_clutter = "item:IsClutter";
        const property_path rights_management_license_data = "item:RightsManagementLicenseData";
        const property_path block_status = "item:BlockStatus";
        const property_path has_blocked_images = "item:HasBlockedImages";
        const property_path web_client_read_from_query_string = "item:WebClientReadFormQueryString";
        const property_path web_client_edit_from_query_string = "item:WebClientEditFormQueryString";
        const property_path text_body = "item:TextBody";
        const property_path icon_index = "item:IconIndex";
        const property_path mime_content_utf8 = "item:MimeContentUTF8";
    };

    struct message_property_path final : public internal::no_assign
    {
        const property_path conversation_index = "message:ConversationIndex";
        const property_path conversation_topic = "message:ConversationTopic";
        const property_path internet_message_id = "message:InternetMessageId";
        const property_path is_read = "message:IsRead";
        const property_path is_response_requested = "message:IsResponseRequested";
        const property_path is_read_receipt_requested = "message:IsReadReceiptRequested";
        const property_path is_delivery_receipt_requested = "message:IsDeliveryReceiptRequested";
        const property_path received_by = "message:ReceivedBy";
        const property_path received_representing = "message:ReceivedRepresenting";
        const property_path references = "message:References";
        const property_path reply_to = "message:ReplyTo";
        const property_path from = "message:From";
        const property_path sender = "message:Sender";
        const property_path to_recipients = "message:ToRecipients";
        const property_path cc_recipients = "message:CcRecipients";
        const property_path bcc_recipients = "message:BccRecipients";
        const property_path approval_request_data = "message:ApprovalRequestData";
        const property_path voting_information = "message:VotingInformation";
        const property_path reminder_message_data = "message:ReminderMessageData";
    };

    struct meeting_property_path final : public internal::no_assign
    {
        const property_path associated_calendar_item_id = "meeting:AssociatedCalendarItemId";
        const property_path is_delegated = "meeting:IsDelegated";
        const property_path is_out_of_date = "meeting:IsOutOfDate";
        const property_path has_been_processed = "meeting:HasBeenProcessed";
        const property_path response_type = "meeting:ResponseType";
        const property_path proposed_start = "meeting:ProposedStart";
        const property_path proposed_end = "meeting:PropsedEnd";
    };

    struct meeting_request_property_path final : public internal::no_assign
    {
        const property_path meeting_request_type = "meetingRequest:MeetingRequestType";
        const property_path intended_free_busy_status = "meetingRequest:IntendedFreeBusyStatus";
        const property_path change_highlights = "meetingRequest:ChangeHighlights";
    };

    struct calendar_property_path final : public internal::no_assign
    {
        const property_path start = "calendar:Start";
        const property_path end = "calendar:End";
        const property_path original_start = "calendar:OriginalStart";
        const property_path start_wall_clock = "calendar:StartWallClock";
        const property_path end_wall_clock = "calendar:EndWallClock";
        const property_path start_time_zone_id = "calendar:StartTimeZoneId";
        const property_path end_time_zone_id = "calendar:EndTimeZoneId";
        const property_path is_all_day_event = "calendar:IsAllDayEvent";
        const property_path legacy_free_busy_status = "calendar:LegacyFreeBusyStatus";
        const property_path location = "calendar:Location";
        const property_path when = "calendar:When";
        const property_path is_meeting = "calendar:IsMeeting";
        const property_path is_cancelled = "calendar:IsCancelled";
        const property_path is_recurring = "calendar:IsRecurring";
        const property_path meeting_request_was_sent = "calendar:MeetingRequestWasSent";
        const property_path is_response_requested = "calendar:IsResponseRequested";
        const property_path calendar_item_type = "calendar:CalendarItemType";
        const property_path my_response_type = "calendar:MyResponseType";
        const property_path organizer = "calendar:Organizer";
        const property_path required_attendees = "calendar:RequiredAttendees";
        const property_path optional_attendees = "calendar:OptionalAttendees";
        const property_path resources = "calendar:Resources";
        const property_path conflicting_meeting_count = "calendar:ConflictingMeetingCount";
        const property_path adjacent_meeting_count = "calendar:AdjacentMeetingCount";
        const property_path conflicting_meetings = "calendar:ConflictingMeetings";
        const property_path adjacent_meetings = "calendar:AdjacentMeetings";
        const property_path duration = "calendar:Duration";
        const property_path time_zone = "calendar:TimeZone";
        const property_path appointment_reply_time = "calendar:AppointmentReplyTime";
        const property_path appointment_sequence_number = "calendar:AppointmentSequenceNumber";
        const property_path appointment_state = "calendar:AppointmentState";
        const property_path recurrence = "calendar:Recurrence";
        const property_path first_occurrence = "calendar:FirstOccurrence";
        const property_path last_occurrence = "calendar:LastOccurrence";
        const property_path modified_occurrences = "calendar:ModifiedOccurrences";
        const property_path deleted_occurrences = "calendar:DeletedOccurrences";
        const property_path meeting_time_zone = "calendar:MeetingTimeZone";
        const property_path conference_type = "calendar:ConferenceType";
        const property_path allow_new_time_proposal = "calendar:AllowNewTimeProposal";
        const property_path is_online_meeting = "calendar:IsOnlineMeeting";
        const property_path meeting_workspace_url = "calendar:MeetingWorkspaceUrl";
        const property_path net_show_url = "calendar:NetShowUrl";
        const property_path uid = "calendar:UID";
        const property_path recurrence_id = "calendar:RecurrenceId";
        const property_path date_time_stamp = "calendar:DateTimeStamp";
        const property_path start_time_zone = "calendar:StartTimeZone";
        const property_path end_time_zone = "calendar:EndTimeZone";
        const property_path join_online_meeting_url = "calendar:JoinOnlineMeetingUrl";
        const property_path online_meeting_settings = "calendar:OnlineMeetingSettings";
        const property_path is_organizer = "calendar:IsOrganizer";
    };

    struct task_property_path final : public internal::no_assign
    {
        const property_path actual_work = "task:ActualWork";
        const property_path assigned_time = "task:AssignedTime";
        const property_path billing_information = "task:BillingInformation";
        const property_path change_count = "task:ChangeCount";
        const property_path companies = "task:Companies";
        const property_path complete_date = "task:CompleteDate";
        const property_path contacts = "task:Contacts";
        const property_path delegation_state = "task:DelegationState";
        const property_path delegator = "task:Delegator";
        const property_path due_date = "task:DueDate";
        const property_path is_assignment_editable = "task:IsAssignmentEditable";
        const property_path is_complete = "task:IsComplete";
        const property_path is_recurring = "task:IsRecurring";
        const property_path is_team_task = "task:IsTeamTask";
        const property_path mileage = "task:Mileage";
        const property_path owner = "task:Owner";
        const property_path percent_complete = "task:PercentComplete";
        const property_path recurrence = "task:Recurrence";
        const property_path start_date = "task:StartDate";
        const property_path status = "task:Status";
        const property_path status_description = "task:StatusDescription";
        const property_path total_work = "task:TotalWork";
    };

    struct contact_property_path final : public internal::no_assign
    {
        const property_path alias = "contacts:Alias";
        const property_path assistant_name = "contacts:AssistantName";
        const property_path birthday = "contacts:Birthday";
        const property_path business_home_page = "contacts:BusinessHomePage";
        const property_path children = "contacts:Children";
        const property_path companies = "contacts:Companies";
        const property_path company_name = "contacts:CompanyName";
        const property_path complete_name = "contacts:CompleteName";
        const property_path contact_source = "contacts:ContactSource";
        const property_path culture = "contacts:Culture";
        const property_path department = "contacts:Department";
        const property_path display_name = "contacts:DisplayName";
        const property_path directory_id = "contacts:DirectoryId";
        const property_path direct_reports = "contacts:DirectReports";
        const property_path email_addresses = "contacts:EmailAddresses";
        const indexed_property_path email_address_1{ "contacts:EmailAddress", "EmailAddress1" };
        const indexed_property_path email_address_2{ "contacts:EmailAddress", "EmailAddress2" };
        const indexed_property_path email_address_3{ "contacts:EmailAddress", "EmailAddress3" };
        const property_path file_as = "contacts:FileAs";
        const property_path file_as_mapping = "contacts:FileAsMapping";
        const property_path generation = "contacts:Generation";
        const property_path given_name = "contacts:GivenName";
        const property_path im_addresses = "contacts:ImAddresses";
        const property_path initials = "contacts:Initials";
        const property_path job_title = "contacts:JobTitle";
        const property_path manager = "contacts:Manager";
        const property_path manager_mailbox = "contacts:ManagerMailbox";
        const property_path middle_name = "contacts:MiddleName";
        const property_path mileage = "contacts:Mileage";
        const property_path ms_exchange_certificate = "contacts:MSExchangeCertificate";
        const property_path nickname = "contacts:Nickname";
        const property_path notes = "contacts:Notes";
        const property_path office_location = "contacts:OfficeLocation";
        const property_path phone_numbers = "contacts:PhoneNumbers";
        const property_path phonetic_full_name = "contacts:PhoneticFullName";
        const property_path phonetic_first_name = "contacts:PhoneticFirstName";
        const property_path phonetic_last_name = "contacts:PhoneticLastName";
        const property_path photo = "contacts:Photo";
        const property_path physical_address = "contacts:PhysicalAddresses";
        const property_path postal_adress_index = "contacts:PostalAddressIndex";
        const property_path profession = "contacts:Profession";
        const property_path spouse_name = "contacts:SpouseName";
        const property_path surname = "contacts:Surname";
        const property_path wedding_anniversary = "contacts:WeddingAnniversary";
        const property_path smime_certificate = "contacts:UserSMIMECertificate";
        const property_path has_picture = "contacts:HasPicture";
    };

    struct distribution_list_property_path final : public internal::no_assign
    {
        const property_path members = "distributionlist:Members";
    };

    struct post_item_property_path final : public internal::no_assign
    {
        const property_path posted_time = "postitem:PostedTime";
    };

    struct conversation_property_path final : public internal::no_assign
    {
        const property_path conversation_id = "conversation:ConversationId";
        const property_path conversation_topic = "conversation:ConversationTopic";
        const property_path unique_recipients = "conversation:UniqueRecipients";
        const property_path global_unique_recipients = "conversation:GlobalUniqueRecipients";
        const property_path unique_unread_senders = "conversation:UniqueUnreadSenders";
        const property_path global_unique_unread_readers = "conversation:GlobalUniqueUnreadSenders";
        const property_path unique_senders = "conversation:UniqueSenders";
        const property_path global_unique_senders = "conversation:GlobalUniqueSenders";
        const property_path last_delivery_time = "conversation:LastDeliveryTime";
        const property_path global_last_delivery_time = "conversation:GlobalLastDeliveryTime";
        const property_path categories = "conversation:Categories";
        const property_path global_categories = "conversation:GlobalCategories";
        const property_path flag_status = "conversation:FlagStatus";
        const property_path global_flag_status = "conversation:GlobalFlagStatus";
        const property_path has_attachments = "conversation:HasAttachments";
        const property_path global_has_attachments = "conversation:GlobalHasAttachments";
        const property_path has_irm = "conversation:HasIrm";
        const property_path global_has_irm = "conversation:GlobalHasIrm";
        const property_path message_count = "conversation:MessageCount";
        const property_path global_message_count = "conversation:GlobalMessageCount";
        const property_path unread_count = "conversation:UnreadCount";
        const property_path global_unread_count = "conversation:GlobalUnreadCount";
        const property_path size = "conversation:Size";
        const property_path global_size = "conversation:GlobalSize";
        const property_path item_classes = "conversation:ItemClasses";
        const property_path global_item_classes = "conversation:GlobalItemClasses";
        const property_path importance = "conversation:Importance";
        const property_path global_importance = "conversation:GlobalImportance";
        const property_path item_ids = "conversation:ItemIds";
        const property_path global_item_ids = "conversation:GlobalItemIds";
        const property_path last_modified_time = "conversation:LastModifiedTime";
        const property_path instance_key = "conversation:InstanceKey";
        const property_path preview = "conversation:Preview";
        const property_path global_parent_folder_id = "conversation:GlobalParentFolderId";
        const property_path next_predicted_action = "conversation:NextPredictedAction";
        const property_path grouping_action = "conversation:GroupingAction";
        const property_path icon_index = "conversation:IconIndex";
        const property_path global_icon_index = "conversation:GlobalIconIndex";
        const property_path draft_item_ids = "conversation:DraftItemIds";
        const property_path has_clutter = "conversation:HasClutter";
    };

    // Represents a single property
    //
    // Used in ews::service::update_item method call
    class property final
    {
    public:
        // Use this constructor if you want to delete a property from an item
        explicit property(property_path path)
            : path_(std::move(path)),
              value_()
        {
        }

        // Use this constructor (and following overloads) whenever you want to
        // set or update an item's property
        property(property_path path, std::string value)
            : path_(std::move(path)),
              value_(std::move(value))
        {
        }

        property(property_path path, const char* value)
            : path_(std::move(path)),
              value_(std::string(value))
        {
        }

        property(property_path path, int value)
            : path_(std::move(path)),
              value_(std::to_string(value))
        {
        }

        property(property_path path, long value)
            : path_(std::move(path)),
              value_(std::to_string(value))
        {
        }

        property(property_path path, long long value)
            : path_(std::move(path)),
              value_(std::to_string(value))
        {
        }

        property(property_path path, unsigned value)
            : path_(std::move(path)),
              value_(std::to_string(value))
        {
        }

        property(property_path path, unsigned long value)
            : path_(std::move(path)),
              value_(std::to_string(value))
        {
        }

        property(property_path path, unsigned long long value)
            : path_(std::move(path)),
              value_(std::to_string(value))
        {
        }

        property(property_path path, float value)
            : path_(std::move(path)),
              value_(std::to_string(value))
        {
        }

        property(property_path path, double value)
            : path_(std::move(path)),
              value_(std::to_string(value))
        {
        }

        property(property_path path, long double value)
            : path_(std::move(path)),
              value_(std::to_string(value))
        {
        }

        property(property_path path, bool value)
            : path_(std::move(path)),
              value_(value ? "true" : "false")
        {
        }

        property(property_path path, const body& value)
            : path_(std::move(path)),
              value_(value.to_xml("t"))
        {
        }

        property(property_path path, const std::vector<email_address>& value)
            : path_(std::move(path)),
              value_()
        {
            std::stringstream sstr;
            for (const auto& addr : value)
            {
                sstr << addr.to_xml("t");
            }
            value_ = sstr.str();
        }

        std::string to_xml(const char* xmlns) const
        {
            std::stringstream sstr;
            auto pref = std::string();
            if (xmlns)
            {
                pref = std::string(xmlns) + ":";
            }
            sstr << "<" << pref << "FieldURI FieldURI=\"";
            sstr << path().field_uri() << "\"/>";
            sstr << "<" << pref << path().class_name() << ">";
            sstr << "<" << pref << path().property_name() << ">";
            sstr << value_;
            sstr << "</" << pref << path().property_name() << ">";
            sstr << "</" << pref << path().class_name() << ">";
            return sstr.str();
        }

        bool empty_value() const EWS_NOEXCEPT { return value_.empty(); }

        const property_path& path() const EWS_NOEXCEPT { return path_; }

    private:
        property_path path_;
        std::string value_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(!std::is_default_constructible<property>::value, "");
    static_assert(std::is_copy_constructible<property>::value, "");
    static_assert(std::is_copy_assignable<property>::value, "");
    static_assert(std::is_move_constructible<property>::value, "");
    static_assert(std::is_move_assignable<property>::value, "");
#endif

    // Base-class for
    //
    //   - exists
    //   - excludes
    //   - is_equal_to
    //   - is_not_equal_to
    //   - is_greater_than
    //   - is_greater_than_or_equal_to
    //   - is_less_than
    //   - is_less_than_or_equal_to
    //   - contains
    //   - not
    //   - and
    //   - or
    //
    // search expressions.
    class restriction
    {
    public:
        ~restriction() = default;

        std::string to_xml(const char* xmlns=nullptr) const
        {
            return func_(xmlns);
        }

    protected:
        explicit restriction(std::function<std::string (const char*)>&& func)
            : func_(std::move(func))
        {
        }

    private:
        std::function<std::string (const char*)> func_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(!std::is_default_constructible<restriction>::value, "");
    static_assert(std::is_copy_constructible<restriction>::value, "");
    static_assert(std::is_copy_assignable<restriction>::value, "");
    static_assert(std::is_move_constructible<restriction>::value, "");
    static_assert(std::is_move_assignable<restriction>::value, "");
#endif

    // A search expression that compares a property with either a constant value
    // or another property and evaluates to true if they are equal.
    class is_equal_to final : public restriction
    {
    public:
        is_equal_to(property_path path, bool b)
            : restriction([=](const char* xmlns) -> std::string
                    {
                        std::stringstream sstr;
                        auto pref = std::string();
                        if (xmlns)
                        {
                            pref = std::string(xmlns) + ":";
                        }
                        sstr << "<" << pref << "IsEqualTo><" << pref;
                        sstr << "FieldURI FieldURI=\"";
                        sstr << path.field_uri();
                        sstr << "\"/><" << pref << "FieldURIOrConstant><";
                        sstr << pref << "Constant Value=\"";
                        sstr << std::boolalpha << b;
                        sstr << "\"/></" << pref << "FieldURIOrConstant></";
                        sstr << pref << "IsEqualTo>";
                        return sstr.str();
                    })
        {
        }

        is_equal_to(property_path path, const char* str)
            : restriction([=](const char* xmlns) -> std::string
                    {
                        std::stringstream sstr;
                        const char* pref = "";
                        if (xmlns)
                        {
                            pref = "t:";
                        }
                        sstr << "<" << pref << "IsEqualTo><" << pref;
                        sstr << "FieldURI FieldURI=\"";
                        sstr << path.field_uri();
                        sstr << "\"/><" << pref << "FieldURIOrConstant><";
                        sstr << pref << "Constant Value=\"";
                        sstr << str;
                        sstr << "\"/></" << pref << "FieldURIOrConstant></";
                        sstr << pref << "IsEqualTo>";
                        return sstr.str();
                    })
        {
        }

        is_equal_to(indexed_property_path path, const char* str)
            : restriction([=](const char* xmlns) -> std::string
                    {
                        std::stringstream sstr;
                        const char* pref = "";
                        if (xmlns)
                        {
                            pref = "t:";
                        }
                        sstr << "<" << pref << "IsEqualTo><" << pref;
                        sstr << "IndexedFieldURI FieldURI=\"";
                        sstr << path.field_uri();
                        sstr << "\" FieldIndex=\"" << path.field_index();
                        sstr << "\"/><" << pref << "FieldURIOrConstant><";
                        sstr << pref << "Constant Value=\"";
                        sstr << str;
                        sstr << "\"/></" << pref << "FieldURIOrConstant></";
                        sstr << pref << "IsEqualTo>";
                        return sstr.str();
                    })
        {
        }

        is_equal_to(property_path path, date_time when)
            : restriction([=](const char* xmlns) -> std::string
                    {
                        std::stringstream sstr;
                        const char* pref = "";
                        if (xmlns)
                        {
                            pref = "t:";
                        }
                        sstr << "<" << pref << "IsEqualTo><" << pref;
                        sstr << "FieldURI FieldURI=\"";
                        sstr << path.field_uri();
                        sstr << "\"/><" << pref << "FieldURIOrConstant><";
                        sstr << pref << "Constant Value=\"";
                        sstr << when.to_string();
                        sstr << "\"/></" << pref << "FieldURIOrConstant></";
                        sstr << pref << "IsEqualTo>";
                        return sstr.str();
                    })
        {
        }

        // TODO: is_equal_to(property_path, property_path) {}
    };

    // Represents a <FileAttachment> or a <ItemAttachment>
    class attachment final
    {
    public:
        enum class type { item, file };

        attachment() = default;

        const attachment_id& id() const EWS_NOEXCEPT { return id_; }

        const std::string& name() const EWS_NOEXCEPT { return name_; }

        const std::string& content_type() const EWS_NOEXCEPT
        {
            return content_type_;
        }

        // Returns either ews::attachment::type::file or
        // ews::attachment::type::item
        type get_type() const EWS_NOEXCEPT { return type_; }

        // If this is a <FileAttachment>, returns the Base64-encoded contents
        // of the file attachment. If this is an <ItemAttachment>, the empty
        // string.
        const std::string& content() const EWS_NOEXCEPT
        {
            return content_;
        }

        // If this is a <FileAttachment>, returns the size in bytes of the file
        // attachment; otherwise 0.
        std::size_t content_size() const EWS_NOEXCEPT { return content_size_; }

        // If this is a <FileAttachment>, writes content to file.  Does nothing
        // if this is an <ItemAttachment>.  Returns the number of bytes
        // written.
        std::size_t write_content_to_file(const std::string& file_path)
        {
            if (get_type() == type::item)
            {
                return 0U;
            }

            const auto raw_bytes = internal::base64::decode(content_);

            auto ofstr = std::ofstream(file_path,
                                       std::ofstream::out | std::ios::binary);
            if (!ofstr.is_open())
            {
                if (file_path.empty())
                {
                    throw exception(
                        "Could not open file for writing: no file name given");
                }

                throw exception(
                        "Could not open file for writing: " + file_path);
            }

            std::copy(begin(raw_bytes), end(raw_bytes),
                      std::ostreambuf_iterator<char>(ofstr));
            ofstr.close();
            return raw_bytes.size();
        }

        // std::string to_xml(const char* xmlns=nullptr) const
        // {}

        static attachment from_xml_element(const rapidxml::xml_node<>& elem)
        {
            using uri = internal::uri<>;

            const auto elem_name = std::string(elem.local_name(),
                                               elem.local_name_size());
            EWS_ASSERT((elem_name == "FileAttachment"
                     || elem_name == "ItemAttachment")
                    && "Expected <FileAttachment> or <ItemAttachment>");

            auto obj = attachment();
            if (elem_name == "FileAttachment")
            {
                obj.type_ = type::file;
                auto id_node = elem.first_node_ns(uri::microsoft::types(),
                                                 "AttachmentId");
                EWS_ASSERT(id_node && "Expected <AttachmentId>");
                obj.id_ = attachment_id::from_xml_element(*id_node);

                auto name_node = elem.first_node_ns(uri::microsoft::types(),
                                                    "Name");
                if (name_node)
                {
                    obj.name_ = std::string(name_node->value(),
                                            name_node->value_size());
                }

                auto content_type_node =
                    elem.first_node_ns(uri::microsoft::types(), "ContentType");
                if (content_type_node)
                {
                    obj.content_type_ =
                        std::string(content_type_node->value(),
                                    content_type_node->value_size());
                }

                auto content_node = elem.first_node_ns(uri::microsoft::types(),
                                                       "Content");
                if (content_node)
                {
                    obj.content_ = std::string(content_node->value(),
                                               content_node->value_size());
                }
            }
            else
            {
                obj.type_ = type::item;
            }
            return obj;
        }

        // Creates a new <FileAttachment> from a given file.
        //
        // Returns a new <FileAttachment> that you can pass to
        // ews::service::create_attachment in order to create the attachment on
        // the server.
        static attachment from_file(std::string file_path,
                                    std::string content_type,
                                    std::string name)
        {
            // Try open file
            auto ifstr = std::ifstream(file_path,
                                       std::ifstream::in | std::ios::binary);
            if (!ifstr.is_open())
            {
                throw exception(
                        "Could not open file for reading: " + file_path);
            }

            // Stop eating newlines in binary mode
            ifstr.unsetf(std::ios::skipws);

            // Determine size
            ifstr.seekg(0, std::ios::end);
            const auto file_size = ifstr.tellg();
            ifstr.seekg(0, std::ios::beg);

            // Allocate buffer
            auto buffer = std::vector<unsigned char>();
            buffer.reserve(file_size);

            // Read
            buffer.insert(begin(buffer),
                          std::istream_iterator<unsigned char>(ifstr),
                          std::istream_iterator<unsigned char>());
            ifstr.close();

            // And encode
            auto content = internal::base64::encode(buffer);

            auto obj = attachment();
            obj.type_ = type::file;
            obj.name_ = std::move(name);
            obj.content_type_ = std::move(content_type);
            obj.content_ = std::move(content);
            obj.content_size_ = buffer.size();
            return obj;
        }

        // Creates a new <ItemAttachment> from a given item.
        //
        // It is not necessary for the item to already exist in the Exchange
        // store. If it doesn't, it will be automatically created.
        static attachment from_item(const item& the_item, std::string name)
        {
            (void)the_item;

            // Creating a new <ItemAttachment> with the <CreateAttachment>
            // method is pretty similar to a <CreateItem> method call. However,
            // most of the times we do not want to create item attachments out
            // of thin air but attach an _existing_ item.
            //
            // If we want create an attachment from an existing item, we need
            // to first call <GetItem> before we call <CreateItem> and put the
            // complete item from the response into the <CreateAttachment>
            // call.
            //
            // There is a shortcut: use <BaseShape>IdOnly</BaseShape> and
            // <AdditionalProperties> with item::MimeContent in <GetItem> call,
            // remove <ItemId> from the response and pass that to
            // <CreateAttachment>.

            auto obj = attachment();
            obj.name_ = std::move(name);
            return obj;
        }

    private:
        attachment_id id_;
        std::string name_;
        std::string content_type_;
        std::string content_;
        std::size_t content_size_;
        type type_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<attachment>::value, "");
    static_assert(std::is_copy_constructible<attachment>::value, "");
    static_assert(std::is_copy_assignable<attachment>::value, "");
    static_assert(std::is_move_constructible<attachment>::value, "");
    static_assert(std::is_move_assignable<attachment>::value, "");
#endif

    // The service class contains all methods that can be performed on an
    // Exchange server.
    //
    // Will get a *huge* public interface over time, e.g.,
    //
    // - create_item
    // - find_conversation
    // - find_folder
    // - find_item
    // - find_people
    // - get_contact
    // - get_task
    //
    // and so on and so on.
    template <typename RequestHandler = internal::http_request>
    class basic_service final
    {
    public:
        // FIXME: credentials are stored plain-text in memory
        //
        // That'll be bad. We wouldn't want random Joe at first-level support to
        // see plain-text passwords and user-names just because the process
        // crashed and some automatic mechanism sent a minidump over the wire.
        // What are our options? Security-by-obscurity: we could hash
        // credentials with a hash of the process-id or something.
        basic_service(std::string server_uri,
                      std::string domain,
                      std::string username,
                      std::string password)
            : server_uri_(std::move(server_uri)),
              domain_(std::move(domain)),
              username_(std::move(username)),
              password_(std::move(password)),
              server_version_("Exchange2013_SP1")
        {
        }

        void set_request_server_version(server_version vers)
        {
            server_version_ = internal::server_version_to_str(vers);
        }

        server_version get_request_server_version() const
        {
            return internal::str_to_server_version(server_version_);
        }

        // Gets a task from the Exchange store
        task get_task(const item_id& id)
        {
            return get_item_impl<task>(id, base_shape::all_properties);
        }

        // Gets a contact from the Exchange store
        contact get_contact(const item_id& id)
        {
            return get_item_impl<contact>(id, base_shape::all_properties);
        }

        // Gets a message item from the Exchange store
        message get_message(const item_id& id)
        {
            return get_item_impl<message>(id, base_shape::all_properties);
        }

        // Delete an arbitrary item from the Exchange store
        void delete_item(item&& the_item)
        {
            delete_item_impl<item>(std::move(the_item));
        }

        // Delete a task item from the Exchange store
        void delete_task(task&& the_task,
                         delete_type del_type = delete_type::hard_delete,
                         affected_task_occurrences affected =
                            affected_task_occurrences::all_occurrences)
        {
            using internal::delete_item_response_message;

            const std::string request_string =
R"(
    <m:DeleteItem
      DeleteType=")" + delete_type_str(del_type) + R"("
      AffectedTaskOccurrences=")" + affected_task_occurrences_str(affected) + R"(">
      <m:ItemIds>)" + the_task.get_item_id().to_xml("t") + R"(</m:ItemIds>
    </m:DeleteItem>
)";
            auto response = request(request_string);
            const auto response_message =
                delete_item_response_message::parse(response);
            if (!response_message.success())
            {
                throw exchange_error(response_message.get_response_code());
            }
            the_task = task();
        }

        // Delete a contact from the Exchange store
        void delete_contact(contact&& the_contact)
        {
            delete_item_impl(std::move(the_contact));
        }

        // Delete a message item from the Exchange store
        void delete_message(message&& the_message)
        {
            delete_item_impl(std::move(the_message));
        }

        // Following items can be created in Exchange:
        //
        // - Calendar items
        // - E-mail messages
        // - Meeting requests
        // - Tasks
        // - Contacts
        //
        // Only purpose of this overload-set is to prevent exploding template
        // code in errors messages in caller's code

        // Creates a new task item from the given object in the Exchange store
        // and returns it's item_id if successful
        item_id create_item(const task& the_task)
        {
            return create_item_impl(the_task);
        }

        // Creates a new contact item from the given object in the Exchange
        // store and returns it's item_id if successful
        item_id create_item(const contact& the_contact)
        {
            return create_item_impl(the_contact);
        }

        item_id create_item(const message& the_message,
                            ews::message_disposition disposition)
        {
            using internal::create_item_response_message;

            auto response =
                request(the_message.create_item_request_string(disposition));
#ifdef EWS_ENABLE_VERBOSE
            std::cerr << response.payload() << std::endl;
#endif
            const auto response_message =
                create_item_response_message::parse(response);
            if (!response_message.success())
            {
                throw exchange_error(response_message.get_response_code());
            }
            EWS_ASSERT(!response_message.items().empty()
                    && "Expected a message item");
            return response_message.items().front();
        }

        std::vector<item_id> find_item(const folder_id& parent_folder_id)
        {
            using find_item_response_message =
                internal::find_item_response_message;

            const std::string request_string =
R"(
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:ItemShape>
      <m:ParentFolderIds>)" + parent_folder_id.to_xml("t") + R"(</m:ParentFolderIds>
    </m:FindItem>
)";
            auto response = request(request_string);
#ifdef EWS_ENABLE_VERBOSE
            std::cerr << response.payload() << std::endl;
#endif
            const auto response_message =
                find_item_response_message::parse(response);
            if (!response_message.success())
            {
                throw exchange_error(response_message.get_response_code());
            }
            return response_message.items();
        }

        std::vector<item_id> find_item(const folder_id& parent_folder_id,
                                       restriction filter)
        {
            using find_item_response_message =
                internal::find_item_response_message;

            const std::string request_string =
R"(
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:ItemShape>
      <m:Restriction>)" + filter.to_xml("t") + R"(</m:Restriction>
      <m:ParentFolderIds>)" + parent_folder_id.to_xml("t") + R"(</m:ParentFolderIds>
    </m:FindItem>
)";
            auto response = request(request_string);
#ifdef EWS_ENABLE_VERBOSE
            std::cerr << response.payload() << std::endl;
#endif
            const auto response_message =
                find_item_response_message::parse(response);
            if (!response_message.success())
            {
                throw exchange_error(response_message.get_response_code());
            }
            return response_message.items();
        }

        // TODO: currently, this can only do <SetItemField>, need to support
        // <AppendToItemField> and <DeleteItemField>
        item_id
        update_item(item_id id,
                    property prop,
                    conflict_resolution res = conflict_resolution::auto_resolve)
        {
            auto item_change_open_tag = "<t:SetItemField>";
            auto item_change_close_tag = "</t:SetItemField>";
            if (   prop.path() == "calendar:OptionalAttendees"
                || prop.path() == "calendar:RequiredAttendees"
                || prop.path() == "calendar:Resources"
                || prop.path() == "item:Body"
                || prop.path() == "message:ToRecipients"
                || prop.path() == "message:CcRecipients"
                || prop.path() == "message:BccRecipients"
                || prop.path() == "message:ReplyTo")
            {
                item_change_open_tag = "<t:AppendToItemField>";
                item_change_close_tag = "</t:AppendToItemField>";
            }

            const std::string request_string =
R"(
    <m:UpdateItem
        MessageDisposition="SaveOnly"
        ConflictResolution=")" + conflict_resolution_str(res) + R"(">
      <m:ItemChanges>
        <t:ItemChange>
          )" + id.to_xml("t") + R"(
          <t:Updates>
            )" + item_change_open_tag + R"(
              )" + prop.to_xml("t") + R"(
            )" + item_change_close_tag + R"(
          </t:Updates>
        </t:ItemChange>
      </m:ItemChanges>
    </m:UpdateItem>
)";

            auto response = request(request_string);
#ifdef EWS_ENABLE_VERBOSE
            std::cerr << response.payload() << std::endl;
#endif
            const auto response_message =
                internal::update_item_response_message::parse(response);
            if (!response_message.success())
            {
                throw exchange_error(response_message.get_response_code());
            }
            EWS_ASSERT(!response_message.items().empty()
                    && "Expected at least one item");
            return response_message.items().front();
        }

        // content_type: the (RFC 2046) MIME content type of the attachment. On
        // Windows you can use HKEY_CLASSES_ROOT/MIME/Database/Content Type
        // registry hive to get the content type from a file extension. On a
        // UNIX see magic(5) and file(1).
        attachment_id create_attachment(const item& parent_item,
                                        const attachment& a)
        {
            (void)parent_item;
            (void)a;
            return attachment_id();
        }

        std::vector<attachment_id>
        create_attachment(const item& parent_item,
                          const std::vector<attachment>& attachments)
        {
            (void)parent_item;
            (void)attachments;
            return std::vector<attachment_id>();
        }

    private:
        std::string server_uri_;
        std::string domain_;
        std::string username_;
        std::string password_;
        std::string server_version_;

        // Helper for doing requests.  Adds the right headers, credentials, and
        // checks the response for faults.
        internal::http_response request(const std::string& request_string)
        {
            const auto soap_headers = std::vector<std::string> {
                "<t:RequestServerVersion Version=\"" + server_version_ + "\"/>"
            };
            auto response =
                internal::make_raw_soap_request<RequestHandler>(server_uri_,
                                                                username_,
                                                                password_,
                                                                domain_,
                                                                request_string,
                                                                soap_headers);
            internal::raise_exception_if_soap_fault(response);
            return response;
        }

        // Gets an item from the server.
        template <typename ItemType>
        ItemType get_item_impl(const item_id& id, base_shape shape)
        {
            using get_item_response_message =
                internal::get_item_response_message<ItemType>;

            // TODO: remove <AdditionalProperties> below, add parameter(s) or
            // overload to allow users customization
            const std::string request_string =
R"(
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>)" + base_shape_str(shape) + R"(</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:MimeContent"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>)" + id.to_xml("t") + R"(</m:ItemIds>
    </m:GetItem>
)";
            auto response = request(request_string);
#ifdef EWS_ENABLE_VERBOSE
            std::cerr << response.payload() << std::endl;
#endif
            const auto response_message =
                get_item_response_message::parse(response);
            if (!response_message.success())
            {
                throw exchange_error(response_message.get_response_code());
            }
            EWS_ASSERT(!response_message.items().empty()
                    && "Expected at least one item");
            return response_message.items().front();
        }

        // Creates an item on the server and returns it's item_id.
        template <typename ItemType>
        item_id create_item_impl(const ItemType& the_item)
        {
            using internal::create_item_response_message;

            auto response = request(the_item.create_item_request_string());
#ifdef EWS_ENABLE_VERBOSE
            std::cerr << response.payload() << std::endl;
#endif
            const auto response_message =
                create_item_response_message::parse(response);
            if (!response_message.success())
            {
                throw exchange_error(response_message.get_response_code());
            }
            EWS_ASSERT(!response_message.items().empty()
                    && "Expected at least one item");
            return response_message.items().front();
        }

        template <typename ItemType>
        void delete_item_impl(ItemType&& the_item)
        {
            using internal::delete_item_response_message;

            const std::string request_string =
R"(
    <m:DeleteItem>
      <m:ItemIds>)" + the_item.get_item_id().to_xml("t") + R"(</m:ItemIds>
    </m:DeleteItem>
)";
            auto response = request(request_string);
#ifdef EWS_ENABLE_VERBOSE
            std::cerr << response.payload() << std::endl;
#endif
            const auto response_message =
                delete_item_response_message::parse(response);
            if (!response_message.success())
            {
                throw exchange_error(response_message.get_response_code());
            }
            the_item = ItemType();
        }
    };

    typedef basic_service<> service;

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(!std::is_default_constructible<service>::value, "");
    static_assert(std::is_copy_constructible<service>::value, "");
    static_assert(std::is_copy_assignable<service>::value, "");
    static_assert(std::is_move_constructible<service>::value, "");
    static_assert(std::is_move_assignable<service>::value, "");
#endif

    // Implementations

    namespace internal
    {
        inline void ntlm_credentials::certify(http_request* request) const
        {
            EWS_ASSERT(request != nullptr);

            // CURLOPT_USERPWD: domain\username:password
            std::string login = domain_ + "\\" + username_ + ":" + password_;
            request->set_option(CURLOPT_USERPWD, login.c_str());
            request->set_option(CURLOPT_HTTPAUTH, CURLAUTH_NTLM);
        }

        // FIXME: a CreateItemResponse can contain multiple ResponseMessages
        inline create_item_response_message
        create_item_response_message::parse(http_response& response)
        {
            using uri = internal::uri<>;
            using internal::get_element_by_qname;

            const auto& doc = response.payload();
            auto elem = get_element_by_qname(doc,
                                             "CreateItemResponseMessage",
                                             uri::microsoft::messages());

#ifdef EWS_ENABLE_VERBOSE
            // FIXME: sometimes assertion fails for reasons unknown to me,
            // temporarily log response code and XML payload here
            if (!elem)
            {
                std::cerr
                    << "Parsing CreateItemResponseMessage failed, response code: "
                    << response.code() << ", payload:\n" << doc << std::endl;
            }
#endif

            EWS_ASSERT(elem &&
                    "Expected <CreateItemResponseMessage>, got nullptr");
            const auto result = parse_response_class_and_code(*elem);
            auto item_ids = std::vector<item_id>();
            for_each_item(*elem,
                    [&item_ids](const rapidxml::xml_node<>& item_elem)
            {
                auto item_id_elem = item_elem.first_node();
                EWS_ASSERT(item_id_elem && "Expected <ItemId> element");
                item_ids.emplace_back(
                        item_id::from_xml_element(*item_id_elem));
            });
            return create_item_response_message(result.first,
                                                result.second,
                                                std::move(item_ids));
        }

        inline find_item_response_message
        find_item_response_message::parse(http_response& response)
        {
            using uri = internal::uri<>;
            using internal::get_element_by_qname;

            const auto& doc = response.payload();
            auto elem = get_element_by_qname(doc,
                                             "FindItemResponseMessage",
                                             uri::microsoft::messages());

            EWS_ASSERT(elem &&
                "Expected <FindItemResponseMessage>, got nullptr");
            const auto result = parse_response_class_and_code(*elem);

            auto root_folder = elem->first_node_ns(uri::microsoft::messages(),
                                                   "RootFolder");

            auto items_elem =
                root_folder->first_node_ns(uri::microsoft::types(), "Items");
            EWS_ASSERT(items_elem && "Expected <t:Items> element");

            auto items = std::vector<item_id>();
            for (auto item_elem = items_elem->first_node(); item_elem;
                 item_elem = item_elem->next_sibling())
            {
                EWS_ASSERT(item_elem && "Expected an element, got nullptr");
                auto item_id_elem = item_elem->first_node();
                EWS_ASSERT(item_id_elem && "Expected <ItemId> element");
                items.emplace_back(
                    item_id::from_xml_element(*item_id_elem));
            }
            return find_item_response_message(result.first,
                                              result.second,
                                              std::move(items));
        }

        inline update_item_response_message
        update_item_response_message::parse(http_response& response)
        {
            using uri = internal::uri<>;
            using internal::get_element_by_qname;

            const auto& doc = response.payload();
            auto elem = get_element_by_qname(doc,
                                             "UpdateItemResponseMessage",
                                             uri::microsoft::messages());

            EWS_ASSERT(elem &&
                "Expected <UpdateItemResponseMessage>, got nullptr");
            const auto result = parse_response_class_and_code(*elem);

            auto items_elem =
                elem->first_node_ns(uri::microsoft::messages(), "Items");
            EWS_ASSERT(items_elem && "Expected <m:Items> element");

            auto items = std::vector<item_id>();
            for (auto item_elem = items_elem->first_node(); item_elem;
                 item_elem = item_elem->next_sibling())
            {
                EWS_ASSERT(item_elem && "Expected an element, got nullptr");
                auto item_id_elem = item_elem->first_node();
                EWS_ASSERT(item_id_elem && "Expected <ItemId> element");
                items.emplace_back(
                    item_id::from_xml_element(*item_id_elem));
            }
            return update_item_response_message(result.first,
                                                result.second,
                                                std::move(items));
        }

        template <typename ItemType>
        inline
        get_item_response_message<ItemType>
        get_item_response_message<ItemType>::parse(http_response& response)
        {
            using uri = internal::uri<>;
            using internal::get_element_by_qname;

            const auto& doc = response.payload();
            auto elem = get_element_by_qname(doc,
                                             "GetItemResponseMessage",
                                             uri::microsoft::messages());
#ifdef EWS_ENABLE_VERBOSE
            if (!elem)
            {
                std::cerr
                    << "Parsing GetItemResponseMessage failed, response code: "
                    << response.code() << ", payload:\n" << doc << std::endl;
            }
#endif
            EWS_ASSERT(elem &&
                    "Expected <GetItemResponseMessage>, got nullptr");
            const auto result = parse_response_class_and_code(*elem);
            auto items = std::vector<ItemType>();
            for_each_item(*elem, [&items](const rapidxml::xml_node<>& item_elem)
            {
                items.emplace_back(ItemType::from_xml_element(item_elem));
            });
            return get_item_response_message(result.first,
                                             result.second,
                                             std::move(items));
        }
    }
}
