
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

#pragma once

#include <algorithm>
#include <cctype>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <exception>
#include <fstream>
#include <functional>
#include <iomanip>
#include <ios>
#include <iterator>
#include <memory>
#include <sstream>
#include <stdexcept>
#include <string>
#include <tuple>
#include <type_traits>
#include <utility>
#include <vector>

#include <curl/curl.h>

#include "rapidxml/rapidxml.hpp"
#include "rapidxml/rapidxml_print.hpp"

// Print more detailed error messages, HTTP request/response etc to stderr
#ifdef EWS_ENABLE_VERBOSE
#include <iostream>
#include <ostream>
#endif

#include "ews_fwd.hpp"

// Macro for verifying expressions at run-time. Calls assert() with 'expr'.
// Allows turning assertions off, even if -DNDEBUG wasn't given at
// compile-time.  This macro does nothing unless EWS_ENABLE_ASSERTS was defined
// during compilation.
#define EWS_ASSERT(expr) ((void)0)
#ifndef NDEBUG
#ifdef EWS_ENABLE_ASSERTS
#include <cassert>
#undef EWS_ASSERT
#define EWS_ASSERT(expr) assert(expr)
#endif
#endif // !NDEBUG

// To work on windows with or without NOMINMAX defined.
// Allows bare min(a,b), max(a,b) to be used with no
// qualifying namespace.
#if !defined(_WIN32) || defined(NOMINMAX)
namespace ews
{
template <class T>
inline const T& min(const T& a, const T& b)
{
    return std::min(a, b);
}

template <class T>
inline const T& max(const T& a, const T& b)
{
    return std::max(a, b);
}
} // namespace ews
#endif // !_WIN32 || NOMINMAX

// Workarounds for missing functions in Android NDK
#ifdef __ANDROID__
namespace std
{
    template <class T>
    inline std::string to_string(const T& val)
    {
        std::ostringstream os;
        os << val;
        return os.str();
    }

    inline int stoi(const std::string& str)
    {
        return atoi(str.c_str());
    }

    // This is an incomplete implementation, supporting only bases 8, 10, and 16
    inline unsigned long stoul(const std::string& str, std::size_t* pos = 0, int base = 10)
    {
        std::istringstream is(str);
        unsigned long ul;
        is >> std::setbase(base) >> std::skipws >> ul;
        if (pos)
        {
            *pos = is.gcount();
        }
        return ul;
    }
}
#endif // __ANDROID__


//! Contains all classes, functions, and enumerations of this library
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
        {
        }
        catch (std::exception&)
        {
            destructor_function();
        }

#ifdef EWS_HAS_DEFAULT_AND_DELETE
        on_scope_exit() = delete;
        on_scope_exit(const on_scope_exit&) = delete;
        on_scope_exit& operator=(const on_scope_exit&) = delete;
#else
    private:
        on_scope_exit(const on_scope_exit&);            // Never defined
        on_scope_exit& operator=(const on_scope_exit&); // Never defined

    public:
#endif

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

        void release() EWS_NOEXCEPT { func_ = nullptr; }

    private:
        std::function<void(void)> func_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(!std::is_copy_constructible<on_scope_exit>::value, "");
    static_assert(!std::is_copy_assignable<on_scope_exit>::value, "");
    static_assert(!std::is_move_constructible<on_scope_exit>::value, "");
    static_assert(!std::is_move_assignable<on_scope_exit>::value, "");
    static_assert(!std::is_default_constructible<on_scope_exit>::value, "");
#endif

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
            static const auto chars = std::string("ABCDEFGHIJKLMNOPQRSTUVWXYZ"
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
                    char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
                    char_array_4[1] = ((char_array_3[0] & 0x03) << 4) +
                                      ((char_array_3[1] & 0xf0) >> 4);
                    char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) +
                                      ((char_array_3[2] & 0xc0) >> 6);
                    char_array_4[3] = char_array_3[2] & 0x3f;

                    for (i = 0; (i < 4); i++)
                    {
                        ret += base64_chars[char_array_4[i]];
                    }
                    i = 0;
                }
            }

            if (i)
            {
                for (j = i; j < 3; j++)
                {
                    char_array_3[j] = '\0';
                }

                char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
                char_array_4[1] = ((char_array_3[0] & 0x03) << 4) +
                                  ((char_array_3[1] & 0xf0) >> 4);
                char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) +
                                  ((char_array_3[2] & 0xc0) >> 6);
                char_array_4[3] = char_array_3[2] & 0x3f;

                for (j = 0; (j < i + 1); j++)
                {
                    ret += base64_chars[char_array_4[j]];
                }

                while ((i++ < 3))
                {
                    ret += '=';
                }
            }

            return ret;
        }

        inline std::vector<unsigned char>
        decode(const std::string& encoded_string)
        {
            const auto& base64_chars = valid_chars();
            auto in_len = encoded_string.size();
            int i = 0;
            int j = 0;
            int in = 0;
            unsigned char char_array_4[4];
            unsigned char char_array_3[3];
            std::vector<unsigned char> ret;

            while (in_len-- && (encoded_string[in] != '=') &&
                   is_base64(encoded_string[in]))
            {
                char_array_4[i++] = encoded_string[in];
                in++;

                if (i == 4)
                {
                    for (i = 0; i < 4; i++)
                    {
                        char_array_4[i] = static_cast<unsigned char>(
                            base64_chars.find(char_array_4[i]));
                    }

                    char_array_3[0] = (char_array_4[0] << 2) +
                                      ((char_array_4[1] & 0x30) >> 4);
                    char_array_3[1] = ((char_array_4[1] & 0xf) << 4) +
                                      ((char_array_4[2] & 0x3c) >> 2);
                    char_array_3[2] =
                        ((char_array_4[2] & 0x3) << 6) + char_array_4[3];

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
                    char_array_4[j] = static_cast<unsigned char>(
                        base64_chars.find(char_array_4[j]));
                }

                char_array_3[0] =
                    (char_array_4[0] << 2) + ((char_array_4[1] & 0x30) >> 4);
                char_array_3[1] = ((char_array_4[1] & 0xf) << 4) +
                                  ((char_array_4[2] & 0x3c) >> 2);
                char_array_3[2] =
                    ((char_array_4[2] & 0x3) << 6) + char_array_4[3];

                for (j = 0; (j < i - 1); j++)
                {
                    ret.push_back(char_array_3[j]);
                }
            }

            return ret;
        }
    }

    template <typename T>
    inline bool points_within_array(T* p, T* begin, T* end)
    {
        // Not 100% sure if this works in each and every case
        return std::greater_equal<T*>()(p, begin) && std::less<T*>()(p, end);
    }

    // Forward declarations
    class http_request;
}

//! The ResponseClass attribute of a ResponseMessage
enum class response_class
{
    //! An error has occurred
    error,

    //! Everything went fine
    success,

    //! Something strange but not fatal happened
    warning
};

//! Base-class for all exceptions thrown by this library
class exception : public std::runtime_error
{
public:
    explicit exception(const std::string& what) : std::runtime_error(what) {}

    explicit exception(const char* what) : std::runtime_error(what) {}
};

//! Raised when a response from a server could not be parsed
class xml_parse_error final : public exception
{
public:
    explicit xml_parse_error(const std::string& what) : exception(what) {}
    explicit xml_parse_error(const char* what) : exception(what) {}

#ifndef EWS_DOXYGEN_SHOULD_SKIP_THIS
    static std::string error_message_from(const rapidxml::parse_error& exc,
                                          const std::vector<char>& xml)
    {
        using internal::points_within_array;

        const auto what = std::string(exc.what());
        const auto* where = exc.where<char>();

        if ((where == nullptr) || (*where == '\0'))
        {
            return what;
        }

        std::string msg = what;
        try
        {
            const auto start = &xml[0];
            if (points_within_array(where, start, (start + xml.size() + 1)))
            {
                enum {column_width = 79};
                const auto idx =
                    static_cast<std::size_t>(std::distance(start, where));

                auto doc = std::string(start, xml.size());
                auto lineno = 1U;
                auto charno = 0U;
                std::replace_if(begin(doc), end(doc),
                                [&](char c) -> bool {
                                    charno++;
                                    if (c == '\n')
                                    {
                                        if (charno < idx)
                                        {
                                            lineno++;
                                        }
                                        return true;
                                    }
                                    return false;
                                },
                                ' ');

                // 0-termini probably got replaced by xml_document::parse
                std::replace(begin(doc), end(doc), '\0', '>');

                // Chop off last '>'
                doc = std::string(doc.data(), doc.length() - 1);

                // Construct message
                msg = "in line " + std::to_string(lineno) + ":\n";
                msg += what + '\n';
                const auto pr = shorten(doc, idx, column_width);
                const auto line = pr.first;
                const auto line_index = pr.second;
                msg += line + '\n';

                auto squiggle = std::string(column_width, ' ');
                squiggle[line_index] = '~';
                squiggle = remove_trailing_whitespace(squiggle);
                msg += squiggle + '\n';
            }
        }
        catch (std::exception&)
        {
        }
        return msg;
    }
#endif // EWS_DOXYGEN_SHOULD_SKIP_THIS

private:
    static std::string remove_trailing_whitespace(const std::string& str)
    {
        static const auto whitespace = std::string(" \t");
        const auto end = str.find_last_not_of(whitespace);
        return str.substr(0, end + 1);
    }

    static std::pair<std::string, std::size_t>
    shorten(const std::string& str, std::size_t at, std::size_t columns)
    {
        at = min(at, str.length());
        if (str.length() < columns)
        {
            return std::make_pair(str, at);
        }

        const auto start =
            max(at - (columns / 2), static_cast<std::size_t>(0));
        const auto end = min(at + (columns / 2), str.length());
        EWS_ASSERT(start < end);
        std::string line;
        std::copy(&str[start], &str[end], std::back_inserter(line));
        const auto line_index = columns / 2;
        return std::make_pair(line, line_index);
    }
};

//! \brief Response code enum describes status information about a request
enum class response_code
{
    //! No error occurred
    no_error,

    //! Calling account does not have the rights to perform the action
    //! requested.
    error_access_denied,

    //! Returned when the account in question has been disabled.
    error_account_disabled,

    //! The address space (Domain Name System [DNS] domain name) record for
    //! cross forest availability could not be found in the Microsoft Active
    //! Directory
    error_address_space_not_found,

    //! Operation failed due to issues talking with the Active Directory.
    error_ad_operation,

    //! You should never encounter this response code, which occurs only as
    //! a result of an issue in Exchange Web Services.
    error_ad_session_filter,

    //! The Active Directory is temporarily unavailable. Try your request
    //! again later.
    error_ad_unavailable,

    //! Indicates that Exchange Web Services tried to determine the URL of a
    //! cross forest Client Access Server (CAS) by using the Autodiscover
    //! service, but the call to Autodiscover failed.
    error_auto_discover_failed,

    //! The AffectedTaskOccurrences enumeration value is missing. It is
    //! required when deleting a task so that Exchange Web Services knows
    //! whether you want to delete a single task or all occurrences of a
    //! repeating task.
    error_affected_task_occurrences_required,

    //! You encounter this error only if the size of your attachment exceeds
    //! Int32. MaxValue (in bytes). Of course, you probably have bigger
    //! problems if that is the case.
    error_attachment_size_limit_exceeded,

    //! The availability configuration information for the local Active
    //! Directory forest is missing from the Active Directory.
    error_availability_config_not_found,

    //! Indicates that the previous item in the request failed in such a way
    //! that Exchange Web Services stopped processing the remaining items in
    //! the request. All remaining items are marked with
    //! ErrorBatchProcessingStopped.
    error_batch_processing_stopped,

    //! You are not allowed to move or copy calendar item occurrences.
    error_calendar_cannot_move_or_copy_occurrence,

    //! If the update in question would send out a meeting update, but the
    //! meeting is in the organizer's deleted items folder, then you
    //! encounter this error.
    error_calendar_cannot_update_deleted_item,

    //! When a RecurringMasterId is examined, the OccurrenceId attribute is
    //! examined to see if it refers to a valid occurrence. If it doesn't,
    //! then this response code is returned.
    error_calendar_cannot_use_id_for_occurrence_id,

    //! When an OccurrenceId is examined, the RecurringMasterId attribute is
    //! examined to see if it refers to a valid recurring master. If it
    //! doesn't, then this response code is returned.
    error_calendar_cannot_use_id_for_recurring_master_id,

    //! Returned if the suggested duration of a calendar item exceeds five
    //! years.
    error_calendar_duration_is_too_long,

    //! The end date must be greater than the start date. Otherwise, it
    //! isn't worth having the meeting.
    error_calendar_end_date_is_earlier_than_start_date,

    //! You can apply calendar paging only on a CalendarFolder.
    error_calendar_folder_is_invalid_for_calendar_view,

    //! You should never encounter this response code.
    error_calendar_invalid_attribute_value,

    //! When defining a time change pattern, the values Day, WeekDay and
    //! WeekendDay are invalid.
    error_calendar_invalid_day_for_time_change_pattern,

    //! When defining a weekly recurring pattern, the values Day, Weekday,
    //! and WeekendDay are invalid.
    error_calendar_invalid_day_for_weekly_recurrence,

    //! Indicates that the state of the calendar item recurrence blob in the
    //! Exchange Store is invalid.
    error_calendar_invalid_property_state,

    //! You should never encounter this response code.
    error_calendar_invalid_property_value,

    //! You should never encounter this response code. It indicates that the
    //! internal structure of the objects representing the recurrence is
    //! invalid.
    error_calendar_invalid_recurrence,

    //! Occurs when an invalid time zone is encountered.
    error_calendar_invalid_time_zone,

    //! Accepting a meeting request by using delegate access is not
    //! supported in RTM.
    error_calendar_is_delegated_for_accept,

    //! Declining a meeting request by using delegate access is not
    //! supported in RTM.
    error_calendar_is_delegated_for_decline,

    //! Removing an item from the calendar by using delegate access is not
    //! supported in RTM.
    error_calendar_is_delegated_for_remove,

    //! Tentatively accepting a meeting request by using delegate access is
    //! not supported in RTM.
    error_calendar_is_delegated_for_tentative,

    //! Only the meeting organizer can cancel the meeting, no matter how
    //! much you are dreading it.
    error_calendar_is_not_organizer,

    //! The organizer cannot accept the meeting. Only attendees can accept a
    //! meeting request.
    error_calendar_is_organizer_for_accept,

    //! The organizer cannot decline the meeting. Only attendees can decline
    //! a meeting request.
    error_calendar_is_organizer_for_decline,

    //! The organizer cannot remove the meeting from the calendar by using
    //! RemoveItem. The organizer can do this only by cancelling the meeting
    //! request. Only attendees can remove a calendar item.
    error_calendar_is_organizer_for_remove,

    //! The organizer cannot tentatively accept the meeting request. Only
    //! attendees can tentatively accept a meeting request.
    error_calendar_is_organizer_for_tentative,

    //! Occurs when the occurrence index specified in the OccurenceId does
    //! not correspond to a valid occurrence. For instance, if your
    //! recurrence pattern defines a set of three meeting occurrences, and
    //! you try to access the fifth occurrence, well, there is no fifth
    //! occurrence. So instead, you get this response code.
    error_calendar_occurrence_index_is_out_of_recurrence_range,

    //! Occurs when the occurrence index specified in the OccurrenceId
    //! corresponds to a deleted occurrence.
    error_calendar_occurrence_is_deleted_from_recurrence,

    //! Occurs when a recurrence pattern is defined that contains values for
    //! month, day, week, and so on that are out of range. For example,
    //! specifying the fifteenth week of the month is both silly and an
    //! error.
    error_calendar_out_of_range,

    //! Calendar paging can span a maximum of two years.
    error_calendar_view_range_too_big,

    //! Calendar items can be created only within calendar folders.
    error_cannot_create_calendar_item_in_non_calendar_folder,

    //! Contacts can be created only within contact folders.
    error_cannot_create_contact_in_non_contacts_folder,

    //! Tasks can be created only within Task folders.
    error_cannot_create_task_in_non_task_folder,

    //! Occurs when Exchange Web Services fails to delete the item or folder
    //! in question for some unexpected reason.
    error_cannot_delete_object,

    //! This error indicates that you either tried to delete an occurrence
    //! of a nonrecurring task or tried to delete the last occurrence of a
    //! recurring task.
    error_cannot_delete_task_occurrence,

    //! Exchange Web Services could not open the file attachment
    error_cannot_open_file_attachment,

    //! The Id that was passed represents a folder rather than an item
    error_cannot_use_folder_id_for_item_id,

    //! The id that was passed in represents an item rather than a folder
    error_cannot_user_item_id_for_folder_id,

    //! You will never encounter this response code. Superseded by
    //! ErrorChangeKeyRequiredForWriteOperations.
    error_change_key_required,

    //! When performing certain update operations, you must provide a valid
    //! change key.
    error_change_key_required_for_write_operations,

    //! This response code is returned when Exchange Web Services is unable
    //! to connect to the Mailbox in question.
    error_connection_failed,

    //! Occurs when Exchange Web Services is unable to retrieve the MIME
    //! content for the item in question (GetItem), or is unable to create
    //! the item from the supplied MIME content (CreateItem).
    error_content_conversion_failed,

    //! Indicates that the data in question is corrupt and cannot be acted
    //! upon.
    error_corrupt_data,

    //! Indicates that the caller does not have the right to create the item
    //! in question.
    error_create_item_access_denied,

    //! Indicates that one ore more of the managed folders passed to
    //! CreateManagedFolder failed to be created. The only way to determine
    //! which ones failed is to search for the folders and see which ones do
    //! not exist.
    error_create_managed_folder_partial_completion,

    //! The calling account does not have the proper rights to create the
    //! subfolder in question.
    error_create_subfolder_access_denied,

    //! You cannot move an item or folder from one Mailbox to another.
    error_cross_mailbox_move_copy,

    //! Either the data that you were trying to set exceeded the maximum
    //! size for the property, or the value is large enough to require
    //! streaming and the property does not support streaming (that is,
    //! folder properties).
    error_data_size_limit_exceeded,

    //! An Active Directory operation failed.
    error_data_source_operation,

    //! You cannot delete a distinguished folder
    error_delete_distinguished_folder,

    //! You will never encounter this response code.
    error_delete_items_failed,

    //! There are duplicate values in the folder names array that was passed
    //! into CreateManagedFolder.
    error_duplicate_input_folder_names,

    //! The Mailbox subelement of DistinguishedFolderId pointed to a
    //! different Mailbox than the one you are currently operating on. For
    //! example, you cannot create a search folder that exists in one
    //! Mailbox but considers distinguished folders from another Mailbox in
    //! its search criteria.
    error_email_address_mismatch,

    //! Indicates that the subscription that was created with a particular
    //! watermark is no longer valid.
    error_event_not_found,

    //! Indicates that the subscription referenced by GetEvents has expired.
    error_expired_subscription,

    //! The folder is corrupt and cannot be saved. This means that you set
    //! some strange and invalid property on the folder, or possibly that
    //! the folder was already in some strange state before you tried to
    //! set values on it (UpdateFolder). In any case, this is not a common
    //! error.
    error_folder_corrupt,

    //! Indicates that the folder id passed in does not correspond to a
    //! valid folder, or in delegate access cases that the delegate does
    //! not have permissions to access the folder.
    error_folder_not_found,

    //! Indicates that the property that was requested could not be
    //! retrieved. Note that this does not mean that it just wasn't there.
    //! This means that the property was corrupt in some way such that
    //! retrieving it failed. This is not a common error.
    error_folder_property_request_failed,

    //! The folder could not be created or saved due to invalid state.
    error_folder_save,

    //! The folder could not be created or saved due to invalid state.
    error_folder_save_failed,

    //! The folder could not be created or updated due to invalid property
    //! values. The properties which caused the problem are listed in the
    //! MessageXml element..
    error_folder_save_property_error,

    //! A folder with that name already exists.
    error_folder_exists,

    //! Unable to retrieve Free/Busy information.This should not be common.
    error_free_busy_generation_failed,

    //! You will never encounter this response code.
    error_get_server_security_descriptor_failed,

    //! This response code is always returned within a SOAP fault. It
    //! indicates that the calling account does not have the
    //! ms-Exch-EPI-May-Impersonate right on either the user/contact they
    //! are try to impersonate or the Mailbox database containing the user
    //! Mailbox.
    error_impersonate_user_denied,

    //! This response code is always returned within a SOAP fault. It
    //! indicates that the calling account does not have the ms-Exch-EPI-
    //! Impersonation right on the CAS it is calling.
    error_impersonation_denied,

    //! There was an unexpected error trying to perform Server to Server
    //! authentication. This typically indicates that the service account
    //! running the Exchange Web Services application pool is misconfigured,
    //! that Exchange Web Services cannot talk to the Active Directory, or
    //! that a trust between Active Directory forests is not properly
    //! configured.
    error_impersonation_failed,

    //! Each change description within an UpdateItem or UpdateFolder call
    //! must list one and only one property to update.
    error_incorrect_update_property_count,

    //! Your request contained too many attendees to resolve. The default
    //! mailbox count limit is 100.
    error_individual_mailbox_limit_reached,

    //! Indicates that the Mailbox server is overloaded You should try your
    //! request again later.
    error_insufficient_resources,

    //! This response code means that the Exchange Web Services team members
    //! didn't anticipate all possible scenarios, and therefore Exchange
    //! Web Services encountered a condition that it couldn't recover from.
    error_internal_server_error,

    //! This response code is an interesting one. It essentially means that
    //! the Exchange Web Services team members didn't anticipate all
    //! possible scenarios, but the nature of the unexpected condition is
    //! such that it is likely temporary and so you should try again later.
    error_internal_server_transient_error,

    //! It is unlikely that you will encounter this response code. It means
    //! that Exchange Web Services tried to figure out what level of access
    //! the caller has on the Free/Busy information of another account, but
    //! the access that was returned didn't make sense.
    error_invalid_access_level,

    //! Indicates that the attachment was not found within the attachments
    //! collection on the item in question. You might encounter this if you
    //! have an attachment id, the attachment is deleted, and then you try
    //! to call GetAttachment on the old attachment id.
    error_invalid_attachment_id,

    //! Exchange Web Services supports only simple contains filters against
    //! the attachment table. If you try to retrieve the search parameters
    //! on an existing search folder that has a more complex attachment
    //! table restriction (called a subfilter), then Exchange Web Services
    //! has no way of rendering the XML for that filter, and it returns
    //! this response code. Note that you can still call GetFolder on this
    //! folder, just don't request the SearchParameters property.
    error_invalid_attachment_subfilter,

    //! Exchange Web Services supports only simple contains filters against
    //! the attachment table. If you try to retrieve the search parameters
    //! on an existing search folder that has a more complex attachment
    //! table restriction, then Exchange Web Services has no way of
    //! rendering the XML for that filter. This specific case means that
    //! the attachment subfilter is a contains (text) filter, but the
    //! subfilter is not referring to the attachment display name.
    error_invalid_attachment_subfilter_text_filter,

    //! You should not encounter this error, which has to do with a failure
    //! to proxy an availability request to another CAS
    error_invalid_authorization_context,

    //! Indicates that the passed in change key was invalid. Note that many
    //! methods do not require a change key to be passed. However, if you do
    //! provide one, it must be a valid, though not necessarily up-to-date,
    //! change key.
    error_invalid_change_key,

    //! Indicates that there was an internal error related to trying to
    //! resolve the caller's identity. This is not a common error.
    error_invalid_client_security_context,

    //! Occurs when you try to set the CompleteDate of a task to a date in
    //! the past. When converted to the local time of the CAS, the
    //! CompleteDate cannot be set to a value less than the local time on
    //! the CAS.
    error_invalid_complete_date,

    //! This response code can be returned with two different error
    //! messages: Unable to send cross-forest request for mailbox {mailbox
    //! identifier} because of invalid configuration. When
    //! UseServiceAccount is false, user name cannot be null or empty. You
    //! should never encounter this response code.
    error_invalid_cross_forest_credentials,

    //! An ExchangeImpersonation header was passed in but it did not contain
    //! a security identifier (SID), user principal name (UPN) or primary
    //! SMTP address. You must supply one of these identifiers and of
    //! course, they cannot be empty strings. Note that this response code
    //! is always returned within a SOAP fault.
    error_invalid_exchange_impersonation_header_data,

    //! The bitmask passed into the Excludes restriction was unparsable.
    error_invalid_excludes_restriction,

    //! You should never encounter this response code.
    error_invalid_expression_type_for_sub_filter,

    //! The combination of extended property values that were specified is
    //! invalid or results in a bad extended property URI.
    error_invalid_extended_property,

    //! The value offered for the extended property is inconsistent with the
    //! type specified in the associated extended field URI. For example, if
    //! the PropertyType on the extended field URI is set to String, but you
    //! set the value of the extended property as an array of integers, you
    //! will encounter this error.
    error_invalid_extended_property_value,

    //! You should never encounter this response code
    error_invalid_folder_id,

    //! This response code will occur if:
    //! Numerator > denominator
    //! Numerator < 0
    //! Denominator <= 0
    error_invalid_fractional_paging_parameters,

    //! Returned if you call GetUserAvailability with a FreeBusyViewType of
    //! None
    error_invalid_free_busy_view_type,

    //! Indicates that the Id and/or change key is malformed
    error_invalid_id,

    //! Occurs if you pass in an empty id (Id="")
    error_invalid_id_empty,

    //! Indicates that the Id is malformed.
    error_invalid_id_malformed,

    //! Here is an example of over-engineering. Once again, this indicates
    //! that the structure of the id is malformed The moniker referred to in
    //! the name of this response code is contained within the id and
    //! indicates which Mailbox the id belongs to. Exchange Web Services
    //! checks the length of this moniker and fails it if the byte count is
    //! more than expected. The check is good, but there is no reason to
    //! have a separate response code for this condition.
    error_invalid_id_moniker_too_long,

    //! The AttachmentId specified does not refer to an item attachment.
    error_invalid_id_not_an_item_attachment_id,

    //! You should never encounter this response code. If you do, it
    //! indicates that a contact in your Mailbox is so corrupt that Exchange
    //! Web Services really can't tell the difference between it and a
    //! waffle maker.
    error_invalid_id_returned_by_resolve_names,

    //! Treat this like ErrorInvalidIdMalformed. Indicates that the id was
    //! malformed
    error_invalid_id_store_object_id_too_long,

    //! Exchange Web Services allows for attachment hierarchies to be a
    //! maximum of 255 levels deep. If the attachment hierarchy on an item
    //! exceeds 255 levels, you will get this response code.
    error_invalid_id_too_many_attachment_levels,

    //! You will never encounter this response code.
    error_invalid_id_xml,

    //! Indicates that the offset was < 0.
    error_invalid_indexed_paging_parameters,

    //! You will never encounter this response code. At one point in the
    //! Exchange Web Services development cycle, you could update the
    //! Internet message headers via UpdateItem. Because that is no longer
    //! the case, this response code is unused.
    error_invalid_internet_header_child_nodes,

    //! Indicates that you tried to create an item attachment by using an
    //! unsupported item type. Supported item types for item attachments are
    //! Item, Message, CalendarItem, Contact, and Task. For instance, if you
    //! try to create a MeetingMessage attachment, you encounter this
    //! response code. In fact, the schema indicates that MeetingMessage,
    //! MeetingRequest, MeetingResponse, and MeetingCancellation are all
    //! permissible. Don't believe it.
    error_invalid_item_for_operation_create_item_attachment,

    //! Indicates that you tried to create an unsupported item. In addition
    //! to response objects, Supported items include Item, Message,
    //! CalendarItem, Task, and Contact. For example, you cannot use
    //! CreateItem to create a DistributionList. In addition, certain types
    //! of items are created as a side effect of doing another action.
    //! Meeting messages, for instance, are created as a result of sending a
    //! calendar item to attendees. You never explicitly create a meeting
    //! message.
    error_invalid_item_for_operation_create_item,

    //! This response code is returned if: You created an AcceptItem
    //! response object and referenced an item type other than a meeting
    //! request or a calendar item. You tried to accept a calendar item
    //! occurrence that is in the deleted items folder.
    error_invalid_item_for_operation_accept_item,

    //! You created a CancelItem response object and referenced an item type
    //! other than a calendar item.
    error_invalid_item_for_operation_cancel_item,

    //! This response code is returned if: You created a DeclineItem
    //! response object referencing an item with a type other than a
    //! meeting request or a calendar item. You tried to decline a calendar
    //! item occurrence that is in the deleted items folder.
    error_invalid_item_for_operation_decline_item,

    //! The ItemId passed to ExpandDL does not represent a distribution
    //! list. For example, you cannot expand a Message.
    error_invalid_item_for_operation_expand_dl,

    //! You created a RemoveItem response object reference an item with a
    //! type other than a meeting cancellation.
    error_invalid_item_for_operation_remove_item,

    //! You tried to send an item with a type that does not derive from
    //! MessageItem. Only items whose ItemClass begins with "IPM.Note"
    //! are sendable
    error_invalid_item_for_operation_send_item,

    //! This response code is returned if: You created a
    //! TentativelyAcceptItem response object referencing an item whose
    //! type is not a meeting request or a calendar item. You tried to
    //! tentatively accept a calendar item occurrence that is in the
    //! deleted items folder.
    error_invalid_item_for_operation_tentative,

    //! Indicates that the structure of the managed folder is corrupt and
    //! cannot be rendered.
    error_invalid_managed_folder_property,

    //! Indicates that the quota that is set on the managed folder is less
    //! than zero, indicating a corrupt managed folder.
    error_invalid_managed_folder_quota,

    //! Indicates that the size that is set on the managed folder is less
    //! than zero, indicating a corrupt managed folder.
    error_invalid_managed_folder_size,

    //! Indicates that the supplied merged free/busy internal value is
    //! invalid. Default minimum is 5 minutes. Default maximum is 1440
    //! minutes.
    error_invalid_merged_free_busy_interval,

    //! Indicates that the name passed into ResolveNames was invalid. For
    //! instance, a zero length string, a single space, a comma, and a dash
    //! are all invalid names. Vakue? Yes, that is part of the message text,
    //! although it should obviously be "value."
    error_invalid_name_for_name_resolution,

    //! Indicates that there is a problem with the NetworkService account on
    //! the CAS. This response code is quite rare and has been seen only in
    //! the wild by the most vigilant of hunters.
    error_invalid_network_service_context,

    //! You will never encounter this response code.
    error_invalid_oof_parameter,

    //! Indicates that you specified a MaxRows value that is <= 0.
    error_invalid_paging_max_rows,

    //! You tried to create a folder within a search folder.
    error_invalid_parent_folder,

    //! You tried to set the percentage complete property to an invalid
    //! value (must be between 0 and 100 inclusive).
    error_invalid_percent_complete_value,

    //! The property that you are trying to append to does not support
    //! appending. Currently, the only properties that support appending
    //! are:
    //! \li Recipient collections (ToRecipients, CcRecipients,
    //! BccRecipients)
    //! \li Attendee collections (RequiredAttendees, OptionalAttendees,
    //! Resources)
    //! \li Body
    //! \li ReplyTo
    error_invalid_property_append,

    //! The property that you are trying to delete does not support
    //! deleting. An example of this is trying to delete the ItemId of an
    //! item.
    error_invalid_property_delete,

    //! You cannot pass in a flags property to an Exists filter. The flags
    //! properties are IsDraft, IsSubmitted, IsUnmodified, IsResend,
    //! IsFromMe, and IsRead. Use IsEqualTo instead. The reason that flag
    //! don't make sense in an Exists filter is that each of these flags is
    //! actually a bit within a single property. So, calling Exists() on one
    //! of these flags is like asking if a given bit exists within a value,
    //! which is different than asking if that value exists or if the bit is
    //! set. Likely you really mean to see if the bit is set, and you should
    //! use the IsEqualTo filter expression instead.
    error_invalid_property_for_exists,

    //! Indicates that the property you are trying to manipulate does not
    //! support whatever operation your are trying to perform on it.
    error_invalid_property_for_operation,

    //! Indicates that you requested a property in the response shape, and
    //! that property is not within the schema of the item type in question.
    //! For example, requesting calendar:OptionalAttendees in the response
    //! shape of GetItem when binding to a message would result in this
    //! error.
    error_invalid_property_request,

    //! The property you are trying to set is read-only.
    error_invalid_property_set,

    //! You cannot update a Message that has already been sent.
    error_invalid_property_update_sent_message,

    //! You cannot call GetEvents or Unsubscribe on a push subscription id.
    //! To unsubscribe from a push subscription, you must respond to a push
    //! request with an unsubscribe response, or simply disconnect your Web
    //! service and wait for the push notifications to time out.
    error_invalid_pull_subscription_id,

    //! The URL provided as a callback for the push subscription has a bad
    //! format. The following conditions must be met for Exchange Web
    //! Services to accept the URL:
    //!
    //! \li String length > 0 and < 2083
    //! \li Protocol is HTTP or HTTPS
    //! \li Must be parsable by the System.Uri.NET Framework class
    error_invalid_push_subscription_url,

    //! You should never encounter this response code. If you do, the
    //! recipient collection on your message or attendee collection on your
    //! calendar item is invalid.
    error_invalid_recipients,

    //! Indicates that the search folder in question has a recipient table
    //! filter that Exchange Web Services cannot represent. The response
    //! code is a little misleading—there is nothing invalid about the
    //! search folder restriction. To get around this issue, call GetFolder
    //! without requesting the search parameters.
    error_invalid_recipient_subfilter,

    //! Indicates that the search folder in question has a recipient table
    //! filter that Exchange Web Services cannot represent. The error code
    //! is a little misleading—there is nothing invalid about the search
    //! folder restriction. To get around this, issue, call GetFolder
    //! without requesting the search parameters.
    error_invalid_recipient_subfilter_comparison,

    //! Indicates that the search folder in question has a recipient table
    //! filter that Exchange Web Services cannot represent. The response
    //! code is a little misleading—there is nothing invalid about the
    //! search folder restriction To get around this,issue, call GetFolder
    //! without requesting the search parameters.
    error_invalid_recipient_subfilter_order,

    //! Can you guess our comments on this one? Indicates that the search
    //! folder in question has a recipient table filter that Exchange Web
    //! Services cannot represent. The error code is a little
    //! misleading—there is nothing invalid about the search folder
    //! restriction. To get around this issue, call GetFolder without
    //! requesting the search parameters.
    error_invalid_recipient_subfilter_text_filter,

    //! You can only reply to/forward a Message, CalendarItem, or their
    //! descendants. If you are referencing a CalendarItem and you are the
    //! organizer, you can only forward the item. If you are referencing a
    //! draft message, you cannot reply to the item. For read receipt
    //! suppression, you can reference only a Message or descendant.
    error_invalid_reference_item,

    //! Indicates that the SOAP request has a SOAP Action header, but
    //! nothing in the SOAP body. Note that the SOAP Action header is not
    //! required because Exchange Web Services can determine the method to
    //! call from the local name of the root element in the SOAP body.
    //! However, you must supply content in the SOAP body or else there is
    //! really nothing for Exchange Web Services to act on.
    error_invalid_request,

    //! You will never encounter this response code.
    error_invalid_restriction,

    //! Indicates that the RoutingType element that was passed for an
    //! EmailAddressType is not a valid routing type. Typically, routing
    //! type will be set to Simple Mail Transfer Protocol (SMTP).
    error_invalid_routing_type,

    //! The specified duration end time must be greater than the start time
    //! and must occur in the future.
    error_invalid_scheduled_oof_duration,

    //! Indicates that the security descriptor on the calendar folder in the
    //! Store is corrupt.
    error_invalid_security_descriptor,

    //! The SaveItemToFolder attribute is false, but you included a
    //! SavedItemFolderId.
    error_invalid_send_item_save_settings,

    //! Because you never use token serialization in your application, you
    //! should never encounter this response code. The response code occurs
    //! if the token passed in the SOAP header is malformed, doesn't refer
    //! to a valid account in the Active Directory, or is missing the
    //! primary group SID.
    error_invalid_serialized_access_token,

    //! ExchangeImpersonation element have an invalid structure.
    error_invalid_sid,

    //! The passed in SMTP address is not parsable.
    error_invalid_smtp_address,

    //! You will never encounter this response code.
    error_invalid_subfilter_type,

    //! You will never encounter this response code.
    error_invalid_subfilter_type_not_attendee_type,

    //! You will never encounter this response code.
    error_invalid_subfilter_type_not_recipient_type,

    //! Indicates that the subscription is no longer valid. This could be
    //! due to the CAS having been rebooted or because the subscription has
    //! expired.
    error_invalid_subscription,

    //! Indicates that the sync state data is corrupt. You need to resync
    //! without the sync state. Make sure that if you are persisting sync
    //! state somewhere, you are not accidentally altering the information.
    error_invalid_sync_state_data,

    //! The specified time interval is invalid (schema type Duration). The
    //! start time must be greater than or equal to the end time.
    error_invalid_time_interval,

    //! The user OOF settings are invalid due to a missing internal or
    //! external reply.
    error_invalid_user_oof_settings,

    //! Indicates that the UPN passed in the ExchangeImpersonation SOAP
    //! header did not map to a valid account.
    error_invalid_user_principal_name,

    //! Indicates that the SID passed in the ExchangeImpersonation SOAP
    //! header was either invalid or did not map to a valid account.
    error_invalid_user_sid,

    //! You will never encounter this response code.
    error_invalid_user_sid_missing_upn,

    //! Indicates that the comparison value is invalid for the property your
    //! are comparing against. For instance, DateTimeCreated > "true"
    //! would cause this response code to be returned. You also encounter
    //! this response code if you specify an enumeration property in the
    //! comparison, but the value you are comparing against is not a valid
    //! value for that enumeration.
    error_invalid_value_for_property,

    //! Indicates that the supplied watermark is corrupt.
    error_invalid_watermark,

    //! Indicates that conflict resolution was unable to resolve changes for
    //! the properties in question. This typically means that someone
    //! changed the item in the Store, and you are dealing with a stale
    //! copy. Retrieve the updated change key and try again.
    error_irresolvable_conflict,

    //! Indicates that the state of the object is corrupt and cannot be
    //! retrieved. When retrieving an item, typically only certain
    //! properties will be in this state (i.e. Body, MimeContent). Try
    //! omitting those properties and retrying the operation.
    error_item_corrupt,

    //! Indicates that the item in question was not found, or potentially
    //! that it really does exist, and you just don't have rights to access
    //! it.
    error_item_not_found,

    //! Exchange Web Services tried to retrieve a given property on an item
    //! or folder but failed to do so. Note that this means that some value
    //! was there, but Exchange Web Services was unable to retrieve it.
    error_item_property_request_failed,

    //! Attempts to save the item/folder failed.
    error_item_save,

    //! Attempts to save the item/folder failed because of invalid property
    //! values. The response includes the offending property paths.
    error_item_save_property_error,

    //! You will never encounter this response code.
    error_legacy_mailbox_free_busy_view_type_not_merged,

    //! You will never encounter this response code.
    error_local_server_object_not_found,

    //! Indicates that the availability service was unable to log on as
    //! Network Service to proxy requests to the appropriate sites/forests.
    //! This typically indicates a configuration error.
    error_logon_as_network_service_failed,

    //! Indicates that the Mailbox information in the Active Directory is
    //! misconfigured.
    error_mailbox_configuration,

    //! Indicates that the MailboxData array in the request is empty. You
    //! must supply at least one Mailbox identifier.
    error_mailbox_data_array_empty,

    //! You can supply a maximum of 100 entries in the MailboxData array.
    error_mailbox_data_array_too_big,

    //! Failed to connect to the Mailbox to get the calendar view
    //! information.
    error_mailbox_logon_failed,

    //! The Mailbox in question is currently being moved. Try your request
    //! again once the move is complete.
    error_mailbox_move_in_progress,

    //! The Mailbox database is offline, corrupt, shutting down, or involved
    //! in a plethora of other conditions that make the Mailbox unavailable.
    error_mailbox_store_unavailable,

    //! Could not map the MailboxData information to a valid Mailbox
    //! account.
    error_mail_recipient_not_found,

    //! The managed folder that you are trying to create already exists in
    //! your Mailbox.
    error_managed_folder_already_exists,

    //! The folder name specified in the request does not map to a managed
    //! folder definition in the Active Directory. You can create instances
    //! of managed folders only for folders defined in the Active Directory.
    //! Check the name and try again.
    error_managed_folder_not_found,

    //! Managed folders typically exist within the root managed folders
    //! folder. The root must be properly created and functional in order to
    //! deal with managed folders through Exchange Web Services. Typically,
    //! this configuration happens transparently when you start dealing with
    //! managed folders.
    //! This response code indicates that the managed folders root was
    //! deleted from the Mailbox or that there is already a folder in the
    //! same parent folder with the name of the managed folder root. This
    //! response code also occurs if the attempt to create the root managed
    //! folder fails.
    error_managed_folders_root_failure,

    //! Indicates that the suggestions engine encountered a problem when it
    //! was trying to generate the suggestions.
    error_meeting_suggestion_generation_failed,

    //! When creating or updating an item that descends from MessageType,
    //! you must supply the MessageDisposition attribute on the request to
    //! indicate what operations should be performed on the message. Note
    //! that this attribute is not required for any other type of item.
    error_message_disposition_required,

    //! Indicates that the message you are trying to send exceeds the
    //! allowable limits.
    error_message_size_exceeded,

    //! For CreateItem, the MIME content was not valid iCalendar content For
    //! GetItem, the MIME content could not be generated.
    error_mime_content_conversion_failed,

    //! The MIME content to set is invalid.
    error_mime_content_invalid,

    //! The MIME content in the request is not a valid Base64 string.
    error_mime_content_invalid_base64_string,

    //! A required argument was missing from the request. The response
    //! message text indicates which argument to check.
    error_missing_argument,

    //! Indicates that you specified a distinguished folder id in the
    //! request, but the account that made the request does not have a
    //! Mailbox on the system. In that case, you must supply a Mailbox
    //! subelement under DistinguishedFolderId.
    error_missing_email_address,

    //! This is really the same failure as ErrorMissingEmailAddress except
    //! that ErrorMissingEmailAddressForManagedFolder is returned from
    //! CreateManagedFolder.
    error_missing_email_address_for_managed_folder,

    //! Indicates that the attendee or recipient does not have the
    //! EmailAddress element set. You must set this element when making
    //! requests. The other two elements within EmailAddressType are
    //! optional (name and routing type).
    error_missing_information_email_address,

    //! When creating a response object such as ForwardItem, you must supply
    //! a reference item id.
    error_missing_information_reference_item_id,

    //! When creating an item attachment, you must include a child element
    //! indicating the item that you want to create. This response code is
    //! returned if you omit this element.
    error_missing_item_for_create_item_attachment,

    //! The policy ids property (Property Tag: 0x6732, String) for the
    //! folder in question is missing. You should consider this a corrupt
    //! folder.
    error_missing_managed_folder_id,

    //! Indicates you tried to send an item with no recipients. Note that if
    //! you call CreateItem with a message disposition that causes the
    //! message to be sent, you get a different response code
    //! (ErrorInvalidRecipients).
    error_missing_recipients,

    //! Indicates that the move or copy operation failed. A move occurs in
    //! CreateItem when you accept a meeting request that is in the Deleted
    //! Items folder. In addition, if you decline a meeting request, cancel
    //! a calendar item, or remove a meeting from your calendar, it gets
    //! moved to the deleted items folder.
    error_move_copy_failed,

    //! You cannot move a distinguished folder.
    error_move_distinguished_folder,

    //! This is not actually an error. It should be categorized as a
    //! warning. This response code indicates that the ambiguous name that
    //! you specified matched more than one contact or distribution list.
    //! This is also the only "error" response code that includes response
    //! data (the matched names).
    error_name_resolution_multiple_results,

    //! Indicates that the effective caller does not have a Mailbox on the
    //! system. Name resolution considers both the Active Directory as well
    //! as the Store Mailbox.
    error_name_resolution_no_mailbox,

    //! The ambiguous name did not match any contacts in either the Active
    //! Directory or the Mailbox.
    error_name_resolution_no_results,

    //! There was no calendar folder for the Mailbox in question.
    error_no_calendar,

    //! You can set the FolderClass only when creating a generic folder. For
    //! typed folders (i.e. CalendarFolder and TaskFolder, the folder class
    //! is implied. Note that if you try to set the folder class to a
    //! different folder type via UpdateFolder, you get
    //! ErrorObjectTypeChanged—so don't even try it (we knew you were
    //! going there...). Exchange Web Services should enable you to create a
    //! more specialized—but consistent—folder class when creating a
    //! strongly typed folder. To get around this issue, use a generic
    //! folder type but set the folder class to the value you need.
    //! Exchange Web Services then creates the correct strongly typed
    //! folder.
    error_no_folder_class_override,

    //! Indicates that the caller does not have free/busy viewing rights on
    //! the calendar folder in question.
    error_no_free_busy_access,

    //! The SMTP address has no mailbox associated with it.
    //!
    //! This response code is returned in two cases:
    //!
    //! \li (CreateManagedFolder) You specified a Mailbox element in your
    //! request but you either omitted the EmailAddress child element or the
    //! value in the EmailAddress is empty
    //! \li The SMTP address does not map to a valid mailbox
    error_non_existent_mailbox,

    //! For requests that take an SMTP address, the address must be the
    //! primary SMTP address representing the Mailbox. Non-primary SMTP
    //! addresses are not permitted. The response includes the correct SMTP
    //! address to use. Because Exchange Web Services returns the primary
    //! SMTP address, it makes you wonder why Exchange Web Services didn't
    //! just use the proxy address you passed in… Note that this
    //! requirement may be removed in a future release.
    error_non_primary_smtp_address,

    //! Messaging Application Programming Interface (MAPI) properties in the
    //! custom range (0x8000 and greater) cannot be referenced by property
    //! tags. You must use PropertySetId or DistinguishedPropertySetId along
    //! with PropertyName or PropertyId.
    error_no_property_tag_for_custom_properties,

    //! The operation could not complete due to insufficient memory.
    error_not_enough_memory,

    //! For CreateItem, you cannot set the ItemClass so that it is
    //! inconsistent with the strongly typed item (i.e. Message or Contact).
    //! It must be consistent. For UpdateItem/Folder, you cannot change the
    //! item or folder class such that the type of the item/folder will
    //! change. You can change the item/folder class to a more derived
    //! instance of the same type (for example, IPM.Note to IPM.Note.Foo).
    //! Note that with CreateFolder, if you try to override the folder class
    //! so that it is different than the strongly typed folder element, you
    //! get an ErrorNoFolderClassOverride. Treat ErrorObjectTypeChanged and
    //! ErrorNoFolderClassOverride in the same manner.
    error_object_type_changed,

    //! Indicates that the time allotment for a given occurrence overlaps
    //! with one of its neighbors.
    error_occurrence_crossing_boundary,

    //! Indicates that the time allotment for a given occurrence is too
    //! long, which causes the occurrence to overlap with its neighbor.
    //! This response code also occurs if the length in minutes of a given
    //! occurrence is larger than Int32.MaxValue.
    error_occurrence_time_span_too_big,

    //! You will never encounter this response code.
    error_parent_folder_id_required,

    //! The parent folder id that you specified does not exist.
    error_parent_folder_not_found,

    //! You must change your password before you can access this Mailbox.
    //! This occurs when a new account has been created, and the
    //! administrator indicated that the user must change the password at
    //! first logon. You cannot change a password through Exchange Web
    //! Services. You must use a user application such as Outlook Web Access
    //! (OWA) to change your password.
    error_password_change_required,

    //! The password associated with the calling account has expired.. You
    //! need to change your password. You cannot change a password through
    //! Exchange Web Services. You must use a user application such as
    //! Outlook Web Access to change your password.
    error_password_expired,

    //! Update failed due to invalid property values. The response message
    //! includes the offending property paths.
    error_property_update,

    //! You will never encounter this response code.
    error_property_validation_failure,

    //! You will likely never encounter this response code. This response
    //! code indicates that the request that Exchange Web Services sent to
    //! another CAS when trying to fulfill a GetUserAvailability request was
    //! invalid. This response code likely indicates a configuration or
    //! rights error, or someone trying unsuccessfully to mimic an
    //! availability proxy request.
    error_proxy_request_not_allowed,

    //! The recipient passed to GetUserAvailability is located on a legacy
    //! Exchange server (prior to Exchange Server 2007). As such, Exchange
    //! Web Services needed to contact the public folder server to retrieve
    //! free/busy information for that recipient. Unfortunately, this call
    //! failed, resulting in Exchange Web Services returning a response code
    //! of ErrorPublicFolderRequestProcessingFailed.
    error_public_folder_request_processing_failed,

    //! The recipient in question is located on a legacy Exchange server
    //! (prior to Exchange -2007). As such, Exchange Web Services needed to
    //! contact the public folder server to retrieve free/busy information
    //! for that recipient. However, the organizational unit in question did
    //! not have a public folder server associated with it.
    error_public_folder_server_not_found,

    //! Restrictions can contain a maximum of 255 filter expressions. If you
    //! try to bind to an existing search folder that exceeds this limit,
    //! you encounter this response code.
    error_query_filter_too_long,

    //! The Mailbox quota has been exceeded.
    error_quota_exceeded,

    //! The process for reading events was aborted due to an internal
    //! failure. You should recreate the subscription based on a last known
    //! watermark.
    error_read_events_failed,

    //! You cannot suppress a read receipt if the message sender did not
    //! request a read receipt on the message.
    error_read_receipt_not_pending,

    //! The end date for the recurrence was out of range (it is limited to
    //! September 1, 4500).
    error_recurrence_end_date_too_big,

    //! The recurrence has no occurrence instances in the specified range.
    error_recurrence_has_no_occurrence,

    //! You will never encounter this response code.
    error_request_aborted,

    //! During GetUserAvailability processing, the request was deemed larger
    //! than it should be. You should not encounter this response code.
    error_request_stream_too_big,

    //! Indicates that one or more of the required properties is missing
    //! during a CreateAttachment call. The response XML indicates which
    //! property path was not set.
    error_required_property_missing,

    //! You will never encounter this response code. Just as a piece of
    //! trivia, the Exchange Web Services design team used this response
    //! code for debug builds to ensure that their responses were schema
    //! compliant. If Exchange Web Services expects you to send schema-
    //! compliant XML, then the least Exchange Web Services can do is be
    //! compliant in return.
    error_response_schema_validation,

    //! A restriction can have a maximum of 255 filter elements.
    error_restriction_too_long,

    //! Exchange Web Services cannot evaluate the restriction you supplied.
    //! The restriction might appear simple, but Exchange Web Services does
    //! not agree with you.
    error_restriction_too_complex,

    //! The number of calendar entries for a given recipient exceeds the
    //! allowable limit (1000). Reduce the window and try again.
    error_result_set_too_big,

    //! Indicates that the folder you want to save the item to does not
    //! exist.
    error_saved_item_folder_not_found,

    //! Exchange Web Services validates all incoming requests against the
    //! schema files (types.xsd, messages.xsd). Any instance documents that
    //! are not compliant are rejected, and this response code is returned.
    //! Note that this response code is always returned within a SOAP fault.
    error_schema_validation,

    //! Indicates that the search folder has been created, but the search
    //! criteria was never set on the folder. This condition occurs only
    //! when you access corrupt search folders that were created with
    //! another application programming interface (API) or client. Exchange
    //! Web Services does not enable you to create search folders with this
    //! condition To fix a search folder that has not been initialized,
    //! call UpdateFolder and set the SearchParameters to include the
    //! restriction that should be on the folder.
    error_search_folder_not_initialized,

    //! The caller does not have Send As rights for the account in question.
    error_send_as_denied,

    //! When you are the organizer and are deleting a calendar item, you
    //! must set the SendMeetingCancellations attribute on the DeleteItem
    //! request to indicate whether meeting cancellations will be sent to
    //! the meeting attendees. If you are using the proxy classes don't
    //! forget to set the SendMeetingCancellationsSpecified property to
    //! true.
    error_send_meeting_cancellations_required,

    //! When you are the organizer and are updating a calendar item, you
    //! must set the SendMeetingInvitationsOrCancellations attribute on the
    //! UpdateItem request. If you are using the proxy classes don't forget
    //! to set the SendMeetingInvitationsOrCancellationsSpecified attribute
    //! to true.
    error_send_meeting_invitations_or_cancellations_required,

    //! When creating a calendar item, you must set the
    //! SendMeetingInvitiations attribute on the CreateItem request. If you
    //! are using the proxy classes don't forget to set the
    //! SendMeetingInvitationsSpecified attribute to true.
    error_send_meeting_invitations_required,

    //! After the organizer sends a meeting request, that request cannot be
    //! updated. If the organizer wants to modify the meeting, you need to
    //! modify the calendar item, not the meeting request.
    error_sent_meeting_request_update,

    //! After the task initiator sends a task request, that request cannot
    //! be updated. However, you should not encounter this response code
    //! because Exchange Web Services does not support task assignment at
    //! this point.
    error_sent_task_request_update,

    //! The server is busy, potentially due to virus scan operations. It is
    //! unlikely that you will encounter this response code.
    error_server_busy,

    //! You must supply an up-to-date change key when calling the applicable
    //! methods. You either did not supply a change key, or the change key
    //! you supplied is stale. Call GetItem to retrieve an updated change
    //! key and then try your operation again.
    error_stale_object,

    //! You tried to access a subscription by using an account that did not
    //! create that subscription. Each subscription is tied to its creator.
    //! It does not matter which rights one account has on the Mailbox in
    //! question. Jane's subscriptions can only be accessed by Jane.
    error_subscription_access_denied,

    //! You can cannot create a subscription if you are not the owner or do
    //! not have owner access to the Mailbox in question.
    error_subscription_delegate_access_not_supported,

    //! The specified subscription does not exist which could mean that the
    //! subscription expired, the Exchange Web Services process was
    //! restarted, or you passed in an invalid subscription. If you
    //! encounter this response code, recreate the subscription by using
    //! the last watermark that you have.
    error_subscription_not_found,

    //! Indicates that the folder id you specified in your SyncFolderItems
    //! request does not exist.
    error_sync_folder_not_found,

    //! The time window specified is larger than the allowable limit (42 by
    //! default).
    error_time_interval_too_big,

    //! The specified destination folder does not exist
    error_to_folder_not_found,

    //! The calling account does not have the ms-Exch-EPI-TokenSerialization
    //! right on the CAS that is being called. Of course, because you are
    //! not using token serialization in your application, you should never
    //! encounter this response code. Right?
    error_token_serialization_denied,

    //! You will never encounter this response code.
    error_unable_to_get_user_oof_settings,

    //! You tried to set the Culture property to a value that is not
    //! parsable by the System.Globalization.CultureInfo class.
    error_unsupported_culture,

    //! MAPI property types Error, Null, Object and ObjectArray are
    //! unsupported.
    error_unsupported_mapi_property_type,

    //! You can retrieve or set MIME content only for a post, message, or
    //! calendar item.
    error_unsupported_mime_conversion,

    //! Indicates that the property path cannot be used within a
    //! restriction.
    error_unsupported_path_for_query,

    //! Indicates that the property path cannot be use for sorting or
    //! grouping operations.
    error_unsupported_path_for_sort_group,

    //! You should never encounter this response code.
    error_unsupported_property_definition,

    //! Exchange Web Services cannot render the existing search folder
    //! restriction. This response code does not mean that anything is wrong
    //! with the search folder restriction. You can still call FindItem on
    //! the search folder to retrieve the items in the search folder; you
    //! just can't get the actual restriction clause.
    error_unsupported_query_filter,

    //! You supplied a recurrence pattern that is not supported for tasks.
    error_unsupported_recurrence,

    //! You should never encounter this response code.
    error_unsupported_sub_filter,

    //! You should never encounter this response code. It indicates that
    //! Exchange Web Services found a property type in the Store that it
    //! cannot generate XML for.
    error_unsupported_type_for_conversion,

    //! The single property path listed in a change description must match
    //! the single property that is being set within the actual Item/Folder
    //! element.
    error_update_property_mismatch,

    //! The Exchange Store detected a virus in the message you are trying to
    //! deal with.
    error_virus_detected,

    //! The Exchange Store detected a virus in the message and deleted it.
    error_virus_message_deleted,

    //! You will never encounter this response code. This was left over from
    //! the development cycle before the Exchange Web Services team had
    //! implemented voice mail folder support. Yes, there was a time when
    //! all of this was not implemented.
    error_voice_mail_not_implemented,

    //! You will never encounter this response code. It originally meant
    //! that you intended to send your Web request from Arizona, but it
    //! actually came from Minnesota instead.*
    error_web_request_in_invalid_state,

    //! Indicates that there was a failure when Exchange Web Services was
    //! talking with unmanaged code. Of course, you cannot see the inner
    //! exception because this is a SOAP response.
    error_win32_interop_error,

    //! You will never encounter this response code.
    error_working_hours_save_failed,

    //! You will never encounter this response code.
    error_working_hours_xml_malformed
};

namespace internal
{
    inline response_code str_to_response_code(const std::string& str)
    {
        if (str == "NoError")
        {
            return response_code::no_error;
        }
        if (str == "ErrorAccessDenied")
        {
            return response_code::error_access_denied;
        }
        if (str == "ErrorAccountDisabled")
        {
            return response_code::error_account_disabled;
        }
        if (str == "ErrorAddressSpaceNotFound")
        {
            return response_code::error_address_space_not_found;
        }
        if (str == "ErrorADOperation")
        {
            return response_code::error_ad_operation;
        }
        if (str == "ErrorADSessionFilter")
        {
            return response_code::error_ad_session_filter;
        }
        if (str == "ErrorADUnavailable")
        {
            return response_code::error_ad_unavailable;
        }
        if (str == "ErrorAutoDiscoverFailed")
        {
            return response_code::error_auto_discover_failed;
        }
        if (str == "ErrorAffectedTaskOccurrencesRequired")
        {
            return response_code::error_affected_task_occurrences_required;
        }
        if (str == "ErrorAttachmentSizeLimitExceeded")
        {
            return response_code::error_attachment_size_limit_exceeded;
        }
        if (str == "ErrorAvailabilityConfigNotFound")
        {
            return response_code::error_availability_config_not_found;
        }
        if (str == "ErrorBatchProcessingStopped")
        {
            return response_code::error_batch_processing_stopped;
        }
        if (str == "ErrorCalendarCannotMoveOrCopyOccurrence")
        {
            return response_code::error_calendar_cannot_move_or_copy_occurrence;
        }
        if (str == "ErrorCalendarCannotUpdateDeletedItem")
        {
            return response_code::error_calendar_cannot_update_deleted_item;
        }
        if (str == "ErrorCalendarCannotUseIdForOccurrenceId")
        {
            return response_code::
                error_calendar_cannot_use_id_for_occurrence_id;
        }
        if (str == "ErrorCalendarCannotUseIdForRecurringMasterId")
        {
            return response_code::
                error_calendar_cannot_use_id_for_recurring_master_id;
        }
        if (str == "ErrorCalendarDurationIsTooLong")
        {
            return response_code::error_calendar_duration_is_too_long;
        }
        if (str == "ErrorCalendarEndDateIsEarlierThanStartDate")
        {
            return response_code::
                error_calendar_end_date_is_earlier_than_start_date;
        }
        if (str == "ErrorCalendarFolderIsInvalidForCalendarView")
        {
            return response_code::
                error_calendar_folder_is_invalid_for_calendar_view;
        }
        if (str == "ErrorCalendarInvalidAttributeValue")
        {
            return response_code::error_calendar_invalid_attribute_value;
        }
        if (str == "ErrorCalendarInvalidDayForTimeChangePattern")
        {
            return response_code::
                error_calendar_invalid_day_for_time_change_pattern;
        }
        if (str == "ErrorCalendarInvalidDayForWeeklyRecurrence")
        {
            return response_code::
                error_calendar_invalid_day_for_weekly_recurrence;
        }
        if (str == "ErrorCalendarInvalidPropertyState")
        {
            return response_code::error_calendar_invalid_property_state;
        }
        if (str == "ErrorCalendarInvalidPropertyValue")
        {
            return response_code::error_calendar_invalid_property_value;
        }
        if (str == "ErrorCalendarInvalidRecurrence")
        {
            return response_code::error_calendar_invalid_recurrence;
        }
        if (str == "ErrorCalendarInvalidTimeZone")
        {
            return response_code::error_calendar_invalid_time_zone;
        }
        if (str == "ErrorCalendarIsDelegatedForAccept")
        {
            return response_code::error_calendar_is_delegated_for_accept;
        }
        if (str == "ErrorCalendarIsDelegatedForDecline")
        {
            return response_code::error_calendar_is_delegated_for_decline;
        }
        if (str == "ErrorCalendarIsDelegatedForRemove")
        {
            return response_code::error_calendar_is_delegated_for_remove;
        }
        if (str == "ErrorCalendarIsDelegatedForTentative")
        {
            return response_code::error_calendar_is_delegated_for_tentative;
        }
        if (str == "ErrorCalendarIsNotOrganizer")
        {
            return response_code::error_calendar_is_not_organizer;
        }
        if (str == "ErrorCalendarIsOrganizerForAccept")
        {
            return response_code::error_calendar_is_organizer_for_accept;
        }
        if (str == "ErrorCalendarIsOrganizerForDecline")
        {
            return response_code::error_calendar_is_organizer_for_decline;
        }
        if (str == "ErrorCalendarIsOrganizerForRemove")
        {
            return response_code::error_calendar_is_organizer_for_remove;
        }
        if (str == "ErrorCalendarIsOrganizerForTentative")
        {
            return response_code::error_calendar_is_organizer_for_tentative;
        }
        if (str == "ErrorCalendarOccurrenceIndexIsOutOfRecurrenceRange")
        {
            return response_code::
                error_calendar_occurrence_index_is_out_of_recurrence_range;
        }
        if (str == "ErrorCalendarOccurrenceIsDeletedFromRecurrence")
        {
            return response_code::
                error_calendar_occurrence_is_deleted_from_recurrence;
        }
        if (str == "ErrorCalendarOutOfRange")
        {
            return response_code::error_calendar_out_of_range;
        }
        if (str == "ErrorCalendarViewRangeTooBig")
        {
            return response_code::error_calendar_view_range_too_big;
        }
        if (str == "ErrorCannotCreateCalendarItemInNonCalendarFolder")
        {
            return response_code::
                error_cannot_create_calendar_item_in_non_calendar_folder;
        }
        if (str == "ErrorCannotCreateContactInNonContactsFolder")
        {
            return response_code::
                error_cannot_create_contact_in_non_contacts_folder;
        }
        if (str == "ErrorCannotCreateTaskInNonTaskFolder")
        {
            return response_code::error_cannot_create_task_in_non_task_folder;
        }
        if (str == "ErrorCannotDeleteObject")
        {
            return response_code::error_cannot_delete_object;
        }
        if (str == "ErrorCannotDeleteTaskOccurrence")
        {
            return response_code::error_cannot_delete_task_occurrence;
        }
        if (str == "ErrorCannotOpenFileAttachment")
        {
            return response_code::error_cannot_open_file_attachment;
        }
        if (str == "ErrorCannotUseFolderIdForItemId")
        {
            return response_code::error_cannot_use_folder_id_for_item_id;
        }
        if (str == "ErrorCannotUserItemIdForFolderId")
        {
            return response_code::error_cannot_user_item_id_for_folder_id;
        }
        if (str == "ErrorChangeKeyRequired")
        {
            return response_code::error_change_key_required;
        }
        if (str == "ErrorChangeKeyRequiredForWriteOperations")
        {
            return response_code::
                error_change_key_required_for_write_operations;
        }
        if (str == "ErrorConnectionFailed")
        {
            return response_code::error_connection_failed;
        }
        if (str == "ErrorContentConversionFailed")
        {
            return response_code::error_content_conversion_failed;
        }
        if (str == "ErrorCorruptData")
        {
            return response_code::error_corrupt_data;
        }
        if (str == "ErrorCreateItemAccessDenied")
        {
            return response_code::error_create_item_access_denied;
        }
        if (str == "ErrorCreateManagedFolderPartialCompletion")
        {
            return response_code::
                error_create_managed_folder_partial_completion;
        }
        if (str == "ErrorCreateSubfolderAccessDenied")
        {
            return response_code::error_create_subfolder_access_denied;
        }
        if (str == "ErrorCrossMailboxMoveCopy")
        {
            return response_code::error_cross_mailbox_move_copy;
        }
        if (str == "ErrorDataSizeLimitExceeded")
        {
            return response_code::error_data_size_limit_exceeded;
        }
        if (str == "ErrorDataSourceOperation")
        {
            return response_code::error_data_source_operation;
        }
        if (str == "ErrorDeleteDistinguishedFolder")
        {
            return response_code::error_delete_distinguished_folder;
        }
        if (str == "ErrorDeleteItemsFailed")
        {
            return response_code::error_delete_items_failed;
        }
        if (str == "ErrorDuplicateInputFolderNames")
        {
            return response_code::error_duplicate_input_folder_names;
        }
        if (str == "ErrorEmailAddressMismatch")
        {
            return response_code::error_email_address_mismatch;
        }
        if (str == "ErrorEventNotFound")
        {
            return response_code::error_event_not_found;
        }
        if (str == "ErrorExpiredSubscription")
        {
            return response_code::error_expired_subscription;
        }
        if (str == "ErrorFolderCorrupt")
        {
            return response_code::error_folder_corrupt;
        }
        if (str == "ErrorFolderNotFound")
        {
            return response_code::error_folder_not_found;
        }
        if (str == "ErrorFolderPropertyRequestFailed")
        {
            return response_code::error_folder_property_request_failed;
        }
        if (str == "ErrorFolderSave")
        {
            return response_code::error_folder_save;
        }
        if (str == "ErrorFolderSaveFailed")
        {
            return response_code::error_folder_save_failed;
        }
        if (str == "ErrorFolderSavePropertyError")
        {
            return response_code::error_folder_save_property_error;
        }
        if (str == "ErrorFolderExists")
        {
            return response_code::error_folder_exists;
        }
        if (str == "ErrorFreeBusyGenerationFailed")
        {
            return response_code::error_free_busy_generation_failed;
        }
        if (str == "ErrorGetServerSecurityDescriptorFailed")
        {
            return response_code::error_get_server_security_descriptor_failed;
        }
        if (str == "ErrorImpersonateUserDenied")
        {
            return response_code::error_impersonate_user_denied;
        }
        if (str == "ErrorImpersonationDenied")
        {
            return response_code::error_impersonation_denied;
        }
        if (str == "ErrorImpersonationFailed")
        {
            return response_code::error_impersonation_failed;
        }
        if (str == "ErrorIncorrectUpdatePropertyCount")
        {
            return response_code::error_incorrect_update_property_count;
        }
        if (str == "ErrorIndividualMailboxLimitReached")
        {
            return response_code::error_individual_mailbox_limit_reached;
        }
        if (str == "ErrorInsufficientResources")
        {
            return response_code::error_insufficient_resources;
        }
        if (str == "ErrorInternalServerError")
        {
            return response_code::error_internal_server_error;
        }
        if (str == "ErrorInternalServerTransientError")
        {
            return response_code::error_internal_server_transient_error;
        }
        if (str == "ErrorInvalidAccessLevel")
        {
            return response_code::error_invalid_access_level;
        }
        if (str == "ErrorInvalidAttachmentId")
        {
            return response_code::error_invalid_attachment_id;
        }
        if (str == "ErrorInvalidAttachmentSubfilter")
        {
            return response_code::error_invalid_attachment_subfilter;
        }
        if (str == "ErrorInvalidAttachmentSubfilterTextFilter")
        {
            return response_code::
                error_invalid_attachment_subfilter_text_filter;
        }
        if (str == "ErrorInvalidAuthorizationContext")
        {
            return response_code::error_invalid_authorization_context;
        }
        if (str == "ErrorInvalidChangeKey")
        {
            return response_code::error_invalid_change_key;
        }
        if (str == "ErrorInvalidClientSecurityContext")
        {
            return response_code::error_invalid_client_security_context;
        }
        if (str == "ErrorInvalidCompleteDate")
        {
            return response_code::error_invalid_complete_date;
        }
        if (str == "ErrorInvalidCrossForestCredentials")
        {
            return response_code::error_invalid_cross_forest_credentials;
        }
        if (str == "ErrorInvalidExchangeImpersonationHeaderData")
        {
            return response_code::
                error_invalid_exchange_impersonation_header_data;
        }
        if (str == "ErrorInvalidExcludesRestriction")
        {
            return response_code::error_invalid_excludes_restriction;
        }
        if (str == "ErrorInvalidExpressionTypeForSubFilter")
        {
            return response_code::error_invalid_expression_type_for_sub_filter;
        }
        if (str == "ErrorInvalidExtendedProperty")
        {
            return response_code::error_invalid_extended_property;
        }
        if (str == "ErrorInvalidExtendedPropertyValue")
        {
            return response_code::error_invalid_extended_property_value;
        }
        if (str == "ErrorInvalidFolderId")
        {
            return response_code::error_invalid_folder_id;
        }
        if (str == "ErrorInvalidFractionalPagingParameters")
        {
            return response_code::error_invalid_fractional_paging_parameters;
        }
        if (str == "ErrorInvalidFreeBusyViewType")
        {
            return response_code::error_invalid_free_busy_view_type;
        }
        if (str == "ErrorInvalidId")
        {
            return response_code::error_invalid_id;
        }
        if (str == "ErrorInvalidIdEmpty")
        {
            return response_code::error_invalid_id_empty;
        }
        if (str == "ErrorInvalidIdMalformed")
        {
            return response_code::error_invalid_id_malformed;
        }
        if (str == "ErrorInvalidIdMonikerTooLong")
        {
            return response_code::error_invalid_id_moniker_too_long;
        }
        if (str == "ErrorInvalidIdNotAnItemAttachmentId")
        {
            return response_code::error_invalid_id_not_an_item_attachment_id;
        }
        if (str == "ErrorInvalidIdReturnedByResolveNames")
        {
            return response_code::error_invalid_id_returned_by_resolve_names;
        }
        if (str == "ErrorInvalidIdStoreObjectIdTooLong")
        {
            return response_code::error_invalid_id_store_object_id_too_long;
        }
        if (str == "ErrorInvalidIdTooManyAttachmentLevels")
        {
            return response_code::error_invalid_id_too_many_attachment_levels;
        }
        if (str == "ErrorInvalidIdXml")
        {
            return response_code::error_invalid_id_xml;
        }
        if (str == "ErrorInvalidIndexedPagingParameters")
        {
            return response_code::error_invalid_indexed_paging_parameters;
        }
        if (str == "ErrorInvalidInternetHeaderChildNodes")
        {
            return response_code::error_invalid_internet_header_child_nodes;
        }
        if (str == "ErrorInvalidItemForOperationCreateItemAttachment")
        {
            return response_code::
                error_invalid_item_for_operation_create_item_attachment;
        }
        if (str == "ErrorInvalidItemForOperationCreateItem")
        {
            return response_code::error_invalid_item_for_operation_create_item;
        }
        if (str == "ErrorInvalidItemForOperationAcceptItem")
        {
            return response_code::error_invalid_item_for_operation_accept_item;
        }
        if (str == "ErrorInvalidItemForOperationCancelItem")
        {
            return response_code::error_invalid_item_for_operation_cancel_item;
        }
        if (str == "ErrorInvalidItemForOperationDeclineItem")
        {
            return response_code::error_invalid_item_for_operation_decline_item;
        }
        if (str == "ErrorInvalidItemForOperationExpandDL")
        {
            return response_code::error_invalid_item_for_operation_expand_dl;
        }
        if (str == "ErrorInvalidItemForOperationRemoveItem")
        {
            return response_code::error_invalid_item_for_operation_remove_item;
        }
        if (str == "ErrorInvalidItemForOperationSendItem")
        {
            return response_code::error_invalid_item_for_operation_send_item;
        }
        if (str == "ErrorInvalidItemForOperationTentative")
        {
            return response_code::error_invalid_item_for_operation_tentative;
        }
        if (str == "ErrorInvalidManagedFolderProperty")
        {
            return response_code::error_invalid_managed_folder_property;
        }
        if (str == "ErrorInvalidManagedFolderQuota")
        {
            return response_code::error_invalid_managed_folder_quota;
        }
        if (str == "ErrorInvalidManagedFolderSize")
        {
            return response_code::error_invalid_managed_folder_size;
        }
        if (str == "ErrorInvalidMergedFreeBusyInterval")
        {
            return response_code::error_invalid_merged_free_busy_interval;
        }
        if (str == "ErrorInvalidNameForNameResolution")
        {
            return response_code::error_invalid_name_for_name_resolution;
        }
        if (str == "ErrorInvalidNetworkServiceContext")
        {
            return response_code::error_invalid_network_service_context;
        }
        if (str == "ErrorInvalidOofParameter")
        {
            return response_code::error_invalid_oof_parameter;
        }
        if (str == "ErrorInvalidPagingMaxRows")
        {
            return response_code::error_invalid_paging_max_rows;
        }
        if (str == "ErrorInvalidParentFolder")
        {
            return response_code::error_invalid_parent_folder;
        }
        if (str == "ErrorInvalidPercentCompleteValue")
        {
            return response_code::error_invalid_percent_complete_value;
        }
        if (str == "ErrorInvalidPropertyAppend")
        {
            return response_code::error_invalid_property_append;
        }
        if (str == "ErrorInvalidPropertyDelete")
        {
            return response_code::error_invalid_property_delete;
        }
        if (str == "ErrorInvalidPropertyForExists")
        {
            return response_code::error_invalid_property_for_exists;
        }
        if (str == "ErrorInvalidPropertyForOperation")
        {
            return response_code::error_invalid_property_for_operation;
        }
        if (str == "ErrorInvalidPropertyRequest")
        {
            return response_code::error_invalid_property_request;
        }
        if (str == "ErrorInvalidPropertySet")
        {
            return response_code::error_invalid_property_set;
        }
        if (str == "ErrorInvalidPropertyUpdateSentMessage")
        {
            return response_code::error_invalid_property_update_sent_message;
        }
        if (str == "ErrorInvalidPullSubscriptionId")
        {
            return response_code::error_invalid_pull_subscription_id;
        }
        if (str == "ErrorInvalidPushSubscriptionUrl")
        {
            return response_code::error_invalid_push_subscription_url;
        }
        if (str == "ErrorInvalidRecipients")
        {
            return response_code::error_invalid_recipients;
        }
        if (str == "ErrorInvalidRecipientSubfilter")
        {
            return response_code::error_invalid_recipient_subfilter;
        }
        if (str == "ErrorInvalidRecipientSubfilterComparison")
        {
            return response_code::error_invalid_recipient_subfilter_comparison;
        }
        if (str == "ErrorInvalidRecipientSubfilterOrder")
        {
            return response_code::error_invalid_recipient_subfilter_order;
        }
        if (str == "ErrorInvalidRecipientSubfilterTextFilter")
        {
            return response_code::error_invalid_recipient_subfilter_text_filter;
        }
        if (str == "ErrorInvalidReferenceItem")
        {
            return response_code::error_invalid_reference_item;
        }
        if (str == "ErrorInvalidRequest")
        {
            return response_code::error_invalid_request;
        }
        if (str == "ErrorInvalidRestriction")
        {
            return response_code::error_invalid_restriction;
        }
        if (str == "ErrorInvalidRoutingType")
        {
            return response_code::error_invalid_routing_type;
        }
        if (str == "ErrorInvalidScheduledOofDuration")
        {
            return response_code::error_invalid_scheduled_oof_duration;
        }
        if (str == "ErrorInvalidSecurityDescriptor")
        {
            return response_code::error_invalid_security_descriptor;
        }
        if (str == "ErrorInvalidSendItemSaveSettings")
        {
            return response_code::error_invalid_send_item_save_settings;
        }
        if (str == "ErrorInvalidSerializedAccessToken")
        {
            return response_code::error_invalid_serialized_access_token;
        }
        if (str == "ErrorInvalidSid")
        {
            return response_code::error_invalid_sid;
        }
        if (str == "ErrorInvalidSmtpAddress")
        {
            return response_code::error_invalid_smtp_address;
        }
        if (str == "ErrorInvalidSubfilterType")
        {
            return response_code::error_invalid_subfilter_type;
        }
        if (str == "ErrorInvalidSubfilterTypeNotAttendeeType")
        {
            return response_code::
                error_invalid_subfilter_type_not_attendee_type;
        }
        if (str == "ErrorInvalidSubfilterTypeNotRecipientType")
        {
            return response_code::
                error_invalid_subfilter_type_not_recipient_type;
        }
        if (str == "ErrorInvalidSubscription")
        {
            return response_code::error_invalid_subscription;
        }
        if (str == "ErrorInvalidSyncStateData")
        {
            return response_code::error_invalid_sync_state_data;
        }
        if (str == "ErrorInvalidTimeInterval")
        {
            return response_code::error_invalid_time_interval;
        }
        if (str == "ErrorInvalidUserOofSettings")
        {
            return response_code::error_invalid_user_oof_settings;
        }
        if (str == "ErrorInvalidUserPrincipalName")
        {
            return response_code::error_invalid_user_principal_name;
        }
        if (str == "ErrorInvalidUserSid")
        {
            return response_code::error_invalid_user_sid;
        }
        if (str == "ErrorInvalidUserSidMissingUPN")
        {
            return response_code::error_invalid_user_sid_missing_upn;
        }
        if (str == "ErrorInvalidValueForProperty")
        {
            return response_code::error_invalid_value_for_property;
        }
        if (str == "ErrorInvalidWatermark")
        {
            return response_code::error_invalid_watermark;
        }
        if (str == "ErrorIrresolvableConflict")
        {
            return response_code::error_irresolvable_conflict;
        }
        if (str == "ErrorItemCorrupt")
        {
            return response_code::error_item_corrupt;
        }
        if (str == "ErrorItemNotFound")
        {
            return response_code::error_item_not_found;
        }
        if (str == "ErrorItemPropertyRequestFailed")
        {
            return response_code::error_item_property_request_failed;
        }
        if (str == "ErrorItemSave")
        {
            return response_code::error_item_save;
        }
        if (str == "ErrorItemSavePropertyError")
        {
            return response_code::error_item_save_property_error;
        }
        if (str == "ErrorLegacyMailboxFreeBusyViewTypeNotMerged")
        {
            return response_code::
                error_legacy_mailbox_free_busy_view_type_not_merged;
        }
        if (str == "ErrorLocalServerObjectNotFound")
        {
            return response_code::error_local_server_object_not_found;
        }
        if (str == "ErrorLogonAsNetworkServiceFailed")
        {
            return response_code::error_logon_as_network_service_failed;
        }
        if (str == "ErrorMailboxConfiguration")
        {
            return response_code::error_mailbox_configuration;
        }
        if (str == "ErrorMailboxDataArrayEmpty")
        {
            return response_code::error_mailbox_data_array_empty;
        }
        if (str == "ErrorMailboxDataArrayTooBig")
        {
            return response_code::error_mailbox_data_array_too_big;
        }
        if (str == "ErrorMailboxLogonFailed")
        {
            return response_code::error_mailbox_logon_failed;
        }
        if (str == "ErrorMailboxMoveInProgress")
        {
            return response_code::error_mailbox_move_in_progress;
        }
        if (str == "ErrorMailboxStoreUnavailable")
        {
            return response_code::error_mailbox_store_unavailable;
        }
        if (str == "ErrorMailRecipientNotFound")
        {
            return response_code::error_mail_recipient_not_found;
        }
        if (str == "ErrorManagedFolderAlreadyExists")
        {
            return response_code::error_managed_folder_already_exists;
        }
        if (str == "ErrorManagedFolderNotFound")
        {
            return response_code::error_managed_folder_not_found;
        }
        if (str == "ErrorManagedFoldersRootFailure")
        {
            return response_code::error_managed_folders_root_failure;
        }
        if (str == "ErrorMeetingSuggestionGenerationFailed")
        {
            return response_code::error_meeting_suggestion_generation_failed;
        }
        if (str == "ErrorMessageDispositionRequired")
        {
            return response_code::error_message_disposition_required;
        }
        if (str == "ErrorMessageSizeExceeded")
        {
            return response_code::error_message_size_exceeded;
        }
        if (str == "ErrorMimeContentConversionFailed")
        {
            return response_code::error_mime_content_conversion_failed;
        }
        if (str == "ErrorMimeContentInvalid")
        {
            return response_code::error_mime_content_invalid;
        }
        if (str == "ErrorMimeContentInvalidBase64String")
        {
            return response_code::error_mime_content_invalid_base64_string;
        }
        if (str == "ErrorMissingArgument")
        {
            return response_code::error_missing_argument;
        }
        if (str == "ErrorMissingEmailAddress")
        {
            return response_code::error_missing_email_address;
        }
        if (str == "ErrorMissingEmailAddressForManagedFolder")
        {
            return response_code::
                error_missing_email_address_for_managed_folder;
        }
        if (str == "ErrorMissingInformationEmailAddress")
        {
            return response_code::error_missing_information_email_address;
        }
        if (str == "ErrorMissingInformationReferenceItemId")
        {
            return response_code::error_missing_information_reference_item_id;
        }
        if (str == "ErrorMissingItemForCreateItemAttachment")
        {
            return response_code::error_missing_item_for_create_item_attachment;
        }
        if (str == "ErrorMissingManagedFolderId")
        {
            return response_code::error_missing_managed_folder_id;
        }
        if (str == "ErrorMissingRecipients")
        {
            return response_code::error_missing_recipients;
        }
        if (str == "ErrorMoveCopyFailed")
        {
            return response_code::error_move_copy_failed;
        }
        if (str == "ErrorMoveDistinguishedFolder")
        {
            return response_code::error_move_distinguished_folder;
        }
        if (str == "ErrorNameResolutionMultipleResults")
        {
            return response_code::error_name_resolution_multiple_results;
        }
        if (str == "ErrorNameResolutionNoMailbox")
        {
            return response_code::error_name_resolution_no_mailbox;
        }
        if (str == "ErrorNameResolutionNoResults")
        {
            return response_code::error_name_resolution_no_results;
        }
        if (str == "ErrorNoCalendar")
        {
            return response_code::error_no_calendar;
        }
        if (str == "ErrorNoFolderClassOverride")
        {
            return response_code::error_no_folder_class_override;
        }
        if (str == "ErrorNoFreeBusyAccess")
        {
            return response_code::error_no_free_busy_access;
        }
        if (str == "ErrorNonExistentMailbox")
        {
            return response_code::error_non_existent_mailbox;
        }
        if (str == "ErrorNonPrimarySmtpAddress")
        {
            return response_code::error_non_primary_smtp_address;
        }
        if (str == "ErrorNoPropertyTagForCustomProperties")
        {
            return response_code::error_no_property_tag_for_custom_properties;
        }
        if (str == "ErrorNotEnoughMemory")
        {
            return response_code::error_not_enough_memory;
        }
        if (str == "ErrorObjectTypeChanged")
        {
            return response_code::error_object_type_changed;
        }
        if (str == "ErrorOccurrenceCrossingBoundary")
        {
            return response_code::error_occurrence_crossing_boundary;
        }
        if (str == "ErrorOccurrenceTimeSpanTooBig")
        {
            return response_code::error_occurrence_time_span_too_big;
        }
        if (str == "ErrorParentFolderIdRequired")
        {
            return response_code::error_parent_folder_id_required;
        }
        if (str == "ErrorParentFolderNotFound")
        {
            return response_code::error_parent_folder_not_found;
        }
        if (str == "ErrorPasswordChangeRequired")
        {
            return response_code::error_password_change_required;
        }
        if (str == "ErrorPasswordExpired")
        {
            return response_code::error_password_expired;
        }
        if (str == "ErrorPropertyUpdate")
        {
            return response_code::error_property_update;
        }
        if (str == "ErrorPropertyValidationFailure")
        {
            return response_code::error_property_validation_failure;
        }
        if (str == "ErrorProxyRequestNotAllowed")
        {
            return response_code::error_proxy_request_not_allowed;
        }
        if (str == "ErrorPublicFolderRequestProcessingFailed")
        {
            return response_code::error_public_folder_request_processing_failed;
        }
        if (str == "ErrorPublicFolderServerNotFound")
        {
            return response_code::error_public_folder_server_not_found;
        }
        if (str == "ErrorQueryFilterTooLong")
        {
            return response_code::error_query_filter_too_long;
        }
        if (str == "ErrorQuotaExceeded")
        {
            return response_code::error_quota_exceeded;
        }
        if (str == "ErrorReadEventsFailed")
        {
            return response_code::error_read_events_failed;
        }
        if (str == "ErrorReadReceiptNotPending")
        {
            return response_code::error_read_receipt_not_pending;
        }
        if (str == "ErrorRecurrenceEndDateTooBig")
        {
            return response_code::error_recurrence_end_date_too_big;
        }
        if (str == "ErrorRecurrenceHasNoOccurrence")
        {
            return response_code::error_recurrence_has_no_occurrence;
        }
        if (str == "ErrorRequestAborted")
        {
            return response_code::error_request_aborted;
        }
        if (str == "ErrorRequestStreamTooBig")
        {
            return response_code::error_request_stream_too_big;
        }
        if (str == "ErrorRequiredPropertyMissing")
        {
            return response_code::error_required_property_missing;
        }
        if (str == "ErrorResponseSchemaValidation")
        {
            return response_code::error_response_schema_validation;
        }
        if (str == "ErrorRestrictionTooLong")
        {
            return response_code::error_restriction_too_long;
        }
        if (str == "ErrorRestrictionTooComplex")
        {
            return response_code::error_restriction_too_complex;
        }
        if (str == "ErrorResultSetTooBig")
        {
            return response_code::error_result_set_too_big;
        }
        if (str == "ErrorSavedItemFolderNotFound")
        {
            return response_code::error_saved_item_folder_not_found;
        }
        if (str == "ErrorSchemaValidation")
        {
            return response_code::error_schema_validation;
        }
        if (str == "ErrorSearchFolderNotInitialized")
        {
            return response_code::error_search_folder_not_initialized;
        }
        if (str == "ErrorSendAsDenied")
        {
            return response_code::error_send_as_denied;
        }
        if (str == "ErrorSendMeetingCancellationsRequired")
        {
            return response_code::error_send_meeting_cancellations_required;
        }
        if (str == "ErrorSendMeetingInvitationsOrCancellationsRequired")
        {
            return response_code::
                error_send_meeting_invitations_or_cancellations_required;
        }
        if (str == "ErrorSendMeetingInvitationsRequired")
        {
            return response_code::error_send_meeting_invitations_required;
        }
        if (str == "ErrorSentMeetingRequestUpdate")
        {
            return response_code::error_sent_meeting_request_update;
        }
        if (str == "ErrorSentTaskRequestUpdate")
        {
            return response_code::error_sent_task_request_update;
        }
        if (str == "ErrorServerBusy")
        {
            return response_code::error_server_busy;
        }
        if (str == "ErrorStaleObject")
        {
            return response_code::error_stale_object;
        }
        if (str == "ErrorSubscriptionAccessDenied")
        {
            return response_code::error_subscription_access_denied;
        }
        if (str == "ErrorSubscriptionDelegateAccessNotSupported")
        {
            return response_code::
                error_subscription_delegate_access_not_supported;
        }
        if (str == "ErrorSubscriptionNotFound")
        {
            return response_code::error_subscription_not_found;
        }
        if (str == "ErrorSyncFolderNotFound")
        {
            return response_code::error_sync_folder_not_found;
        }
        if (str == "ErrorTimeIntervalTooBig")
        {
            return response_code::error_time_interval_too_big;
        }
        if (str == "ErrorToFolderNotFound")
        {
            return response_code::error_to_folder_not_found;
        }
        if (str == "ErrorTokenSerializationDenied")
        {
            return response_code::error_token_serialization_denied;
        }
        if (str == "ErrorUnableToGetUserOofSettings")
        {
            return response_code::error_unable_to_get_user_oof_settings;
        }
        if (str == "ErrorUnsupportedCulture")
        {
            return response_code::error_unsupported_culture;
        }
        if (str == "ErrorUnsupportedMapiPropertyType")
        {
            return response_code::error_unsupported_mapi_property_type;
        }
        if (str == "ErrorUnsupportedMimeConversion")
        {
            return response_code::error_unsupported_mime_conversion;
        }
        if (str == "ErrorUnsupportedPathForQuery")
        {
            return response_code::error_unsupported_path_for_query;
        }
        if (str == "ErrorUnsupportedPathForSortGroup")
        {
            return response_code::error_unsupported_path_for_sort_group;
        }
        if (str == "ErrorUnsupportedPropertyDefinition")
        {
            return response_code::error_unsupported_property_definition;
        }
        if (str == "ErrorUnsupportedQueryFilter")
        {
            return response_code::error_unsupported_query_filter;
        }
        if (str == "ErrorUnsupportedRecurrence")
        {
            return response_code::error_unsupported_recurrence;
        }
        if (str == "ErrorUnsupportedSubFilter")
        {
            return response_code::error_unsupported_sub_filter;
        }
        if (str == "ErrorUnsupportedTypeForConversion")
        {
            return response_code::error_unsupported_type_for_conversion;
        }
        if (str == "ErrorUpdatePropertyMismatch")
        {
            return response_code::error_update_property_mismatch;
        }
        if (str == "ErrorVirusDetected")
        {
            return response_code::error_virus_detected;
        }
        if (str == "ErrorVirusMessageDeleted")
        {
            return response_code::error_virus_message_deleted;
        }
        if (str == "ErrorVoiceMailNotImplemented")
        {
            return response_code::error_voice_mail_not_implemented;
        }
        if (str == "ErrorWebRequestInInvalidState")
        {
            return response_code::error_web_request_in_invalid_state;
        }
        if (str == "ErrorWin32InteropError")
        {
            return response_code::error_win32_interop_error;
        }
        if (str == "ErrorWorkingHoursSaveFailed")
        {
            return response_code::error_working_hours_save_failed;
        }
        if (str == "ErrorWorkingHoursXmlMalformed")
        {
            return response_code::error_working_hours_xml_malformed;
        }

        throw exception("Unrecognized response code: " + str);
    }

    inline std::string enum_to_str(response_code code)
    {
        switch (code)
        {
        case response_code::no_error:
            return "NoError";
        case response_code::error_access_denied:
            return "ErrorAccessDenied";
        case response_code::error_account_disabled:
            return "ErrorAccountDisabled";
        case response_code::error_address_space_not_found:
            return "ErrorAddressSpaceNotFound";
        case response_code::error_ad_operation:
            return "ErrorADOperation";
        case response_code::error_ad_session_filter:
            return "ErrorADSessionFilter";
        case response_code::error_ad_unavailable:
            return "ErrorADUnavailable";
        case response_code::error_auto_discover_failed:
            return "ErrorAutoDiscoverFailed";
        case response_code::error_affected_task_occurrences_required:
            return "ErrorAffectedTaskOccurrencesRequired";
        case response_code::error_attachment_size_limit_exceeded:
            return "ErrorAttachmentSizeLimitExceeded";
        case response_code::error_availability_config_not_found:
            return "ErrorAvailabilityConfigNotFound";
        case response_code::error_batch_processing_stopped:
            return "ErrorBatchProcessingStopped";
        case response_code::error_calendar_cannot_move_or_copy_occurrence:
            return "ErrorCalendarCannotMoveOrCopyOccurrence";
        case response_code::error_calendar_cannot_update_deleted_item:
            return "ErrorCalendarCannotUpdateDeletedItem";
        case response_code::error_calendar_cannot_use_id_for_occurrence_id:
            return "ErrorCalendarCannotUseIdForOccurrenceId";
        case response_code::
            error_calendar_cannot_use_id_for_recurring_master_id:
            return "ErrorCalendarCannotUseIdForRecurringMasterId";
        case response_code::error_calendar_duration_is_too_long:
            return "ErrorCalendarDurationIsTooLong";
        case response_code::error_calendar_end_date_is_earlier_than_start_date:
            return "ErrorCalendarEndDateIsEarlierThanStartDate";
        case response_code::error_calendar_folder_is_invalid_for_calendar_view:
            return "ErrorCalendarFolderIsInvalidForCalendarView";
        case response_code::error_calendar_invalid_attribute_value:
            return "ErrorCalendarInvalidAttributeValue";
        case response_code::error_calendar_invalid_day_for_time_change_pattern:
            return "ErrorCalendarInvalidDayForTimeChangePattern";
        case response_code::error_calendar_invalid_day_for_weekly_recurrence:
            return "ErrorCalendarInvalidDayForWeeklyRecurrence";
        case response_code::error_calendar_invalid_property_state:
            return "ErrorCalendarInvalidPropertyState";
        case response_code::error_calendar_invalid_property_value:
            return "ErrorCalendarInvalidPropertyValue";
        case response_code::error_calendar_invalid_recurrence:
            return "ErrorCalendarInvalidRecurrence";
        case response_code::error_calendar_invalid_time_zone:
            return "ErrorCalendarInvalidTimeZone";
        case response_code::error_calendar_is_delegated_for_accept:
            return "ErrorCalendarIsDelegatedForAccept";
        case response_code::error_calendar_is_delegated_for_decline:
            return "ErrorCalendarIsDelegatedForDecline";
        case response_code::error_calendar_is_delegated_for_remove:
            return "ErrorCalendarIsDelegatedForRemove";
        case response_code::error_calendar_is_delegated_for_tentative:
            return "ErrorCalendarIsDelegatedForTentative";
        case response_code::error_calendar_is_not_organizer:
            return "ErrorCalendarIsNotOrganizer";
        case response_code::error_calendar_is_organizer_for_accept:
            return "ErrorCalendarIsOrganizerForAccept";
        case response_code::error_calendar_is_organizer_for_decline:
            return "ErrorCalendarIsOrganizerForDecline";
        case response_code::error_calendar_is_organizer_for_remove:
            return "ErrorCalendarIsOrganizerForRemove";
        case response_code::error_calendar_is_organizer_for_tentative:
            return "ErrorCalendarIsOrganizerForTentative";
        case response_code::
            error_calendar_occurrence_index_is_out_of_recurrence_range:
            return "ErrorCalendarOccurrenceIndexIsOutOfRecurrenceRange";
        case response_code::
            error_calendar_occurrence_is_deleted_from_recurrence:
            return "ErrorCalendarOccurrenceIsDeletedFromRecurrence";
        case response_code::error_calendar_out_of_range:
            return "ErrorCalendarOutOfRange";
        case response_code::error_calendar_view_range_too_big:
            return "ErrorCalendarViewRangeTooBig";
        case response_code::
            error_cannot_create_calendar_item_in_non_calendar_folder:
            return "ErrorCannotCreateCalendarItemInNonCalendarFolder";
        case response_code::error_cannot_create_contact_in_non_contacts_folder:
            return "ErrorCannotCreateContactInNonContactsFolder";
        case response_code::error_cannot_create_task_in_non_task_folder:
            return "ErrorCannotCreateTaskInNonTaskFolder";
        case response_code::error_cannot_delete_object:
            return "ErrorCannotDeleteObject";
        case response_code::error_cannot_delete_task_occurrence:
            return "ErrorCannotDeleteTaskOccurrence";
        case response_code::error_cannot_open_file_attachment:
            return "ErrorCannotOpenFileAttachment";
        case response_code::error_cannot_use_folder_id_for_item_id:
            return "ErrorCannotUseFolderIdForItemId";
        case response_code::error_cannot_user_item_id_for_folder_id:
            return "ErrorCannotUserItemIdForFolderId";
        case response_code::error_change_key_required:
            return "ErrorChangeKeyRequired";
        case response_code::error_change_key_required_for_write_operations:
            return "ErrorChangeKeyRequiredForWriteOperations";
        case response_code::error_connection_failed:
            return "ErrorConnectionFailed";
        case response_code::error_content_conversion_failed:
            return "ErrorContentConversionFailed";
        case response_code::error_corrupt_data:
            return "ErrorCorruptData";
        case response_code::error_create_item_access_denied:
            return "ErrorCreateItemAccessDenied";
        case response_code::error_create_managed_folder_partial_completion:
            return "ErrorCreateManagedFolderPartialCompletion";
        case response_code::error_create_subfolder_access_denied:
            return "ErrorCreateSubfolderAccessDenied";
        case response_code::error_cross_mailbox_move_copy:
            return "ErrorCrossMailboxMoveCopy";
        case response_code::error_data_size_limit_exceeded:
            return "ErrorDataSizeLimitExceeded";
        case response_code::error_data_source_operation:
            return "ErrorDataSourceOperation";
        case response_code::error_delete_distinguished_folder:
            return "ErrorDeleteDistinguishedFolder";
        case response_code::error_delete_items_failed:
            return "ErrorDeleteItemsFailed";
        case response_code::error_duplicate_input_folder_names:
            return "ErrorDuplicateInputFolderNames";
        case response_code::error_email_address_mismatch:
            return "ErrorEmailAddressMismatch";
        case response_code::error_event_not_found:
            return "ErrorEventNotFound";
        case response_code::error_expired_subscription:
            return "ErrorExpiredSubscription";
        case response_code::error_folder_corrupt:
            return "ErrorFolderCorrupt";
        case response_code::error_folder_not_found:
            return "ErrorFolderNotFound";
        case response_code::error_folder_property_request_failed:
            return "ErrorFolderPropertyRequestFailed";
        case response_code::error_folder_save:
            return "ErrorFolderSave";
        case response_code::error_folder_save_failed:
            return "ErrorFolderSaveFailed";
        case response_code::error_folder_save_property_error:
            return "ErrorFolderSavePropertyError";
        case response_code::error_folder_exists:
            return "ErrorFolderExists";
        case response_code::error_free_busy_generation_failed:
            return "ErrorFreeBusyGenerationFailed";
        case response_code::error_get_server_security_descriptor_failed:
            return "ErrorGetServerSecurityDescriptorFailed";
        case response_code::error_impersonate_user_denied:
            return "ErrorImpersonateUserDenied";
        case response_code::error_impersonation_denied:
            return "ErrorImpersonationDenied";
        case response_code::error_impersonation_failed:
            return "ErrorImpersonationFailed";
        case response_code::error_incorrect_update_property_count:
            return "ErrorIncorrectUpdatePropertyCount";
        case response_code::error_individual_mailbox_limit_reached:
            return "ErrorIndividualMailboxLimitReached";
        case response_code::error_insufficient_resources:
            return "ErrorInsufficientResources";
        case response_code::error_internal_server_error:
            return "ErrorInternalServerError";
        case response_code::error_internal_server_transient_error:
            return "ErrorInternalServerTransientError";
        case response_code::error_invalid_access_level:
            return "ErrorInvalidAccessLevel";
        case response_code::error_invalid_attachment_id:
            return "ErrorInvalidAttachmentId";
        case response_code::error_invalid_attachment_subfilter:
            return "ErrorInvalidAttachmentSubfilter";
        case response_code::error_invalid_attachment_subfilter_text_filter:
            return "ErrorInvalidAttachmentSubfilterTextFilter";
        case response_code::error_invalid_authorization_context:
            return "ErrorInvalidAuthorizationContext";
        case response_code::error_invalid_change_key:
            return "ErrorInvalidChangeKey";
        case response_code::error_invalid_client_security_context:
            return "ErrorInvalidClientSecurityContext";
        case response_code::error_invalid_complete_date:
            return "ErrorInvalidCompleteDate";
        case response_code::error_invalid_cross_forest_credentials:
            return "ErrorInvalidCrossForestCredentials";
        case response_code::error_invalid_exchange_impersonation_header_data:
            return "ErrorInvalidExchangeImpersonationHeaderData";
        case response_code::error_invalid_excludes_restriction:
            return "ErrorInvalidExcludesRestriction";
        case response_code::error_invalid_expression_type_for_sub_filter:
            return "ErrorInvalidExpressionTypeForSubFilter";
        case response_code::error_invalid_extended_property:
            return "ErrorInvalidExtendedProperty";
        case response_code::error_invalid_extended_property_value:
            return "ErrorInvalidExtendedPropertyValue";
        case response_code::error_invalid_folder_id:
            return "ErrorInvalidFolderId";
        case response_code::error_invalid_fractional_paging_parameters:
            return "ErrorInvalidFractionalPagingParameters";
        case response_code::error_invalid_free_busy_view_type:
            return "ErrorInvalidFreeBusyViewType";
        case response_code::error_invalid_id:
            return "ErrorInvalidId";
        case response_code::error_invalid_id_empty:
            return "ErrorInvalidIdEmpty";
        case response_code::error_invalid_id_malformed:
            return "ErrorInvalidIdMalformed";
        case response_code::error_invalid_id_moniker_too_long:
            return "ErrorInvalidIdMonikerTooLong";
        case response_code::error_invalid_id_not_an_item_attachment_id:
            return "ErrorInvalidIdNotAnItemAttachmentId";
        case response_code::error_invalid_id_returned_by_resolve_names:
            return "ErrorInvalidIdReturnedByResolveNames";
        case response_code::error_invalid_id_store_object_id_too_long:
            return "ErrorInvalidIdStoreObjectIdTooLong";
        case response_code::error_invalid_id_too_many_attachment_levels:
            return "ErrorInvalidIdTooManyAttachmentLevels";
        case response_code::error_invalid_id_xml:
            return "ErrorInvalidIdXml";
        case response_code::error_invalid_indexed_paging_parameters:
            return "ErrorInvalidIndexedPagingParameters";
        case response_code::error_invalid_internet_header_child_nodes:
            return "ErrorInvalidInternetHeaderChildNodes";
        case response_code::
            error_invalid_item_for_operation_create_item_attachment:
            return "ErrorInvalidItemForOperationCreateItemAttachment";
        case response_code::error_invalid_item_for_operation_create_item:
            return "ErrorInvalidItemForOperationCreateItem";
        case response_code::error_invalid_item_for_operation_accept_item:
            return "ErrorInvalidItemForOperationAcceptItem";
        case response_code::error_invalid_item_for_operation_cancel_item:
            return "ErrorInvalidItemForOperationCancelItem";
        case response_code::error_invalid_item_for_operation_decline_item:
            return "ErrorInvalidItemForOperationDeclineItem";
        case response_code::error_invalid_item_for_operation_expand_dl:
            return "ErrorInvalidItemForOperationExpandDL";
        case response_code::error_invalid_item_for_operation_remove_item:
            return "ErrorInvalidItemForOperationRemoveItem";
        case response_code::error_invalid_item_for_operation_send_item:
            return "ErrorInvalidItemForOperationSendItem";
        case response_code::error_invalid_item_for_operation_tentative:
            return "ErrorInvalidItemForOperationTentative";
        case response_code::error_invalid_managed_folder_property:
            return "ErrorInvalidManagedFolderProperty";
        case response_code::error_invalid_managed_folder_quota:
            return "ErrorInvalidManagedFolderQuota";
        case response_code::error_invalid_managed_folder_size:
            return "ErrorInvalidManagedFolderSize";
        case response_code::error_invalid_merged_free_busy_interval:
            return "ErrorInvalidMergedFreeBusyInterval";
        case response_code::error_invalid_name_for_name_resolution:
            return "ErrorInvalidNameForNameResolution";
        case response_code::error_invalid_network_service_context:
            return "ErrorInvalidNetworkServiceContext";
        case response_code::error_invalid_oof_parameter:
            return "ErrorInvalidOofParameter";
        case response_code::error_invalid_paging_max_rows:
            return "ErrorInvalidPagingMaxRows";
        case response_code::error_invalid_parent_folder:
            return "ErrorInvalidParentFolder";
        case response_code::error_invalid_percent_complete_value:
            return "ErrorInvalidPercentCompleteValue";
        case response_code::error_invalid_property_append:
            return "ErrorInvalidPropertyAppend";
        case response_code::error_invalid_property_delete:
            return "ErrorInvalidPropertyDelete";
        case response_code::error_invalid_property_for_exists:
            return "ErrorInvalidPropertyForExists";
        case response_code::error_invalid_property_for_operation:
            return "ErrorInvalidPropertyForOperation";
        case response_code::error_invalid_property_request:
            return "ErrorInvalidPropertyRequest";
        case response_code::error_invalid_property_set:
            return "ErrorInvalidPropertySet";
        case response_code::error_invalid_property_update_sent_message:
            return "ErrorInvalidPropertyUpdateSentMessage";
        case response_code::error_invalid_pull_subscription_id:
            return "ErrorInvalidPullSubscriptionId";
        case response_code::error_invalid_push_subscription_url:
            return "ErrorInvalidPushSubscriptionUrl";
        case response_code::error_invalid_recipients:
            return "ErrorInvalidRecipients";
        case response_code::error_invalid_recipient_subfilter:
            return "ErrorInvalidRecipientSubfilter";
        case response_code::error_invalid_recipient_subfilter_comparison:
            return "ErrorInvalidRecipientSubfilterComparison";
        case response_code::error_invalid_recipient_subfilter_order:
            return "ErrorInvalidRecipientSubfilterOrder";
        case response_code::error_invalid_recipient_subfilter_text_filter:
            return "ErrorInvalidRecipientSubfilterTextFilter";
        case response_code::error_invalid_reference_item:
            return "ErrorInvalidReferenceItem";
        case response_code::error_invalid_request:
            return "ErrorInvalidRequest";
        case response_code::error_invalid_restriction:
            return "ErrorInvalidRestriction";
        case response_code::error_invalid_routing_type:
            return "ErrorInvalidRoutingType";
        case response_code::error_invalid_scheduled_oof_duration:
            return "ErrorInvalidScheduledOofDuration";
        case response_code::error_invalid_security_descriptor:
            return "ErrorInvalidSecurityDescriptor";
        case response_code::error_invalid_send_item_save_settings:
            return "ErrorInvalidSendItemSaveSettings";
        case response_code::error_invalid_serialized_access_token:
            return "ErrorInvalidSerializedAccessToken";
        case response_code::error_invalid_sid:
            return "ErrorInvalidSid";
        case response_code::error_invalid_smtp_address:
            return "ErrorInvalidSmtpAddress";
        case response_code::error_invalid_subfilter_type:
            return "ErrorInvalidSubfilterType";
        case response_code::error_invalid_subfilter_type_not_attendee_type:
            return "ErrorInvalidSubfilterTypeNotAttendeeType";
        case response_code::error_invalid_subfilter_type_not_recipient_type:
            return "ErrorInvalidSubfilterTypeNotRecipientType";
        case response_code::error_invalid_subscription:
            return "ErrorInvalidSubscription";
        case response_code::error_invalid_sync_state_data:
            return "ErrorInvalidSyncStateData";
        case response_code::error_invalid_time_interval:
            return "ErrorInvalidTimeInterval";
        case response_code::error_invalid_user_oof_settings:
            return "ErrorInvalidUserOofSettings";
        case response_code::error_invalid_user_principal_name:
            return "ErrorInvalidUserPrincipalName";
        case response_code::error_invalid_user_sid:
            return "ErrorInvalidUserSid";
        case response_code::error_invalid_user_sid_missing_upn:
            return "ErrorInvalidUserSidMissingUPN";
        case response_code::error_invalid_value_for_property:
            return "ErrorInvalidValueForProperty";
        case response_code::error_invalid_watermark:
            return "ErrorInvalidWatermark";
        case response_code::error_irresolvable_conflict:
            return "ErrorIrresolvableConflict";
        case response_code::error_item_corrupt:
            return "ErrorItemCorrupt";
        case response_code::error_item_not_found:
            return "ErrorItemNotFound";
        case response_code::error_item_property_request_failed:
            return "ErrorItemPropertyRequestFailed";
        case response_code::error_item_save:
            return "ErrorItemSave";
        case response_code::error_item_save_property_error:
            return "ErrorItemSavePropertyError";
        case response_code::error_legacy_mailbox_free_busy_view_type_not_merged:
            return "ErrorLegacyMailboxFreeBusyViewTypeNotMerged";
        case response_code::error_local_server_object_not_found:
            return "ErrorLocalServerObjectNotFound";
        case response_code::error_logon_as_network_service_failed:
            return "ErrorLogonAsNetworkServiceFailed";
        case response_code::error_mailbox_configuration:
            return "ErrorMailboxConfiguration";
        case response_code::error_mailbox_data_array_empty:
            return "ErrorMailboxDataArrayEmpty";
        case response_code::error_mailbox_data_array_too_big:
            return "ErrorMailboxDataArrayTooBig";
        case response_code::error_mailbox_logon_failed:
            return "ErrorMailboxLogonFailed";
        case response_code::error_mailbox_move_in_progress:
            return "ErrorMailboxMoveInProgress";
        case response_code::error_mailbox_store_unavailable:
            return "ErrorMailboxStoreUnavailable";
        case response_code::error_mail_recipient_not_found:
            return "ErrorMailRecipientNotFound";
        case response_code::error_managed_folder_already_exists:
            return "ErrorManagedFolderAlreadyExists";
        case response_code::error_managed_folder_not_found:
            return "ErrorManagedFolderNotFound";
        case response_code::error_managed_folders_root_failure:
            return "ErrorManagedFoldersRootFailure";
        case response_code::error_meeting_suggestion_generation_failed:
            return "ErrorMeetingSuggestionGenerationFailed";
        case response_code::error_message_disposition_required:
            return "ErrorMessageDispositionRequired";
        case response_code::error_message_size_exceeded:
            return "ErrorMessageSizeExceeded";
        case response_code::error_mime_content_conversion_failed:
            return "ErrorMimeContentConversionFailed";
        case response_code::error_mime_content_invalid:
            return "ErrorMimeContentInvalid";
        case response_code::error_mime_content_invalid_base64_string:
            return "ErrorMimeContentInvalidBase64String";
        case response_code::error_missing_argument:
            return "ErrorMissingArgument";
        case response_code::error_missing_email_address:
            return "ErrorMissingEmailAddress";
        case response_code::error_missing_email_address_for_managed_folder:
            return "ErrorMissingEmailAddressForManagedFolder";
        case response_code::error_missing_information_email_address:
            return "ErrorMissingInformationEmailAddress";
        case response_code::error_missing_information_reference_item_id:
            return "ErrorMissingInformationReferenceItemId";
        case response_code::error_missing_item_for_create_item_attachment:
            return "ErrorMissingItemForCreateItemAttachment";
        case response_code::error_missing_managed_folder_id:
            return "ErrorMissingManagedFolderId";
        case response_code::error_missing_recipients:
            return "ErrorMissingRecipients";
        case response_code::error_move_copy_failed:
            return "ErrorMoveCopyFailed";
        case response_code::error_move_distinguished_folder:
            return "ErrorMoveDistinguishedFolder";
        case response_code::error_name_resolution_multiple_results:
            return "ErrorNameResolutionMultipleResults";
        case response_code::error_name_resolution_no_mailbox:
            return "ErrorNameResolutionNoMailbox";
        case response_code::error_name_resolution_no_results:
            return "ErrorNameResolutionNoResults";
        case response_code::error_no_calendar:
            return "ErrorNoCalendar";
        case response_code::error_no_folder_class_override:
            return "ErrorNoFolderClassOverride";
        case response_code::error_no_free_busy_access:
            return "ErrorNoFreeBusyAccess";
        case response_code::error_non_existent_mailbox:
            return "ErrorNonExistentMailbox";
        case response_code::error_non_primary_smtp_address:
            return "ErrorNonPrimarySmtpAddress";
        case response_code::error_no_property_tag_for_custom_properties:
            return "ErrorNoPropertyTagForCustomProperties";
        case response_code::error_not_enough_memory:
            return "ErrorNotEnoughMemory";
        case response_code::error_object_type_changed:
            return "ErrorObjectTypeChanged";
        case response_code::error_occurrence_crossing_boundary:
            return "ErrorOccurrenceCrossingBoundary";
        case response_code::error_occurrence_time_span_too_big:
            return "ErrorOccurrenceTimeSpanTooBig";
        case response_code::error_parent_folder_id_required:
            return "ErrorParentFolderIdRequired";
        case response_code::error_parent_folder_not_found:
            return "ErrorParentFolderNotFound";
        case response_code::error_password_change_required:
            return "ErrorPasswordChangeRequired";
        case response_code::error_password_expired:
            return "ErrorPasswordExpired";
        case response_code::error_property_update:
            return "ErrorPropertyUpdate";
        case response_code::error_property_validation_failure:
            return "ErrorPropertyValidationFailure";
        case response_code::error_proxy_request_not_allowed:
            return "ErrorProxyRequestNotAllowed";
        case response_code::error_public_folder_request_processing_failed:
            return "ErrorPublicFolderRequestProcessingFailed";
        case response_code::error_public_folder_server_not_found:
            return "ErrorPublicFolderServerNotFound";
        case response_code::error_query_filter_too_long:
            return "ErrorQueryFilterTooLong";
        case response_code::error_quota_exceeded:
            return "ErrorQuotaExceeded";
        case response_code::error_read_events_failed:
            return "ErrorReadEventsFailed";
        case response_code::error_read_receipt_not_pending:
            return "ErrorReadReceiptNotPending";
        case response_code::error_recurrence_end_date_too_big:
            return "ErrorRecurrenceEndDateTooBig";
        case response_code::error_recurrence_has_no_occurrence:
            return "ErrorRecurrenceHasNoOccurrence";
        case response_code::error_request_aborted:
            return "ErrorRequestAborted";
        case response_code::error_request_stream_too_big:
            return "ErrorRequestStreamTooBig";
        case response_code::error_required_property_missing:
            return "ErrorRequiredPropertyMissing";
        case response_code::error_response_schema_validation:
            return "ErrorResponseSchemaValidation";
        case response_code::error_restriction_too_long:
            return "ErrorRestrictionTooLong";
        case response_code::error_restriction_too_complex:
            return "ErrorRestrictionTooComplex";
        case response_code::error_result_set_too_big:
            return "ErrorResultSetTooBig";
        case response_code::error_saved_item_folder_not_found:
            return "ErrorSavedItemFolderNotFound";
        case response_code::error_schema_validation:
            return "ErrorSchemaValidation";
        case response_code::error_search_folder_not_initialized:
            return "ErrorSearchFolderNotInitialized";
        case response_code::error_send_as_denied:
            return "ErrorSendAsDenied";
        case response_code::error_send_meeting_cancellations_required:
            return "ErrorSendMeetingCancellationsRequired";
        case response_code::
            error_send_meeting_invitations_or_cancellations_required:
            return "ErrorSendMeetingInvitationsOrCancellationsRequired";
        case response_code::error_send_meeting_invitations_required:
            return "ErrorSendMeetingInvitationsRequired";
        case response_code::error_sent_meeting_request_update:
            return "ErrorSentMeetingRequestUpdate";
        case response_code::error_sent_task_request_update:
            return "ErrorSentTaskRequestUpdate";
        case response_code::error_server_busy:
            return "ErrorServerBusy";
        case response_code::error_stale_object:
            return "ErrorStaleObject";
        case response_code::error_subscription_access_denied:
            return "ErrorSubscriptionAccessDenied";
        case response_code::error_subscription_delegate_access_not_supported:
            return "ErrorSubscriptionDelegateAccessNotSupported";
        case response_code::error_subscription_not_found:
            return "ErrorSubscriptionNotFound";
        case response_code::error_sync_folder_not_found:
            return "ErrorSyncFolderNotFound";
        case response_code::error_time_interval_too_big:
            return "ErrorTimeIntervalTooBig";
        case response_code::error_to_folder_not_found:
            return "ErrorToFolderNotFound";
        case response_code::error_token_serialization_denied:
            return "ErrorTokenSerializationDenied";
        case response_code::error_unable_to_get_user_oof_settings:
            return "ErrorUnableToGetUserOofSettings";
        case response_code::error_unsupported_culture:
            return "ErrorUnsupportedCulture";
        case response_code::error_unsupported_mapi_property_type:
            return "ErrorUnsupportedMapiPropertyType";
        case response_code::error_unsupported_mime_conversion:
            return "ErrorUnsupportedMimeConversion";
        case response_code::error_unsupported_path_for_query:
            return "ErrorUnsupportedPathForQuery";
        case response_code::error_unsupported_path_for_sort_group:
            return "ErrorUnsupportedPathForSortGroup";
        case response_code::error_unsupported_property_definition:
            return "ErrorUnsupportedPropertyDefinition";
        case response_code::error_unsupported_query_filter:
            return "ErrorUnsupportedQueryFilter";
        case response_code::error_unsupported_recurrence:
            return "ErrorUnsupportedRecurrence";
        case response_code::error_unsupported_sub_filter:
            return "ErrorUnsupportedSubFilter";
        case response_code::error_unsupported_type_for_conversion:
            return "ErrorUnsupportedTypeForConversion";
        case response_code::error_update_property_mismatch:
            return "ErrorUpdatePropertyMismatch";
        case response_code::error_virus_detected:
            return "ErrorVirusDetected";
        case response_code::error_virus_message_deleted:
            return "ErrorVirusMessageDeleted";
        case response_code::error_voice_mail_not_implemented:
            return "ErrorVoiceMailNotImplemented";
        case response_code::error_web_request_in_invalid_state:
            return "ErrorWebRequestInInvalidState";
        case response_code::error_win32_interop_error:
            return "ErrorWin32InteropError";
        case response_code::error_working_hours_save_failed:
            return "ErrorWorkingHoursSaveFailed";
        case response_code::error_working_hours_xml_malformed:
            return "ErrorWorkingHoursXmlMalformed";
        default:
            throw exception("Unrecognized response code");
        }
    }
}

//! \brief Defines the different values for the
//! <tt>\<RequestServerVersion\></tt> element.
enum class server_version
{
    //! Target the schema files for the initial release version of
    //! Exchange 2007
    exchange_2007,

    //! Target the schema files for Exchange 2007 Service Pack 1 (SP1),
    //! Exchange 2007 Service Pack 2 (SP2), and
    //! Exchange 2007 Service Pack 3 (SP3)
    exchange_2007_sp1,

    //! Target the schema files for Exchange 2010
    exchange_2010,

    //! Target the schema files for Exchange 2010 Service Pack 1 (SP1)
    exchange_2010_sp1,

    //! Target the schema files for Exchange 2010 Service Pack 2 (SP2)
    //! and Exchange 2010 Service Pack 3 (SP3)
    exchange_2010_sp2,

    //! Target the schema files for Exchange 2013
    exchange_2013,

    //! Target the schema files for Exchange 2013 Service Pack 1 (SP1)
    exchange_2013_sp1
};

namespace internal
{
    inline std::string enum_to_str(server_version vers)
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

//! \brief Specifies the set of properties that a GetItem or GetFolder
//! method call will return.
enum class base_shape
{
    //! Return only the item or folder ID
    id_only,

    //! Return the default set of properties
    default_shape,

    //! \brief Return (nearly) all properties.
    //!
    //! Note that some properties still need to be explicitly requested as
    //! additional properties.
    all_properties
};

namespace internal
{
    inline std::string enum_to_str(base_shape shape)
    {
        switch (shape)
        {
        case base_shape::id_only:
            return "IdOnly";
        case base_shape::default_shape:
            return "Default";
        case base_shape::all_properties:
            return "AllProperties";
        default:
            throw exception("Bad enum value");
        }
    }
}

//! \brief Describes how items are deleted from the Exchange store.
//!
//! Side note: we do not provide SoftDelete because that does not make much
//! sense from an EWS perspective.
enum class delete_type
{
    //! The item is removed immediately from the user's mailbox
    hard_delete,

    //! The item is moved to a dedicated "Trash" folder
    move_to_deleted_items
};

namespace internal
{
    inline std::string enum_to_str(delete_type d)
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
}

//! Indicates which occurrences of a recurring series should be deleted
enum class affected_task_occurrences
{
    //! Apply an operation to all occurrences in the series
    all_occurrences,

    //! Apply an operation only to the specified occurrence
    specified_occurrence_only
};

namespace internal
{
    inline std::string enum_to_str(affected_task_occurrences val)
    {
        switch (val)
        {
        case affected_task_occurrences::all_occurrences:
            return "AllOccurrences";
        case affected_task_occurrences::specified_occurrence_only:
            return "SpecifiedOccurrenceOnly";
        default:
            throw exception("Bad enum value");
        }
    }
}

//! Describes how a meeting will be canceled
enum class send_meeting_cancellations
{
    //! The calendar item is deleted without sending a cancellation message
    send_to_none,

    //! The calendar item is deleted and a cancellation message is sent to
    //! all attendees
    send_only_to_all,

    //! The calendar item is deleted and a cancellation message is sent to
    //! all attendees. A copy of the cancellation message is saved in the
    //! Sent Items folder
    send_to_all_and_save_copy
};

typedef send_meeting_cancellations send_meeting_invitations;

namespace internal
{
    inline std::string enum_to_str(send_meeting_cancellations val)
    {
        switch (val)
        {
        case send_meeting_cancellations::send_to_none:
            return "SendToNone";
        case send_meeting_cancellations::send_only_to_all:
            return "SendOnlyToAll";
        case send_meeting_cancellations::send_to_all_and_save_copy:
            return "SendToAllAndSaveCopy";
        default:
            throw exception("Bad enum value");
        }
    }
}

//! \brief The type of conflict resolution to try during an UpdateItem
//! method call.
//!
//! The default value is AutoResolve.
enum class conflict_resolution
{
    //! If there is a conflict, the update operation fails and an error is
    //! returned. The call to update_item never overwrites data that has
    //! changed underneath you!
    never_overwrite,

    //! The update operation automatically resolves any conflict (if it can,
    //! otherwise the request fails)
    auto_resolve,

    //! If there is a conflict, the update operation will overwrite
    //! information. Ignores changes that occurred underneath you; last
    //! writer wins!
    always_overwrite
};

namespace internal
{
    inline std::string enum_to_str(conflict_resolution val)
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
}

//! \brief <tt>\<CreateItem\></tt> and <tt>\<UpdateItem\></tt> methods use
//! this attribute. Only
//! applicable to email messages
enum class message_disposition
{
    //! Save the message in a specified folder or in the Drafts folder if
    //! none is given
    save_only,

    //! Send the message and do not save a copy in the sender's mailbox
    send_only,

    //! Send the message and save a copy in a specified folder or in the
    //! mailbox owner's Sent Items folder if none is given
    send_and_save_copy
};

namespace internal
{
    inline std::string enum_to_str(message_disposition val)
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
}

//! Gets the free/busy status that is associated with the event
enum class free_busy_status
{
    //! The time slot is open for other events
    free,

    //! The time slot is potentially filled
    tentative,

    //! the time slot is filled and attempts by others to schedule a
    //! meeting for this time period should be avoided by an application
    busy,

    //! The user is out-of-office and may not be able to respond to
    //! meeting invitations for new events that occur in the time slot
    out_of_office,

    //! Status is undetermined. You should not explicitly set this.
    //! However the Exchange store might return this value
    no_data
};

namespace internal
{
    inline std::string enum_to_str(free_busy_status val)
    {
        switch (val)
        {
        case free_busy_status::free:
            return "Free";
        case free_busy_status::tentative:
            return "Tentative";
        case free_busy_status::busy:
            return "Busy";
        case free_busy_status::out_of_office:
            return "OOF";
        case free_busy_status::no_data:
            return "NoData";
        default:
            throw exception("Bad enum value");
        }
    }
}

//! The CalendarItemType class represents an Exchange calendar item.
enum class calendar_item_type
{
    //! The item is not associated with a recurring calendar item; default
    //! for new calendar items
    single,

    //! The item is an occurrence of a recurring calendar item
    occurrence,

    //! The item is an exception to a recurring calendar item
    exception,

    //! The item is master for a set of recurring calendar items
    recurring_master
};

//! \brief The ResponseType element represents the type of recipient
//! response that
//! is received for a meeting.
enum class response_type
{
    //! The response type is unknown.
    unknown,
    //! Indicates an organizer response type.
    organizer,
    //! Meeting is tentative accepted.
    tentative,
    //! Meeting is accepted.
    accept,
    //! Meeting is declined.
    decline,
    //! Indicates that no response is received.
    no_response_received
};

namespace internal
{
    inline std::string enum_to_str(response_type val)
    {
        switch (val)
        {
        case response_type::unknown:
            return "Unknown";
        case response_type::organizer:
            return "Organizer";
        case response_type::tentative:
            return "Tentative";
        case response_type::accept:
            return "Accept";
        case response_type::decline:
            return "Decline";
        case response_type::no_response_received:
            return "NoResponseReceived";
        default:
            throw exception("Bad enum value");
        }
    }

    inline response_type str_to_response_type(const std::string& str)
    {
        if (str == "Unknown")
        {
            return response_type::unknown;
        }
        else if (str == "Organizer")
        {
            return response_type::organizer;
        }
        else if (str == "Tentative")
        {
            return response_type::tentative;
        }
        else if (str == "Accept")
        {
            return response_type::accept;
        }
        else if (str == "Decline")
        {
            return response_type::decline;
        }
        else if (str == "NoResponseReceived")
        {
            return response_type::no_response_received;
        }
        else
        {
            throw exception("Bad enum value");
        }
    }
}

//! \brief Well known folder names enumeration. Usually rendered to XML as
//! <tt>\<DistinguishedFolderId></tt> element.
enum class standard_folder
{
    //! The Calendar folder
    calendar,

    //! The Contacts folder
    contacts,

    //! The Deleted Items folder
    deleted_items,

    //! The Drafts folder
    drafts,

    //! The Inbox folder
    inbox,

    //! The Journal folder
    journal,

    //! The Notes folder
    notes,

    //! The Outbox folder
    outbox,

    //! The Sent Items folder
    sent_items,

    //! The Tasks folder
    tasks,

    //! The root of the message folder hierarchy
    msg_folder_root,

    //! The root of the mailbox
    root,

    //! The Junk E-mail folder
    junk_email,

    //! The Search Folders folder, also known as the Finder folder
    search_folders,

    //! The Voicemail folder
    voice_mail,

    //! Following are folders containing recoverable items:

    //! The root of the Recoverable Items folder hierarchy
    recoverable_items_root,

    //! The root of the folder hierarchy of recoverable items that have been
    //! soft-deleted from the Deleted Items folder
    recoverable_items_deletions,

    //! The root of the Recoverable Items versions folder hierarchy in the
    //! archive mailbox
    recoverable_items_versions,

    //! The root of the folder hierarchy of recoverable items that have been
    //! hard-deleted from the Deleted Items folder
    recoverable_items_purges,

    //! The root of the folder hierarchy in the archive mailbox
    archive_root,

    //! The root of the message folder hierarchy in the archive mailbox
    archive_msg_folder_root,

    //! The Deleted Items folder in the archive mailbox
    archive_deleted_items,

    //! Represents the archive Inbox folder. Caution: only versions of
    //! Exchange starting with build number 15.00.0913.09 include this
    //! folder
    archive_inbox,

    //! The root of the Recoverable Items folder hierarchy in the archive
    //! mailbox
    archive_recoverable_items_root,

    //! The root of the folder hierarchy of recoverable items that have been
    //! soft-deleted from the Deleted Items folder of the archive mailbox
    archive_recoverable_items_deletions,

    //! The root of the Recoverable Items versions folder hierarchy in the
    //! archive mailbox
    archive_recoverable_items_versions,

    //! The root of the hierarchy of recoverable items that have been
    //! hard-deleted from the Deleted Items folder of the archive mailbox
    archive_recoverable_items_purges,

    // Following are folders that came with EWS 2013 and Exchange Online:

    //! The Sync Issues folder
    sync_issues,

    //! The Conflicts folder
    conflicts,

    //! The Local Failures folder
    local_failures,

    //! Represents the Server Failures folder
    server_failures,

    //! The recipient cache folder
    recipient_cache,

    //! The quick contacts folder
    quick_contacts,

    //! The conversation history folder
    conversation_history,

    //! Represents the admin audit logs folder
    admin_audit_logs,

    //! The todo search folder
    todo_search,

    //! Represents the My Contacts folder
    my_contacts,

    //! Represents the directory folder
    directory,

    //! Represents the IM contact list folder
    im_contact_list,

    //! Represents the people connect folder
    people_connect,

    //! Represents the Favorites folder
    favorites
};

//! Contains the internal and external EWS URL when using Autodiscover.
//!
//! \sa get_exchange_web_services_url
struct autodiscover_result
{
    std::string internal_ews_url;
    std::string external_ews_url;
};

struct autodiscover_hints
{
    std::string autodiscover_url;
};

//! This enumeration indicates the sensitivity of an item.
enum class sensitivity
{
    //! The item has a normal sensitivity.
    normal,
    //! The item is personal.
    personal,
    //! The item is private.
    priv,
    //! The item is confidential.
    confidential
};

namespace internal
{
    inline std::string enum_to_str(sensitivity s)
    {
        switch (s)
        {
        case sensitivity::normal:
            return "Normal";
        case sensitivity::personal:
            return "Personal";
        case sensitivity::priv:
            return "Private";
        case sensitivity::confidential:
            return "Confidential";
        default:
            throw exception("Bad enum value");
        }
    }

    inline sensitivity str_to_sensitivity(const std::string& str)
    {
        if (str == "Normal")
        {
            return sensitivity::normal;
        }
        else if (str == "Personal")
        {
            return sensitivity::personal;
        }
        else if (str == "Private")
        {
            return sensitivity::priv;
        }
        else if (str == "Confidential")
        {
            return sensitivity::confidential;
        }
        else
        {
            throw exception("Bad enum value");
        }
    }
}

//! \brief This enumeration indicates the importance of an item.
//!
//! Valid values are Low, Normal, High.
enum class importance
{
    //! Low importance
    low,

    //! Normal importance
    normal,

    //! High importance
    high
};

namespace internal
{
    inline std::string enum_to_str(importance i)
    {
        switch (i)
        {
        case importance::low:
            return "Low";
        case importance::normal:
            return "Normal";
        case importance::high:
            return "High";
        default:
            throw exception("Bad enum value");
        }
    }

    inline importance str_to_importance(const std::string& str)
    {
        if (str == "Low")
        {
            return importance::low;
        }
        else if (str == "High")
        {
            return importance::high;
        }
        else if (str == "Normal")
        {
            return importance::normal;
        }
        else
        {
            throw exception("Bad enum value");
        }
    }
}

//! Exception thrown when a request was not successful
class exchange_error final : public exception
{
public:
    explicit exchange_error(response_code code)
        : exception(internal::enum_to_str(code)), code_(code)
    {
    }

    response_code code() const EWS_NOEXCEPT { return code_; }

private:
    response_code code_;
};

namespace internal
{
    inline std::string http_status_code_to_str(int status_code)
    {
        switch (status_code)
        {
        case 100:
            return "Continue";
        case 101:
            return "Switching Protocols";
        case 200:
            return "OK";
        case 201:
            return "Created";
        case 202:
            return "Accepted";
        case 203:
            return "Non-Authoritative Information";
        case 204:
            return "No Content";
        case 205:
            return "Reset Content";
        case 206:
            return "Partial Content";
        case 300:
            return "Multiple Choices";
        case 301:
            return "Moved Permanently";
        case 302:
            return "Found";
        case 303:
            return "See Other";
        case 304:
            return "Not Modified";
        case 305:
            return "Use Proxy";
        case 307:
            return "Temporary Redirect";
        case 400:
            return "Bad Request";
        case 401:
            return "Unauthorized";
        case 402:
            return "Payment Required";
        case 403:
            return "Forbidden";
        case 404:
            return "Not Found";
        case 405:
            return "Method Not Allowed";
        case 406:
            return "Not Acceptable";
        case 407:
            return "Proxy Authentication Required";
        case 408:
            return "Request Timeout";
        case 409:
            return "Conflict";
        case 410:
            return "Gone";
        case 411:
            return "Length Required";
        case 412:
            return "Precondition Failed";
        case 413:
            return "Request Entity Too Large";
        case 414:
            return "Request-URI Too Long";
        case 415:
            return "Unsupported Media Type";
        case 416:
            return "Requested Range Not Satisfiable";
        case 417:
            return "Expectation Failed";
        case 500:
            return "Internal Server Error";
        case 501:
            return "Not Implemented";
        case 502:
            return "Bad Gateway";
        case 503:
            return "Service Unavailable";
        case 504:
            return "Gateway Timeout";
        case 505:
            return "HTTP Version Not Supported";
        }
        return "";
    }
}

//! Exception thrown when a HTTP request was not successful
class http_error final : public exception
{
public:
    explicit http_error(long status_code)
        : exception("HTTP status code: " + std::to_string(status_code) + " (" +
                    internal::http_status_code_to_str(status_code) + ")"),
          code_(status_code)
    {
    }

    long code() const EWS_NOEXCEPT { return code_; }

private:
    long code_;
};

//! A SOAP fault occurred due to a bad request
class soap_fault : public exception
{
public:
    explicit soap_fault(const std::string& what) : exception(what) {}
    explicit soap_fault(const char* what) : exception(what) {}
};

//! \brief A SOAP fault that is raised when we sent invalid XML.
//!
//! This is an internal error and indicates a bug in this library, thus
//! should never happen.
//!
//! Note: because this exception is due to a SOAP fault (sometimes
//! recognized before any server-side XML parsing finished) any included
//! failure message is likely not localized according to any MailboxCulture
//! SOAP header.
class schema_validation_error final : public soap_fault
{
public:
    schema_validation_error(unsigned long line_number,
                            unsigned long line_position, std::string violation)
        : soap_fault("The request failed schema validation"),
          violation_(std::move(violation)), line_pos_(line_position),
          line_no_(line_number)
    {
    }

    //! Line number in request string where the error was found
    unsigned long line_number() const EWS_NOEXCEPT { return line_no_; }

    //! Column number in request string where the error was found
    unsigned long line_position() const EWS_NOEXCEPT { return line_pos_; }

    //! A more detailed explanation of what went wrong
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
        explicit curl_error(const std::string& what) : exception(what) {}
        explicit curl_error(const char* what) : exception(what) {}
    };

    // Helper function; constructs an exception with a meaningful error
    // message from the given result code for the most recent cURL API call.
    //
    // msg: A string that prepends the actual cURL error message.
    // rescode: The result code of a failed cURL operation.
    inline curl_error make_curl_error(const std::string& msg, CURLcode rescode)
    {
        auto reason = std::string(curl_easy_strerror(rescode));
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
        curl_ptr() : handle_(curl_easy_init())
        {
            if (!handle_)
            {
                throw curl_error("Could not start libcurl session");
            }
            // Needed because some EWS servers use self-signed certificates
            curl_easy_setopt(handle_, CURLOPT_SSL_VERIFYPEER, false);
        }

        ~curl_ptr() { curl_easy_cleanup(handle_); }

// Could use curl_easy_duphandle for copying
#ifdef EWS_HAS_DEFAULT_AND_DELETE
        curl_ptr(const curl_ptr&) = delete;
        curl_ptr& operator=(const curl_ptr&) = delete;
#else
    private:
        curl_ptr(const curl_ptr&);            // Never defined
        curl_ptr& operator=(const curl_ptr&); // Never defined

    public:
#endif

        curl_ptr(curl_ptr&& other) : handle_(std::move(other.handle_))
        {
            other.handle_ = nullptr;
        }

        curl_ptr& operator=(curl_ptr&& rhs)
        {
            if (&rhs != this)
            {
                handle_ = std::move(rhs.handle_);
                rhs.handle_ = nullptr;
            }
            return *this;
        }

        CURL* get() const EWS_NOEXCEPT { return handle_; }

    private:
        CURL* handle_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<curl_ptr>::value, "");
    static_assert(!std::is_copy_constructible<curl_ptr>::value, "");
    static_assert(!std::is_copy_assignable<curl_ptr>::value, "");
    static_assert(std::is_move_constructible<curl_ptr>::value, "");
    static_assert(std::is_move_assignable<curl_ptr>::value, "");
#endif

    // RAII wrapper class around cURLs slist construct.
    class curl_string_list final
    {
    public:
        curl_string_list() EWS_NOEXCEPT : slist_(nullptr) {}

        ~curl_string_list() { curl_slist_free_all(slist_); }

#ifdef EWS_HAS_DEFAULT_AND_DELETE
        curl_string_list(const curl_string_list&) = delete;
        curl_string_list& operator=(const curl_string_list&) = delete;
#else
    private:
        curl_string_list(const curl_string_list&);            // N/A
        curl_string_list& operator=(const curl_string_list&); // N/A

    public:
#endif

        curl_string_list(curl_string_list&& other)
            : slist_(std::move(other.slist_))
        {
            other.slist_ = nullptr;
        }

        curl_string_list& operator=(curl_string_list&& rhs)
        {
            if (&rhs != this)
            {
                slist_ = std::move(rhs.slist_);
                rhs.slist_ = nullptr;
            }
            return *this;
        }

        void append(const char* str) EWS_NOEXCEPT
        {
            slist_ = curl_slist_append(slist_, str);
        }

        curl_slist* get() const EWS_NOEXCEPT { return slist_; }

    private:
        curl_slist* slist_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<curl_string_list>::value, "");
    static_assert(!std::is_copy_constructible<curl_string_list>::value, "");
    static_assert(!std::is_copy_assignable<curl_string_list>::value, "");
    static_assert(std::is_move_constructible<curl_string_list>::value, "");
    static_assert(std::is_move_assignable<curl_string_list>::value, "");
#endif

    // String constants
    // TODO: sure this can't be done easier within a header file?
    // We need better handling of static strings (URIs, XML node names,
    // values and stuff, see usage of rapidxml::internal::compare)
    template <typename Ch = char> struct uri
    {
        struct microsoft
        {
            enum
            {
                errors_size = 58U
            };
            static const Ch* errors()
            {
                static const Ch* val = "http://schemas.microsoft.com/"
                                       "exchange/services/2006/errors";
                return val;
            }
            enum
            {
                types_size = 57U
            };
            static const Ch* types()
            {
                static const Ch* const val = "http://schemas.microsoft.com/"
                                             "exchange/services/2006/types";
                return val;
            }
            enum
            {
                messages_size = 60U
            };
            static const Ch* messages()
            {
                static const Ch* const val = "http://schemas.microsoft.com/"
                                             "exchange/services/2006/"
                                             "messages";
                return val;
            }
            enum
            {
                autodiscover_size = 79U
            };
            static const Ch* autodiscover()
            {
                static const Ch* const val = "http://schemas.microsoft.com/"
                                             "exchange/autodiscover/"
                                             "outlook/responseschema/2006a";
                return val;
            }
        };

        struct soapxml
        {
            enum
            {
                envelope_size = 41U
            };
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
            : data_(std::move(data)), code_(code)
        {
            EWS_ASSERT(!data_.empty());
        }

#ifdef EWS_HAS_DEFAULT_AND_DELETE
        ~http_response() = default;

        http_response(const http_response&) = delete;
        http_response& operator=(const http_response&) = delete;
#else
        http_response() {}

    private:
        http_response(const http_response&);            // Never defined
        http_response& operator=(const http_response&); // Never defined

    public:
#endif

        http_response(http_response&& other)
            : data_(std::move(other.data_)), code_(std::move(other.code_))
        {
            other.code_ = 0U;
        }

        http_response& operator=(http_response&& rhs)
        {
            if (&rhs != this)
            {
                data_ = std::move(rhs.data_);
                code_ = std::move(rhs.code_);
            }
            return *this;
        }

        // Returns a reference to the raw byte content in this HTTP
        // response.
        std::vector<char>& content() EWS_NOEXCEPT { return data_; }

        // Returns a reference to the raw byte content in this HTTP
        // response.
        const std::vector<char>& content() const EWS_NOEXCEPT { return data_; }

        // Returns the response code of the HTTP request.
        long code() const EWS_NOEXCEPT { return code_; }

        // Returns whether the response is a SOAP fault.
        //
        // This means the server responded with status code 500 and
        // indicates that the entire request failed (not just a normal EWS
        // error). This can happen e.g. when the request we sent was not
        // schema compliant.
        bool is_soap_fault() const EWS_NOEXCEPT { return code() == 500U; }

        // Returns whether the HTTP response code is 200 (OK).
        bool ok() const EWS_NOEXCEPT { return code() == 200U; }

    private:
        std::vector<char> data_;
        long code_;
    };

    // Loads the XML content from a given HTTP response into a
    // new xml_document and returns it.
    //
    // Note: we are using a mutable temporary buffer internally because
    // we are using RapidXml in destructive mode (the parser modifies
    // source text during the parsing process). Hence, we need to make
    // sure that parsing is done only once!
    inline std::unique_ptr<rapidxml::xml_document<char>>
    parse_response(http_response&& response)
    {
        if (response.content().empty())
        {
            throw xml_parse_error("Cannot parse empty response");
        }

#ifdef EWS_HAS_MAKE_UNIQUE
        auto doc = std::make_unique<rapidxml::xml_document<char>>();
#else
        auto doc = std::unique_ptr<rapidxml::xml_document<char>>(
            new rapidxml::xml_document<char>());
#endif
        try
        {
            static const int flags = 0;
            doc->parse<flags>(&response.content()[0]);
        }
        catch (rapidxml::parse_error& exc)
        {
            // Swallow and erase type
            const auto msg =
                xml_parse_error::error_message_from(exc, response.content());
            throw xml_parse_error(msg);
        }

#ifdef EWS_ENABLE_VERBOSE
        std::cerr << "Response code: " << response.code() << ", Content:\n\'"
                  << *doc << "\'" << std::endl;
#endif

        return doc;
    }

    // TODO: explicitly for nodes in Types XML namespace, document or
    // change interface
    inline rapidxml::xml_node<>& create_node(rapidxml::xml_node<>& parent,
                                             const std::string& name)
    {
        auto doc = parent.document();
        EWS_ASSERT(doc && "parent node needs to be part of a document");
        auto ptr_to_qname = doc->allocate_string(name.c_str());
        auto new_node = doc->allocate_node(rapidxml::node_element);
        new_node->qname(ptr_to_qname, name.length(), ptr_to_qname + 2);
        new_node->namespace_uri(internal::uri<>::microsoft::types(),
                                internal::uri<>::microsoft::types_size);
        parent.append_node(new_node);
        return *new_node;
    }

    inline rapidxml::xml_node<>& create_node(rapidxml::xml_node<>& parent,
                                             const std::string& name,
                                             const std::string& value)
    {
        auto doc = parent.document();
        EWS_ASSERT(doc && "parent node needs to be part of a document");
        auto str = doc->allocate_string(value.c_str());
        auto& new_node = create_node(parent, name);
        new_node.value(str);
        return new_node;
    }

    // Traverse elements, depth first, beginning with given node.
    //
    // Applies given function to every element during traversal, stopping as
    // soon as that function returns true.
    template <typename Function>
    inline bool traverse_elements(const rapidxml::xml_node<>& node,
                                  Function func) EWS_NOEXCEPT
    {
        for (auto child = node.first_node(); child != nullptr;
             child = child->next_sibling())
        {
            if (traverse_elements(*child, func))
            {
                return true;
            }

            if (child->type() == rapidxml::node_element)
            {
                if (func(*child))
                {
                    return true;
                }
            }
        }
        return false;
    }

    // Select element by qualified name, nullptr if there is no such element
    // TODO: how to avoid ambiguities when two elements exist with the same
    // name and namespace but different paths
    inline rapidxml::xml_node<>*
    get_element_by_qname(const rapidxml::xml_node<>& node,
                         const char* local_name,
                         const char* namespace_uri) EWS_NOEXCEPT
    {
        EWS_ASSERT(local_name && namespace_uri);

        rapidxml::xml_node<>* element = nullptr;
        const auto local_name_len = std::strlen(local_name);
        const auto namespace_uri_len = std::strlen(namespace_uri);
        traverse_elements(node, [&](rapidxml::xml_node<>& elem) -> bool {
            using rapidxml::internal::compare;

            if (compare(elem.namespace_uri(), elem.namespace_uri_size(),
                        namespace_uri, namespace_uri_len) &&
                compare(elem.local_name(), elem.local_name_size(), local_name,
                        local_name_len))
            {
                element = std::addressof(elem);
                return true;
            }
            return false;
        });
        return element;
    }

    class credentials
    {
    public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
        virtual ~credentials() = default;
        credentials& operator=(const credentials&) = default;
#else
        virtual ~credentials() {}
#endif
        virtual void certify(http_request*) const = 0;
    };
}

//! \brief This class allows HTTP basic authentication.
//!
//! Basic authentication allows a client application to authenticate with
//! username and password. \b Important: Because the password is transmitted
//! to the server in plain-text, this method is \b not secure unless you
//! provide some form of transport layer security.
//!
//! However, basic authentication can be the correct choice for your
//! application in some circumstances, e.g., for debugging purposes, if your
//! application communicates via TLS encrypted HTTP.
class basic_credentials final : public internal::credentials
{
public:
    basic_credentials(std::string username, std::string password)
        : username_(std::move(username)), password_(std::move(password))
    {
    }

private:
    // Implemented below
    void certify(internal::http_request*) const override;

    std::string username_;
    std::string password_;
};

//! \brief This class allows NTLM authentication.
//!
//! NTLM authentication is only available for Exchange on-premises servers.
//!
//! For applications that run inside the corporate firewall, NTLM
//! authentication provides a built-in means to authenticate against an
//! Exchange server. However, because NTLM requires the client to store the
//! user's password in plain-text for the entire session, it is not the
//! safest method to use.
class ntlm_credentials final : public internal::credentials
{
public:
    ntlm_credentials(std::string username, std::string password,
                     std::string domain)
        : username_(std::move(username)), password_(std::move(password)),
          domain_(std::move(domain))
    {
    }

private:
    // Implemented below
    void certify(internal::http_request*) const override;

    std::string username_;
    std::string password_;
    std::string domain_;
};

namespace internal
{
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

        // Set this HTTP request's content length.
        void set_content_length(std::size_t content_length)
        {
            const std::string str =
                "Content-Length: " + std::to_string(content_length);
            headers_.append(str.c_str());
        }

        // Set credentials for authentication.
        void set_credentials(const credentials& creds) { creds.certify(this); }

        void set_timeout(std::chrono::seconds timeout)
        {
            curl_easy_setopt(handle_.get(), CURLOPT_TIMEOUT, timeout.count());
        }

#ifdef EWS_HAS_VARIADIC_TEMPLATES
        // Small wrapper around curl_easy_setopt(3).
        //
        // Converts return codes to exceptions of type curl_error and allows
        // objects other than http_web_requests to set transfer options
        // without direct access to the internal CURL handle.
        template <typename... Args>
        void set_option(CURLoption option, Args... args)
        {
            auto retcode = curl_easy_setopt(handle_.get(), option,
                                            std::forward<Args>(args)...);
            switch (retcode)
            {
            case CURLE_OK:
                return;

            case CURLE_FAILED_INIT:
            {
                throw make_curl_error("curl_easy_setopt: unsupported option",
                                      retcode);
            }

            default:
            {
                throw make_curl_error("curl_easy_setopt: failed setting option",
                                      retcode);
            }
            };
        }
#else
        template <typename T1> void set_option(CURLoption option, T1 arg1)
        {
            auto retcode = curl_easy_setopt(handle_.get(), option, arg1);
            switch (retcode)
            {
            case CURLE_OK:
                return;

            case CURLE_FAILED_INIT:
            {
                throw make_curl_error("curl_easy_setopt: unsupported option",
                                      retcode);
            }

            default:
            {
                throw make_curl_error("curl_easy_setopt: failed setting option",
                                      retcode);
            }
            };
        }
#if 0 // KAM - curl_easy_setopt does not take 4 args

        template <typename T1, typename T2>
        void set_option(CURLoption option, T1 arg1, T2 arg2)
        {
            auto retcode = curl_easy_setopt(handle_.get(), option, arg1, arg2);
            switch (retcode)
            {
            case CURLE_OK:
                return;

            case CURLE_FAILED_INIT:
            {
                throw make_curl_error("curl_easy_setopt: unsupported option",
                                      retcode);
            }

            default:
            {
                throw make_curl_error("curl_easy_setopt: failed setting option",
                                      retcode);
            }
            };
        }
#endif
#endif

        // Perform the HTTP request and returns the response. This function
        // blocks until the complete response is received or a timeout is
        // reached. Throws curl_error if operation could not be completed.
        //
        // request: The complete request string; you must make sure that
        // the data is encoded the way you want the server to receive it.
        http_response send(const std::string& request)
        {
            auto callback = [](char* ptr, std::size_t size, std::size_t nmemb,
                               void* userdata) -> std::size_t {
                std::vector<char>* buf =
                    reinterpret_cast<std::vector<char>*>(userdata);
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
#endif

            // Set complete request string for HTTP POST method; note: no
            // encoding here
            set_option(CURLOPT_POSTFIELDS, request.c_str());
            set_option(CURLOPT_POSTFIELDSIZE, request.length());

            // Finally, set HTTP headers. We do this as last action here
            // because we want to overwrite implicitly set header lines due
            // to the options set above with our own header lines
            set_option(CURLOPT_HTTPHEADER, headers_.get());

            std::vector<char> response_data;
            set_option(CURLOPT_WRITEFUNCTION,
                       static_cast<std::size_t (*)(
                           char*, std::size_t, std::size_t, void*)>(callback));
            set_option(CURLOPT_WRITEDATA, std::addressof(response_data));

#ifdef EWS_DISABLE_TLS_CERT_VERIFICATION
            // Turn-off verification of the server's authenticity
            set_option(CURLOPT_SSL_VERIFYPEER, 0L);
            set_option(CURLOPT_SSL_VERIFYHOST, 0L);

// Warn in release builds
#ifdef NDEBUG
#ifdef _MSC_VER
#pragma message(                                                               \
    "warning: TLS verification of the server's authenticity is disabled")
#else
#warning "TLS verification of the server's authenticity is disabled"
#endif
#endif
#endif

            auto retcode = curl_easy_perform(handle_.get());
            if (retcode != 0)
            {
                throw make_curl_error("curl_easy_perform", retcode);
            }
            long response_code = 0U;
            curl_easy_getinfo(handle_.get(), CURLINFO_RESPONSE_CODE,
                              &response_code);
            response_data.emplace_back('\0');
            return http_response(std::move(response_code),
                                 std::move(response_data));
        }

    private:
        curl_ptr handle_;
        curl_string_list headers_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(!std::is_default_constructible<http_request>::value, "");
    static_assert(!std::is_copy_constructible<http_request>::value, "");
    static_assert(!std::is_copy_assignable<http_request>::value, "");
    static_assert(std::is_move_constructible<http_request>::value, "");
    static_assert(std::is_move_assignable<http_request>::value, "");
#endif

#ifdef EWS_HAS_DEFAULT_TEMPLATE_ARGS_FOR_FUNCTIONS
    template <typename RequestHandler = http_request>
#else
    template <typename RequestHandler>
#endif
    inline http_response
    make_raw_soap_request(RequestHandler& handler, const std::string& soap_body,
                          const std::vector<std::string>& soap_headers)
    {
        std::stringstream request_stream;
        request_stream
            << "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
               "<soap:Envelope "
               "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" "
               "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" "
               "xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" "
               "xmlns:m=\"http://schemas.microsoft.com/exchange/services/"
               "2006/messages\" "
               "xmlns:t=\"http://schemas.microsoft.com/exchange/services/"
               "2006/types\">";

        if (!soap_headers.empty())
        {
            request_stream << "<soap:Header>";
            for (const auto& header : soap_headers)
            {
                request_stream << header;
            }
            request_stream << "</soap:Header>";
        }

        request_stream << "<soap:Body>";
        request_stream << soap_body;
        request_stream << "</soap:Body>";
        request_stream << "</soap:Envelope>";

#ifdef EWS_ENABLE_VERBOSE
        std::cerr << request_stream.str() << std::endl;
#endif
        return handler.send(request_stream.str());
    }
// Makes a raw SOAP request.
//
// url: The URL of the server to talk to.
// username: The user-name of user.
// password: The user's secret password, plain-text.
// domain: The user's Windows domain.
// soap_body: The contents of the SOAP body (minus the body element);
// this is the actual EWS request.
// soap_headers: Any SOAP headers to add.
//
// Returns the response.
#ifdef EWS_HAS_DEFAULT_TEMPLATE_ARGS_FOR_FUNCTIONS
    template <typename RequestHandler = http_request>
#else
    template <typename RequestHandler>
#endif
    inline http_response
    make_raw_soap_request(const std::string& url, const std::string& username,
                          const std::string& password,
                          const std::string& domain,
                          const std::string& soap_body,
                          const std::vector<std::string>& soap_headers)
    {
        RequestHandler handler(url);
        handler.set_method(RequestHandler::method::POST);
        handler.set_content_type("text/xml; charset=utf-8");
        ntlm_credentials creds(username, password, domain);
        handler.set_credentials(creds);
        return make_raw_soap_request(handler, soap_body, soap_headers);
    }

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
    // A default constructed xml_subtree instance makes only sense when an
    // item class is default constructed. In that case the buffer (and the
    // DOM) is initially empty and elements are added directly to the
    // document's root node.
    class xml_subtree final
    {
    public:
        // Default constructible because item class (and it's descendants)
        // need to be
        xml_subtree() : rawdata_(), doc_(new rapidxml::xml_document<char>) {}

        explicit xml_subtree(const rapidxml::xml_node<char>& origin,
                             std::size_t size_hint = 0U)
            : rawdata_(), doc_(new rapidxml::xml_document<char>)
        {
            reparse(origin, size_hint);
        }

        // Needs to be copy- and copy-assignable because item classes are.
        // However xml_document isn't (and can't be without major rewrite
        // IMHO).  Hence, copying is quite costly as it involves serializing
        // DOM tree to string and re-parsing into new tree
        xml_subtree& operator=(const xml_subtree& rhs)
        {
            if (&rhs != this)
            {
                // FIXME: not strong exception safe
                rawdata_.clear();
                doc_->clear();
                // TODO: can be done with deep_copy
                reparse(*rhs.doc_, rhs.rawdata_.size());
            }
            return *this;
        }

        xml_subtree(const xml_subtree& other)
            : rawdata_(), doc_(new rapidxml::xml_document<char>)
        {
            reparse(*other.doc_, other.rawdata_.size());
        }

        // Returns a pointer to the root node of this sub-tree. Returned
        // pointer can be null
        rapidxml::xml_node<>* root() EWS_NOEXCEPT { return doc_->first_node(); }

        // Returns a pointer to the root node of this sub-tree. Returned
        // pointer can be null
        const rapidxml::xml_node<>* root() const EWS_NOEXCEPT
        {
            return doc_->first_node();
        }

        // Might return nullptr when there is no such element. Client code
        // needs to check returned pointer. Should never throw
        rapidxml::xml_node<char>*
        get_node(const char* node_name) const EWS_NOEXCEPT
        {
            return get_element_by_qname(*doc_, node_name,
                                        uri<>::microsoft::types());
        }

        rapidxml::xml_node<char>*
        get_node(const std::string& node_name) const EWS_NOEXCEPT
        {
            return get_node(node_name.c_str());
        }

        rapidxml::xml_document<char>* document() const EWS_NOEXCEPT
        {
            return doc_.get();
        }

        std::string get_value_as_string(const char* node_name) const
        {
            const auto node = get_node(node_name);
            return node ? std::string(node->value(), node->value_size()) : "";
        }

        void set_or_update(const std::string& node_name,
                           const std::string& node_value)
        {
            using rapidxml::internal::compare;

            auto oldnode = get_element_by_qname(*doc_, node_name.c_str(),
                                                uri<>::microsoft::types());
            if (oldnode && compare(node_value.c_str(), node_value.length(),
                                   oldnode->value(), oldnode->value_size()))
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
            newnode->qname(ptr_to_qname, node_qname.length(), ptr_to_qname + 2);
            newnode->value(ptr_to_value);
            newnode->namespace_uri(uri<>::microsoft::types(),
                                   uri<>::microsoft::types_size);
            if (oldnode)
            {
                auto parent = oldnode->parent();
                parent->insert_node(oldnode, newnode);
                parent->remove_node(oldnode);
            }
            else
            {
                doc_->append_node(newnode);
            }
        }

        std::string to_string() const
        {
            std::string str;
            rapidxml::print(std::back_inserter(str), *doc_,
                            rapidxml::print_no_indenting);
            return str;
        }

        void append_to(rapidxml::xml_node<>& dest) const
        {
            auto target_document = dest.document();
            auto source = doc_->first_node();
            if (source)
            {
                auto new_child = deep_copy(target_document, source);
                dest.append_node(new_child);
            }
        }

    private:
        std::vector<char> rawdata_;
        std::unique_ptr<rapidxml::xml_document<char>> doc_;

        // Custom namespace processor for parsing XML sub-tree. Makes
        // uri::microsoft::types the default namespace.
        struct custom_namespace_processor
            : public rapidxml::internal::xml_namespace_processor<char>
        {
            struct scope : public rapidxml::internal::xml_namespace_processor<
                               char>::scope
            {
                explicit scope(xml_namespace_processor& processor)
                    : rapidxml::internal::xml_namespace_processor<char>::scope(
                          processor)
                {
                    static auto nsattr = make_namespace_attribute();
                    add_namespace_prefix(&nsattr);
                }

                static rapidxml::xml_attribute<char> make_namespace_attribute()
                {
                    static const char* const name = "t";
                    static const char* const value = uri<>::microsoft::types();
                    rapidxml::xml_attribute<char> attr;
                    attr.name(name, 1);
                    attr.value(value, uri<>::microsoft::types_size);
                    return attr;
                }
            };
        };

        void reparse(const rapidxml::xml_node<char>& source,
                     std::size_t size_hint)
        {
            rawdata_.reserve(size_hint);
            rapidxml::print(std::back_inserter(rawdata_), source,
                            rapidxml::print_no_indenting);
            rawdata_.emplace_back('\0');

            // Note: we use xml_document::parse_ns here. This is because we
            // substitute the default namespace processor with a custom one.
            // Reason we do this: when we re-parse only a sub-tree of the
            // original XML document, we loose all enclosing namespace URIs.

            doc_->parse_ns<0, custom_namespace_processor>(&rawdata_[0]);
        }

        static rapidxml::xml_node<>*
        deep_copy(rapidxml::xml_document<>* target_doc,
                  const rapidxml::xml_node<>* source)
        {
            EWS_ASSERT(target_doc);
            EWS_ASSERT(source);

            auto newnode = target_doc->allocate_node(source->type());

            // Copy name and value
            auto name = target_doc->allocate_string(source->name());
            newnode->qname(name, source->name_size(), source->local_name());

            auto value = target_doc->allocate_string(source->value());
            newnode->value(value, source->value_size());

            // Copy child nodes and attributes
            for (auto child = source->first_node(); child;
                 child = child->next_sibling())
            {
                newnode->append_node(deep_copy(target_doc, child));
            }

            for (auto attr = source->first_attribute(); attr;
                 attr = attr->next_attribute())
            {
                newnode->append_attribute(target_doc->allocate_attribute(
                    attr->name(), attr->value(), attr->name_size(),
                    attr->value_size()));
            }

            return newnode;
        }
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<xml_subtree>::value, "");
    static_assert(std::is_copy_constructible<xml_subtree>::value, "");
    static_assert(std::is_copy_assignable<xml_subtree>::value, "");
    static_assert(std::is_move_constructible<xml_subtree>::value, "");
    static_assert(std::is_move_assignable<xml_subtree>::value, "");
#endif

#ifdef EWS_HAS_DEFAULT_TEMPLATE_ARGS_FOR_FUNCTIONS
    template <typename RequestHandler = http_request>
#else
    template <typename RequestHandler>
#endif
    inline autodiscover_result
    get_exchange_web_services_url(const std::string& user_smtp_address,
                                  const basic_credentials& credentials,
                                  unsigned int redirections,
                                  const autodiscover_hints& hints)
    {
        using rapidxml::internal::compare;

        // Check redirection counter. We don't want to get in an endless
        // loop
        if (redirections > 2)
        {
            throw exception("Maximum of two redirections reached");
        }

        // Check SMTP address
        if (user_smtp_address.empty())
        {
            throw exception("Empty SMTP address given");
        }

        std::string autodiscover_url;
        if (hints.autodiscover_url.empty())
        {
            // Get user name and domain part from the SMTP address
            const auto at_sign_idx = user_smtp_address.find_first_of("@");
            if (at_sign_idx == std::string::npos)
            {
                throw exception("No valid SMTP address given");
            }

            const auto username = user_smtp_address.substr(0, at_sign_idx);
            const auto domain = user_smtp_address.substr(
                at_sign_idx + 1, user_smtp_address.size());

            // It is important that we use an HTTPS end-point here because we
            // authenticate with HTTP basic auth; specifically we send the
            // passphrase in plain-text
            autodiscover_url =
                "https://" + domain + "/autodiscover/autodiscover.xml";
        }
        else
        {
            autodiscover_url = hints.autodiscover_url;
        }
        // Create an Outlook provider <Autodiscover/> request
        std::stringstream sstr;
        sstr << "<?xml version=\"1.0\" encoding=\"utf-8\" ?>"
             << "<Autodiscover "
             << "xmlns=\"http://schemas.microsoft.com/exchange/"
                "autodiscover/outlook/requestschema/2006\">"
             << "<Request>"
             << "<EMailAddress>" << user_smtp_address << "</EMailAddress>"
             << "<AcceptableResponseSchema>"
             << internal::uri<>::microsoft::autodiscover()
             << "</AcceptableResponseSchema>"
             << "</Request>"
             << "</Autodiscover>";
        const auto request_string = sstr.str();

        RequestHandler handler(autodiscover_url);
        handler.set_method(RequestHandler::method::POST);
        handler.set_credentials(credentials);
        handler.set_content_type("text/xml; charset=utf-8");
        handler.set_content_length(request_string.size());

#ifdef EWS_ENABLE_VERBOSE
        std::cerr << request_string << std::endl;
#endif

        auto response = handler.send(request_string);
        if (!response.ok())
        {
            throw http_error(response.code());
        }

        const auto doc = parse_response(std::move(response));

        const auto account_node = get_element_by_qname(
            *doc, "Account", internal::uri<>::microsoft::autodiscover());
        if (!account_node)
        {
            // Check for <Error/> element
            const auto error_node = get_element_by_qname(
                *doc, "Error", "http://schemas.microsoft.com/exchange/"
                               "autodiscover/responseschema/2006");
            if (error_node)
            {
                std::string error_code;
                std::string message;

                for (auto node = error_node->first_node(); node;
                     node = node->next_sibling())
                {
                    if (compare(node->local_name(), node->local_name_size(),
                                "ErrorCode", std::strlen("ErrorCode")))
                    {
                        error_code =
                            std::string(node->value(), node->value_size());
                    }
                    else if (compare(node->local_name(),
                                     node->local_name_size(), "Message",
                                     std::strlen("Message")))
                    {
                        message =
                            std::string(node->value(), node->value_size());
                    }

                    if (!error_code.empty() && !message.empty())
                    {
                        throw exception(message + " (error code: " +
                                        error_code + ")");
                    }
                }
            }

            throw exception("Unable to parse response");
        }

        // For each protocol in the response, look for the appropriate
        // protocol type (internal/external) and then look for the
        // corresponding <ASUrl/> element
        std::string protocol;
        autodiscover_result result;
        for (int i = 0; i < 2; i++)
        {
            for (auto protocol_node = account_node->first_node(); protocol_node;
                 protocol_node = protocol_node->next_sibling())
            {
                if (i >= 1)
                {
                    protocol = "EXCH";
                }
                else
                {
                    protocol = "EXPR";
                }
                if (compare(protocol_node->local_name(),
                            protocol_node->local_name_size(), "Protocol",
                            std::strlen("Protocol")))
                {
                    for (auto type_node = protocol_node->first_node();
                         type_node; type_node = type_node->next_sibling())
                    {
                        if (compare(type_node->local_name(),
                                    type_node->local_name_size(), "Type",
                                    std::strlen("Type")) &&
                            compare(type_node->value(), type_node->value_size(),
                                    protocol.c_str(),
                                    std::strlen(protocol.c_str())))
                        {
                            for (auto asurl_node = protocol_node->first_node();
                                 asurl_node;
                                 asurl_node = asurl_node->next_sibling())
                            {
                                if (compare(asurl_node->local_name(),
                                            asurl_node->local_name_size(),
                                            "ASUrl", std::strlen("ASUrl")))
                                {
                                    if (i >= 1)
                                    {
                                        result.internal_ews_url = std::string(
                                            asurl_node->value(),
                                            asurl_node->value_size());
                                        return result;
                                    }
                                    else
                                    {
                                        result.external_ews_url = std::string(
                                            asurl_node->value(),
                                            asurl_node->value_size());
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // If we reach this point, then either autodiscovery returned an
        // error or there is a redirect address to retry the autodiscover
        // lookup

        for (auto redirect_node = account_node->first_node(); redirect_node;
             redirect_node = redirect_node->next_sibling())
        {
            if (compare(redirect_node->local_name(),
                        redirect_node->local_name_size(), "RedirectAddr",
                        std::strlen("RedirectAddr")))
            {
                // Retry
                redirections++;
                const auto redirect_address = std::string(
                    redirect_node->value(), redirect_node->value_size());
                return get_exchange_web_services_url<RequestHandler>(
                    redirect_address, credentials, redirections, hints);
            }
        }

        throw exception("Autodiscovery failed unexpectedly");
    }
}
//! Set-up EWS library.
//!
//! Should be called when application is still in single-threaded context.
//! Calling this function more than once does no harm.
//!
//! Note: Function is not thread-safe
inline void set_up() EWS_NOEXCEPT { curl_global_init(CURL_GLOBAL_DEFAULT); }

//! Clean-up EWS library.
//!
//! You should call this function only when no other thread is running.
//! See libcurl(3) man-page or http://curl.haxx.se/libcurl/c/libcurl.html
//!
//! Note: Function is not thread-safe
inline void tear_down() EWS_NOEXCEPT { curl_global_cleanup(); }

//! \brief Returns the EWS URL by querying the Autodiscover service.
//!
//! \param user_smtp_address User's primary SMTP address
//! \param credentials The user's credentials
//! \return The Exchange Web Services URLs as autodiscover_result properties
#ifdef EWS_HAS_DEFAULT_TEMPLATE_ARGS_FOR_FUNCTIONS
template <typename RequestHandler = internal::http_request>
#else
template <typename RequestHandler>
#endif
inline autodiscover_result
get_exchange_web_services_url(const std::string& user_smtp_address,
                              const basic_credentials& credentials)
{
    ews::autodiscover_hints hints;
    return internal::get_exchange_web_services_url<RequestHandler>(
        user_smtp_address, credentials, 0U, hints);
}

//! \brief Returns the EWS URL by querying the Autodiscover service.
//!
//! \param user_smtp_address User's primary SMTP address
//! \param credentials The user's credentials
//! \param hints The url given by the user
//! \return The Exchange Web Services URLs as autodiscover_result properties
#ifdef EWS_HAS_DEFAULT_TEMPLATE_ARGS_FOR_FUNCTIONS
template <typename RequestHandler = internal::http_request>
#else
template <typename RequestHandler>
#endif
inline autodiscover_result
get_exchange_web_services_url(const std::string& user_smtp_address,
                              const basic_credentials& credentials,
                              const autodiscover_hints& hints)
{
    return internal::get_exchange_web_services_url<RequestHandler>(
        user_smtp_address, credentials, 0U, hints);
}

//! \brief The unique identifier and change key of an item in the
//! Exchange store.
//!
//! The ID uniquely identifies a concrete item throughout the Exchange
//! store and is not expected to change as long as the item exists. The
//! change key on the other hand identifies a specific version of an item.
//! It is expected to be changed whenever a property of the item is
//! changed. The change key is used for synchronization purposes on the
//! server. You only need to take care that the change key you include in a
//! service call is the most current one.
//!
//! Instances of this class are somewhat immutable. You can default
//! construct an item_id in which case valid() will always return false.
//! (Default construction is needed because we need item and it's
//! sub-classes to be default constructible.) Only item_ids that come from
//! an Exchange store are considered to be valid.
class item_id final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    //! Constructs an invalid item_id instance
    item_id() = default;
#else
    item_id() {}
#endif

    //! Constructs an <tt>\<ItemId></tt> from given \p id string.
    explicit item_id(std::string id) : id_(std::move(id)), change_key_() {}

    //! \brief Constructs an <tt>\<ItemId></tt> from given identifier and
    //! change key.
    item_id(std::string id, std::string change_key)
        : id_(std::move(id)), change_key_(std::move(change_key))
    {
    }

    //! Returns the identifier.
    const std::string& id() const EWS_NOEXCEPT { return id_; }

    //! Returns the change key
    const std::string& change_key() const EWS_NOEXCEPT { return change_key_; }

    //! Whether this item_id is expected to be valid
    bool valid() const EWS_NOEXCEPT { return !id_.empty(); }

    //! Serializes this item_id to an XML string
    std::string to_xml() const
    {
        return "<t:ItemId Id=\"" + id() + "\" ChangeKey=\"" + change_key() +
               "\"/>";
    }

    //! Makes an item_id instance from an <tt>\<ItemId></tt> XML element
    static item_id from_xml_element(const rapidxml::xml_node<>& elem)
    {
        auto id_attr = elem.first_attribute("Id");
        EWS_ASSERT(id_attr && "Missing attribute Id in <ItemId>");
        auto id = std::string(id_attr->value(), id_attr->value_size());
        auto ckey_attr = elem.first_attribute("ChangeKey");
        EWS_ASSERT(ckey_attr && "Missing attribute ChangeKey in <ItemId>");
        auto ckey = std::string(ckey_attr->value(), ckey_attr->value_size());
        return item_id(std::move(id), std::move(ckey));
    }

private:
    // case-sensitive; therefore, comparisons between IDs must be
    // case-sensitive or binary
    std::string id_;
    std::string change_key_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(std::is_default_constructible<item_id>::value, "");
static_assert(std::is_copy_constructible<item_id>::value, "");
static_assert(std::is_copy_assignable<item_id>::value, "");
static_assert(std::is_move_constructible<item_id>::value, "");
static_assert(std::is_move_assignable<item_id>::value, "");
#endif

//! \brief Contains the unique identifier of an attachment.
//!
//! The AttachmentId element identifies an item or file attachment. This
//! element is used in CreateAttachment responses.
class attachment_id final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    attachment_id() = default;
#else
    attachment_id() {}
#endif

#ifndef EWS_DOXYGEN_SHOULD_SKIP_THIS
    explicit attachment_id(std::string id) : id_(std::move(id)) {}

    attachment_id(std::string id, item_id root_item_id)
        : id_(std::move(id)), root_item_id_(std::move(root_item_id))
    {
    }
#endif

    //! Returns the string representing the unique identifier of an
    //! attachment
    const std::string& id() const EWS_NOEXCEPT { return id_; }

    //! \brief Returns the item_id of the <em>parent</em> or <em>root</em>
    //! item.
    //!
    //! The root item is the item that contains the attachment.
    //!
    //! \note: the returned item_id is only valid and meaningful when you
    //! obtained this attachment_id in a call to \ref
    //! service::create_attachment.
    const item_id& root_item_id() const EWS_NOEXCEPT { return root_item_id_; }

    //! Whether this attachment_id is valid
    bool valid() const EWS_NOEXCEPT { return !id_.empty(); }

    std::string to_xml() const
    {
        std::stringstream sstr;
        sstr << "<t:AttachmentId Id=\"" << id() << "\"";
        if (root_item_id().valid())
        {
            sstr << " RootItemId=\"" << root_item_id().id()
                 << "\" RootItemChangeKey=\"" << root_item_id().change_key()
                 << "\"";
        }
        sstr << "/>";
        return sstr.str();
    }

    //! Makes an attachment_id instance from an \<AttachmentId> element
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
            root_item_id = std::string(root_item_id_attr->value(),
                                       root_item_id_attr->value_size());
            auto root_item_ckey_attr =
                elem.first_attribute("RootItemChangeKey");
            EWS_ASSERT(root_item_ckey_attr &&
                       "Expected attribute RootItemChangeKey");
            root_item_ckey = std::string(root_item_ckey_attr->value(),
                                         root_item_ckey_attr->value_size());
        }

        return root_item_id.empty()
                   ? attachment_id(std::move(id))
                   : attachment_id(std::move(id),
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

//! \brief Identifies a folder.
//!
//! Renders a <tt>\<FolderId></tt> element. Contains the identifier and
//! change key of a folder.
class folder_id
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    folder_id() = default;
    virtual ~folder_id() = default;
    folder_id(const folder_id&) = default;
#else
    folder_id() {}
    virtual ~folder_id() {}
#endif

    explicit folder_id(std::string id) : id_(std::move(id)), change_key_() {}

    folder_id(std::string id, std::string change_key)
        : id_(std::move(id)), change_key_(std::move(change_key))
    {
    }

    std::string to_xml() const { return this->to_xml_impl(); }

    //! Returns a string identifying a folder in the Exchange store
    const std::string& id() const EWS_NOEXCEPT { return id_; }

    //! Returns a string identifying a version of a folder
    const std::string& change_key() const EWS_NOEXCEPT { return change_key_; }

    //! Whether this folder_id is valid
    bool valid() const EWS_NOEXCEPT { return !id_.empty(); }

    //! Makes a folder_id instance from given XML element
    static folder_id from_xml_element(const rapidxml::xml_node<>& elem)
    {
        auto id_attr = elem.first_attribute("Id");
        EWS_ASSERT(id_attr &&
                   "Expected <ParentFolderId> to have an Id attribute");
        auto id = std::string(id_attr->value(), id_attr->value_size());

        std::string change_key;
        auto ck_attr = elem.first_attribute("ChangeKey");
        if (ck_attr)
        {
            change_key = std::string(ck_attr->value(), ck_attr->value_size());
        }
        return folder_id(std::move(id), std::move(change_key));
    }

#ifndef EWS_DOXYGEN_SHOULD_SKIP_THIS
protected:
    virtual std::string to_xml_impl() const
    {
        std::stringstream sstr;
        sstr << "<t:"
             << "FolderId Id=\"" << id_;
        if (!change_key_.empty())
        {
            sstr << "\" ChangeKey=\"" << change_key_;
        }
        sstr << "\"/>";
        return sstr.str();
    }
#endif

private:
    std::string id_;
    std::string change_key_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(std::is_default_constructible<folder_id>::value, "");
static_assert(std::is_copy_constructible<folder_id>::value, "");
static_assert(std::is_copy_assignable<folder_id>::value, "");
static_assert(std::is_move_constructible<folder_id>::value, "");
static_assert(std::is_move_assignable<folder_id>::value, "");
#endif

//! \brief Renders a <tt>\<DistinguishedFolderId></tt> element.
//!
//! Implicitly convertible from \ref standard_folder.
class distinguished_folder_id final : public folder_id
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    distinguished_folder_id() = delete;
#endif

    // Intentionally not explicit
    distinguished_folder_id(standard_folder folder)
        : folder_id(well_known_name(folder))
    {
    }

    distinguished_folder_id(standard_folder folder, std::string change_key)
        : folder_id(well_known_name(folder), std::move(change_key))
    {
    }

    // TODO: Constructor for EWS delegate access
    // distinguished_folder_id(standard_folder, mailbox) {}

    //! Returns the standard_folder enum for given string
    static standard_folder str_to_standard_folder(const std::string& name)
    {
        if (name == "calendar")
        {
            return standard_folder::calendar;
        }
        else if (name == "contacts")
        {
            return standard_folder::contacts;
        }
        else if (name == "deleteditems")
        {
            return standard_folder::deleted_items;
        }
        else if (name == "drafts")
        {
            return standard_folder::drafts;
        }
        else if (name == "inbox")
        {
            return standard_folder::inbox;
        }
        else if (name == "journal")
        {
            return standard_folder::journal;
        }
        else if (name == "notes")
        {
            return standard_folder::notes;
        }
        else if (name == "outbox")
        {
            return standard_folder::outbox;
        }
        else if (name == "sentitems")
        {
            return standard_folder::sent_items;
        }
        else if (name == "tasks")
        {
            return standard_folder::tasks;
        }
        else if (name == "msgfolderroot")
        {
            return standard_folder::msg_folder_root;
        }
        else if (name == "root")
        {
            return standard_folder::root;
        }
        else if (name == "junkemail")
        {
            return standard_folder::junk_email;
        }
        else if (name == "searchfolders")
        {
            return standard_folder::search_folders;
        }
        else if (name == "voicemail")
        {
            return standard_folder::voice_mail;
        }
        else if (name == "recoverableitemsroot")
        {
            return standard_folder::recoverable_items_root;
        }
        else if (name == "recoverableitemsdeletions")
        {
            return standard_folder::recoverable_items_deletions;
        }
        else if (name == "recoverableitemsversions")
        {
            return standard_folder::recoverable_items_versions;
        }
        else if (name == "recoverableitemspurges")
        {
            return standard_folder::recoverable_items_purges;
        }
        else if (name == "archiveroot")
        {
            return standard_folder::archive_root;
        }
        else if (name == "archivemsgfolderroot")
        {
            return standard_folder::archive_msg_folder_root;
        }
        else if (name == "archivedeleteditems")
        {
            return standard_folder::archive_deleted_items;
        }
        else if (name == "archiveinbox")
        {
            return standard_folder::archive_inbox;
        }
        else if (name == "archiverecoverableitemsroot")
        {
            return standard_folder::archive_recoverable_items_root;
        }
        else if (name == "archiverecoverableitemsdeletions")
        {
            return standard_folder::archive_recoverable_items_deletions;
        }
        else if (name == "archiverecoverableitemsversions")
        {
            return standard_folder::archive_recoverable_items_versions;
        }
        else if (name == "archiverecoverableitemspurges")
        {
            return standard_folder::archive_recoverable_items_purges;
        }
        else if (name == "syncissues")
        {
            return standard_folder::sync_issues;
        }
        else if (name == "conflicts")
        {
            return standard_folder::conflicts;
        }
        else if (name == "localfailures")
        {
            return standard_folder::local_failures;
        }
        else if (name == "serverfailures")
        {
            return standard_folder::server_failures;
        }
        else if (name == "recipientcache")
        {
            return standard_folder::recipient_cache;
        }
        else if (name == "quickcontacts")
        {
            return standard_folder::quick_contacts;
        }
        else if (name == "conversationhistory")
        {
            return standard_folder::conversation_history;
        }
        else if (name == "adminauditlogs")
        {
            return standard_folder::admin_audit_logs;
        }
        else if (name == "todosearch")
        {
            return standard_folder::todo_search;
        }
        else if (name == "mycontacts")
        {
            return standard_folder::my_contacts;
        }
        else if (name == "directory")
        {
            return standard_folder::directory;
        }
        else if (name == "imcontactlist")
        {
            return standard_folder::im_contact_list;
        }
        else if (name == "peopleconnect")
        {
            return standard_folder::people_connect;
        }
        else if (name == "favorites")
        {
            return standard_folder::favorites;
        }

        throw exception("Unrecognized folder name");
    }

    //! Returns the well-known name for given standard_folder as string.
    static std::string well_known_name(standard_folder enumeration)
    {
        std::string name;
        switch (enumeration)
        {
        case standard_folder::calendar:
            name = "calendar";
            break;

        case standard_folder::contacts:
            name = "contacts";
            break;

        case standard_folder::deleted_items:
            name = "deleteditems";
            break;

        case standard_folder::drafts:
            name = "drafts";
            break;

        case standard_folder::inbox:
            name = "inbox";
            break;

        case standard_folder::journal:
            name = "journal";
            break;

        case standard_folder::notes:
            name = "notes";
            break;

        case standard_folder::outbox:
            name = "outbox";
            break;

        case standard_folder::sent_items:
            name = "sentitems";
            break;

        case standard_folder::tasks:
            name = "tasks";
            break;

        case standard_folder::msg_folder_root:
            name = "msgfolderroot";
            break;

        case standard_folder::root:
            name = "root";
            break;

        case standard_folder::junk_email:
            name = "junkemail";
            break;

        case standard_folder::search_folders:
            name = "searchfolders";
            break;

        case standard_folder::voice_mail:
            name = "voicemail";
            break;

        case standard_folder::recoverable_items_root:
            name = "recoverableitemsroot";
            break;

        case standard_folder::recoverable_items_deletions:
            name = "recoverableitemsdeletions";
            break;

        case standard_folder::recoverable_items_versions:
            name = "recoverableitemsversions";
            break;

        case standard_folder::recoverable_items_purges:
            name = "recoverableitemspurges";
            break;

        case standard_folder::archive_root:
            name = "archiveroot";
            break;

        case standard_folder::archive_msg_folder_root:
            name = "archivemsgfolderroot";
            break;

        case standard_folder::archive_deleted_items:
            name = "archivedeleteditems";
            break;

        case standard_folder::archive_inbox:
            name = "archiveinbox";
            break;

        case standard_folder::archive_recoverable_items_root:
            name = "archiverecoverableitemsroot";
            break;

        case standard_folder::archive_recoverable_items_deletions:
            name = "archiverecoverableitemsdeletions";
            break;

        case standard_folder::archive_recoverable_items_versions:
            name = "archiverecoverableitemsversions";
            break;

        case standard_folder::archive_recoverable_items_purges:
            name = "archiverecoverableitemspurges";
            break;

        case standard_folder::sync_issues:
            name = "syncissues";
            break;

        case standard_folder::conflicts:
            name = "conflicts";
            break;

        case standard_folder::local_failures:
            name = "localfailures";
            break;

        case standard_folder::server_failures:
            name = "serverfailures";
            break;

        case standard_folder::recipient_cache:
            name = "recipientcache";
            break;

        case standard_folder::quick_contacts:
            name = "quickcontacts";
            break;

        case standard_folder::conversation_history:
            name = "conversationhistory";
            break;

        case standard_folder::admin_audit_logs:
            name = "adminauditlogs";
            break;

        case standard_folder::todo_search:
            name = "todosearch";
            break;

        case standard_folder::my_contacts:
            name = "mycontacts";
            break;

        case standard_folder::directory:
            name = "directory";
            break;

        case standard_folder::im_contact_list:
            name = "imcontactlist";
            break;

        case standard_folder::people_connect:
            name = "peopleconnect";
            break;

        case standard_folder::favorites:
            name = "favorites";
            break;

        default:
            throw exception("Unrecognized folder name");
        };
        return name;
    }

private:
    std::string to_xml_impl() const override
    {
        std::stringstream sstr;
        sstr << "<t:"
             << "DistinguishedFolderId Id=\"";
        sstr << id();
        if (!change_key().empty())
        {
            sstr << "\" ChangeKey=\"" << change_key();
        }
        sstr << "\"/>";
        return sstr.str();
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<distinguished_folder_id>::value,
              "");
static_assert(std::is_copy_constructible<distinguished_folder_id>::value, "");
static_assert(std::is_copy_assignable<distinguished_folder_id>::value, "");
static_assert(std::is_move_constructible<distinguished_folder_id>::value, "");
static_assert(std::is_move_assignable<distinguished_folder_id>::value, "");
#endif

//! Represents a <tt>\<FileAttachment></tt> or an <tt>\<ItemAttachment></tt>
class attachment final
{
public:
    // Note that we are careful to avoid use of get_element_by_qname() here
    // because it might return elements of the enclosed item and not what
    // we'd expect

    //! Describes whether an attachment contains a file or another item.
    enum class type
    {
        //! An <tt>\<ItemAttachment></tt>
        item,

        //! A <tt>\<FileAttachment></tt>
        file
    };

    attachment() : xml_(), type_(type::item) {}

    //! Returns this attachment's attachment_id
    attachment_id id() const
    {
        const auto node = get_node("AttachmentId");
        return node ? attachment_id::from_xml_element(*node) : attachment_id();
    }

    //! Returns the attachment's name
    std::string name() const
    {
        const auto node = get_node("Name");
        return node ? std::string(node->value(), node->value_size()) : "";
    }

    //! Returns the attachment's content type
    std::string content_type() const
    {
        const auto node = get_node("ContentType");
        return node ? std::string(node->value(), node->value_size()) : "";
    }

    //! \brief Returns the attachment's contents
    //!
    //! If this is a <tt>\<FileAttachment></tt>, returns the Base64-encoded
    //! contents of the file attachment. If this is an \<ItemAttachment>,
    //! the empty string.
    std::string content() const
    {
        const auto node = get_node("Content");
        return node ? std::string(node->value(), node->value_size()) : "";
    }

    //! \brief Returns the attachment's size in bytes
    //!
    //! If this is a <tt>\<FileAttachment></tt>, returns the size in bytes
    //! of the file attachment; otherwise 0.
    std::size_t content_size() const
    {
        const auto node = get_node("Size");
        if (!node)
        {
            return 0U;
        }
        auto size = std::string(node->value(), node->value_size());
        return size.empty() ? 0U : std::stoul(size);
    }

    //! Returns either type::file or type::item
    type get_type() const EWS_NOEXCEPT { return type_; }

    //! \brief Write contents to a file
    //!
    //! If this is a <tt>\<FileAttachment></tt>, writes content to file.
    //! Does nothing if this is an <tt>\<ItemAttachment></tt>.  Returns the
    //! number of bytes written.
    std::size_t write_content_to_file(const std::string& file_path) const
    {
        if (get_type() == type::item)
        {
            return 0U;
        }

        const auto raw_bytes = internal::base64::decode(content());

        std::ofstream ofstr(file_path, std::ofstream::out | std::ios::binary);
        if (!ofstr.is_open())
        {
            if (file_path.empty())
            {
                throw exception(
                    "Could not open file for writing: no file name given");
            }

            throw exception("Could not open file for writing: " + file_path);
        }

        std::copy(begin(raw_bytes), end(raw_bytes),
                  std::ostreambuf_iterator<char>(ofstr));
        ofstr.close();
        return raw_bytes.size();
    }

    //! Returns this attachment serialized to XML
    std::string to_xml() const { return xml_.to_string(); }

    //! Constructs an attachment from a given XML element \p elem.
    static attachment from_xml_element(const rapidxml::xml_node<>& elem)
    {
        const auto elem_name =
            std::string(elem.local_name(), elem.local_name_size());
        EWS_ASSERT(
            (elem_name == "FileAttachment" || elem_name == "ItemAttachment") &&
            "Expected <FileAttachment> or <ItemAttachment>");
        return attachment(elem_name == "FileAttachment" ? type::file
                                                        : type::item,
                          internal::xml_subtree(elem));
    }

    //! \brief Creates a new <tt>\<FileAttachment></tt> from a given file.
    //!
    //! Returns a new <tt>\<FileAttachment></tt> that you can pass to
    //! ews::service::create_attachment in order to create the attachment on
    //! the server.
    //!
    //! \param file_path Path to an existing and readable file
    //! \param content_type The (RFC 2046) MIME content type of the
    //!        attachment
    //! \param name A name for this attachment
    //!
    //! On Windows you can use HKEY_CLASSES_ROOT/MIME/Database/Content Type
    //! registry hive to get the content type from a file extension. On a
    //! UNIX see magic(5) and file(1).
    static attachment from_file(const std::string& file_path,
                                std::string content_type, std::string name)
    {
        using internal::create_node;

        // Try open file
        std::ifstream ifstr(file_path, std::ifstream::in | std::ios::binary);
        if (!ifstr.is_open())
        {
            throw exception("Could not open file for reading: " + file_path);
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

        auto& attachment_node =
            create_node(*obj.xml_.document(), "t:FileAttachment");
        create_node(attachment_node, "t:Name", name);
        create_node(attachment_node, "t:ContentType", content_type);
        create_node(attachment_node, "t:Content", content);
        create_node(attachment_node, "t:Size", std::to_string(buffer.size()));

        return obj;
    }

    static attachment from_item(const item& the_item,
                                const std::string& name); // implemented
                                                          // below

private:
    internal::xml_subtree xml_;
    type type_;

    attachment(type&& t, internal::xml_subtree&& xml)
        : xml_(std::move(xml)), type_(std::move(t))
    {
    }

    rapidxml::xml_node<>* get_node(const char* local_name) const
    {
        using rapidxml::internal::compare;

        const auto attachment_node = xml_.root();
        if (!attachment_node)
        {
            return nullptr;
        }

        rapidxml::xml_node<>* node = nullptr;
        for (auto child = attachment_node->first_node(); child != nullptr;
             child = child->next_sibling())
        {
            if (compare(child->local_name(), child->local_name_size(),
                        local_name, std::strlen(local_name)))
            {
                node = child;
                break;
            }
        }
        return node;
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(std::is_default_constructible<attachment>::value, "");
static_assert(std::is_copy_constructible<attachment>::value, "");
static_assert(std::is_copy_assignable<attachment>::value, "");
static_assert(std::is_move_constructible<attachment>::value, "");
static_assert(std::is_move_assignable<attachment>::value, "");
#endif

namespace internal
{
    // Parse response class and response code from given element.
    inline std::pair<response_class, response_code>
    parse_response_class_and_code(const rapidxml::xml_node<>& elem)
    {
        using rapidxml::internal::compare;

        auto cls = response_class::success;
        auto response_class_attr = elem.first_attribute("ResponseClass");
        if (compare(response_class_attr->value(),
                    response_class_attr->value_size(), "Error", 5))
        {
            cls = response_class::error;
        }
        else if (compare(response_class_attr->value(),
                         response_class_attr->value_size(), "Warning", 7))
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
            auto response_code_elem = elem.first_node_ns(
                uri<>::microsoft::messages(), "ResponseCode");
            EWS_ASSERT(response_code_elem && "Expected <ResponseCode> element");
            code = str_to_response_code(response_code_elem->value());
        }

        return std::make_pair(cls, code);
    }

    // Iterate over <Items> array and execute given function for each node.
    //
    // items_elem: an <Items> element
    // func: A callable that is invoked for each item in the response
    // message's <Items> array. A const rapidxml::xml_node& is passed to
    // that callable.
    template <typename Func>
    inline void for_each_child_node(const rapidxml::xml_node<>& parent_node,
                                    Func func)
    {
        for (auto child = parent_node.first_node(); child;
             child = child->next_sibling())
        {
            func(*child);
        }
    }

    // Base-class for all response messages
    class response_message_base
    {
    public:
        response_message_base(response_class cls, response_code code)
            : cls_(cls), code_(code)
        {
        }

        response_class get_response_class() const EWS_NOEXCEPT { return cls_; }

        bool success() const EWS_NOEXCEPT
        {
            return get_response_class() == response_class::success;
        }

        response_code get_response_code() const EWS_NOEXCEPT { return code_; }

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

        response_message_with_items(response_class cls, response_code code,
                                    std::vector<item_type> items)
            : response_message_base(cls, code), items_(std::move(items))
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
        static create_item_response_message parse(http_response&&);

    private:
        create_item_response_message(response_class cls, response_code code,
                                     std::vector<item_id> items)
            : response_message_with_items<item_id>(cls, code, std::move(items))
        {
        }
    };

    class find_item_response_message final
        : public response_message_with_items<item_id>
    {
    public:
        // implemented below
        static find_item_response_message parse(http_response&&);

    private:
        find_item_response_message(response_class cls, response_code code,
                                   std::vector<item_id> items)
            : response_message_with_items<item_id>(cls, code, std::move(items))
        {
        }
    };

    class find_calendar_item_response_message final
        : public response_message_with_items<calendar_item>
    {
    public:
        // implemented below
        static find_calendar_item_response_message parse(http_response&);

    private:
        find_calendar_item_response_message(response_class cls,
                                            response_code code,
                                            std::vector<calendar_item> items)
            : response_message_with_items<calendar_item>(cls, code,
                                                         std::move(items))
        {
        }
    };

    class update_item_response_message final
        : public response_message_with_items<item_id>
    {
    public:
        // implemented below
        static update_item_response_message parse(http_response&&);

    private:
        update_item_response_message(response_class cls, response_code code,
                                     std::vector<item_id> items)
            : response_message_with_items<item_id>(cls, code, std::move(items))
        {
        }
    };

    template <typename ItemType>
    class get_item_response_message final
        : public response_message_with_items<ItemType>
    {
    public:
        // implemented below
        static get_item_response_message parse(http_response&&);

    private:
        get_item_response_message(response_class cls, response_code code,
                                  std::vector<ItemType> items)
            : response_message_with_items<ItemType>(cls, code, std::move(items))
        {
        }
    };

    template <typename ItemType> class get_item_response_messages final
    {
    public:
        typedef ItemType item_type;
        typedef std::tuple<response_class, response_code,
                           std::vector<item_type>>
            response_message;

        std::vector<item_type> items() const
        {
            std::vector<item_type> items;
            items.reserve(messages_.size()); // Seems like there is always one
                                             // item per message
            for (const auto& msg : messages_)
            {
                const auto& msg_items = std::get<2>(msg);
                std::copy(begin(msg_items), end(msg_items),
                          std::back_inserter(items));
            }
            return items;
        }

        bool success() const
        {
            // Consequently, this means we're aborting here because of a
            // warning. Is this desired? Don't think so. At least it is
            // consistent with response_message_base::success().

            return std::all_of(
                begin(messages_), end(messages_), [](const response_message& msg) {
                    return std::get<0>(msg) == response_class::success;
                });
        }

        response_code first_error_or_warning() const
        {
            auto it = std::find_if_not(
                begin(messages_), end(messages_), [](const response_message& msg) {
                    return std::get<0>(msg) == response_class::success;
                });
            return it == end(messages_) ? response_code::no_error
                                        : std::get<1>(*it);
        }

        // implemented below
        static get_item_response_messages parse(http_response&&);

    private:
        explicit get_item_response_messages(
            std::vector<response_message>&& messages)
            : messages_(std::move(messages))
        {
        }

        std::vector<response_message> messages_;
    };

    class create_attachment_response_message final
        : public response_message_base
    {
    public:
        static create_attachment_response_message
        parse(http_response&& response)
        {
            const auto doc = parse_response(std::move(response));
            auto elem =
                get_element_by_qname(*doc, "CreateAttachmentResponseMessage",
                                     uri<>::microsoft::messages());

            EWS_ASSERT(
                elem &&
                "Expected <CreateAttachmentResponseMessage>, got nullptr");

            const auto cls_and_code = parse_response_class_and_code(*elem);

            auto attachments_element = elem->first_node_ns(
                uri<>::microsoft::messages(), "Attachments");
            EWS_ASSERT(attachments_element && "Expected <Attachments> element");

            auto ids = std::vector<attachment_id>();
            for (auto attachment_elem = attachments_element->first_node();
                 attachment_elem;
                 attachment_elem = attachment_elem->next_sibling())
            {
                auto attachment_id_elem = attachment_elem->first_node_ns(
                    uri<>::microsoft::types(), "AttachmentId");
                EWS_ASSERT(attachment_id_elem &&
                           "Expected <AttachmentId> in response");
                ids.emplace_back(
                    attachment_id::from_xml_element(*attachment_id_elem));
            }
            return create_attachment_response_message(
                cls_and_code.first, cls_and_code.second, std::move(ids));
        }

        const std::vector<attachment_id>& attachment_ids() const EWS_NOEXCEPT
        {
            return ids_;
        }

    private:
        create_attachment_response_message(
            response_class cls, response_code code,
            std::vector<attachment_id> attachment_ids)
            : response_message_base(cls, code), ids_(std::move(attachment_ids))
        {
        }

        std::vector<attachment_id> ids_;
    };

    class get_attachment_response_message final : public response_message_base
    {
    public:
        static get_attachment_response_message parse(http_response&& response)
        {
            const auto doc = parse_response(std::move(response));
            auto elem =
                get_element_by_qname(*doc, "GetAttachmentResponseMessage",
                                     uri<>::microsoft::messages());

            EWS_ASSERT(elem &&
                       "Expected <GetAttachmentResponseMessage>, got nullptr");

            const auto cls_and_code = parse_response_class_and_code(*elem);

            auto attachments_element = elem->first_node_ns(
                uri<>::microsoft::messages(), "Attachments");
            EWS_ASSERT(attachments_element && "Expected <Attachments> element");
            auto attachments = std::vector<attachment>();
            for (auto attachment_elem = attachments_element->first_node();
                 attachment_elem;
                 attachment_elem = attachment_elem->next_sibling())
            {
                attachments.emplace_back(
                    attachment::from_xml_element(*attachment_elem));
            }
            return get_attachment_response_message(cls_and_code.first,
                                                   cls_and_code.second,
                                                   std::move(attachments));
        }

        const std::vector<attachment>& attachments() const EWS_NOEXCEPT
        {
            return attachments_;
        }

    private:
        get_attachment_response_message(response_class cls, response_code code,
                                        std::vector<attachment> attachments)
            : response_message_base(cls, code),
              attachments_(std::move(attachments))
        {
        }

        std::vector<attachment> attachments_;
    };

    class send_item_response_message final : public response_message_base
    {
    public:
        static send_item_response_message parse(http_response&& response)
        {
            const auto doc = parse_response(std::move(response));
            auto elem = get_element_by_qname(*doc, "SendItemResponseMessage",
                                             uri<>::microsoft::messages());

            EWS_ASSERT(elem &&
                       "Expected <SendItemResponseMessage>, got nullptr");
            const auto result = parse_response_class_and_code(*elem);
            return send_item_response_message(result.first, result.second);
        }

    private:
        send_item_response_message(response_class cls, response_code code)
            : response_message_base(cls, code)
        {
        }
    };

    class delete_item_response_message final : public response_message_base
    {
    public:
        static delete_item_response_message parse(http_response&& response)
        {
            const auto doc = parse_response(std::move(response));
            auto elem = get_element_by_qname(*doc, "DeleteItemResponseMessage",
                                             uri<>::microsoft::messages());

            EWS_ASSERT(elem &&
                       "Expected <DeleteItemResponseMessage>, got nullptr");
            const auto result = parse_response_class_and_code(*elem);
            return delete_item_response_message(result.first, result.second);
        }

    private:
        delete_item_response_message(response_class cls, response_code code)
            : response_message_base(cls, code)
        {
        }
    };

    class delete_attachment_response_message final
        : public response_message_base
    {
    public:
        static delete_attachment_response_message
        parse(http_response&& response)
        {
            const auto doc = parse_response(std::move(response));
            auto elem =
                get_element_by_qname(*doc, "DeleteAttachmentResponseMessage",
                                     uri<>::microsoft::messages());

            EWS_ASSERT(
                elem &&
                "Expected <DeleteAttachmentResponseMessage>, got nullptr");
            const auto cls_and_code = parse_response_class_and_code(*elem);

            auto root_item_id = item_id();
            auto root_item_id_elem =
                elem->first_node_ns(uri<>::microsoft::messages(), "RootItemId");
            if (root_item_id_elem)
            {
                auto id_attr = root_item_id_elem->first_attribute("RootItemId");
                EWS_ASSERT(id_attr && "Expected RootItemId attribute");
                auto id = std::string(id_attr->value(), id_attr->value_size());

                auto change_key_attr =
                    root_item_id_elem->first_attribute("RootItemChangeKey");
                EWS_ASSERT(change_key_attr &&
                           "Expected RootItemChangeKey attribute");
                auto change_key = std::string(change_key_attr->value(),
                                              change_key_attr->value_size());

                root_item_id = item_id(std::move(id), std::move(change_key));
            }

            return delete_attachment_response_message(
                cls_and_code.first, cls_and_code.second, root_item_id);
        }

        item_id get_root_item_id() const { return root_item_id_; }

    private:
        delete_attachment_response_message(response_class cls,
                                           response_code code,
                                           item_id root_item_id)
            : response_message_base(cls, code),
              root_item_id_(std::move(root_item_id))
        {
        }

        item_id root_item_id_;
    };
}

//! \brief A thin wrapper for xs:dateTime formatted strings.
//!
//! Note About Dates in EWS
//!
//! Microsoft EWS uses date and date/time string representations as
//! described in http://www.w3.org/TR/xmlschema-2/, notably xs:dateTime (or
//! http://www.w3.org/2001/XMLSchema:dateTime) and xs:date (also known as
//! http://www.w3.org/2001/XMLSchema:date).
//!
//! For example, the lexical representation of xs:date is
//!
//!     '-'? yyyy '-' mm '-' dd zzzzzz?
//!
//! whereas the z represents the timezone. Two examples of date strings are:
//! 2000-01-16Z and 1981-07-02 (the Z means Zulu time which is the same as
//! UTC). xs:dateTime is formatted accordingly, just with a time component;
//! you get the idea.
//!
//! This library does not interpret, parse, or in any way touch date nor
//! date/time strings in any circumstance. This library provides the
//! \ref date_time class, which acts solely as a thin wrapper to make the
//! signatures of public API functions more type-rich and easier to
//! understand. date_time is implicitly convertible from std::string.
//!
//! If your date or date/time strings are not formatted properly, Microsoft
//! EWS will likely give you a SOAP fault which this library transports to
//! you as an exception of type ews::soap_fault.
class date_time final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    date_time() = default;
#else
    date_time() {}
#endif

    date_time(std::string str) // intentionally not explicit
        : val_(std::move(str))
    {
    }

    const std::string& to_string() const EWS_NOEXCEPT { return val_; }

    inline bool is_set() const EWS_NOEXCEPT { return !val_.empty(); }

private:
    friend bool operator==(const date_time&, const date_time&);
    std::string val_;
};

inline bool operator==(const date_time& lhs, const date_time& rhs)
{
    return lhs.val_ == rhs.val_;
}

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(std::is_default_constructible<date_time>::value, "");
static_assert(std::is_copy_constructible<date_time>::value, "");
static_assert(std::is_copy_assignable<date_time>::value, "");
static_assert(std::is_move_constructible<date_time>::value, "");
static_assert(std::is_move_assignable<date_time>::value, "");
#endif

//! \brief A xs:date formatted string
//!
//! Exactly the same type as date_time. Used to indicate that an API
//! expects a date without a time value
typedef date_time date;

//! \brief Specifies a time interval
//!
//! A thin wrapper around xs:duration formatted strings.
//!
//! The time interval is specified in the following form
//! <tt>PnYnMnDTnHnMnS</tt> where:
//!
//! \li P indicates the period(required)
//! \li nY indicates the number of years
//! \li nM indicates the number of months
//! \li nD indicates the number of days
//! \li T indicates the start of a time section (required if you are going
//!     to specify hours, minutes, or seconds)
//! \li nH indicates the number of hours
//! \li nM indicates the number of minutes
//! \li nS indicates the number of seconds
class duration final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    duration() = default;
#else
    duration() {}
#endif

    duration(std::string str) // intentionally not explicit
        : val_(std::move(str))
    {
    }

    const std::string& to_string() const EWS_NOEXCEPT { return val_; }

    inline bool is_set() const EWS_NOEXCEPT { return !val_.empty(); }

private:
    friend bool operator==(const duration&, const duration&);
    std::string val_;
};

inline bool operator==(const duration& lhs, const duration& rhs)
{
    return lhs.val_ == rhs.val_;
}

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(std::is_default_constructible<duration>::value, "");
static_assert(std::is_copy_constructible<duration>::value, "");
static_assert(std::is_copy_assignable<duration>::value, "");
static_assert(std::is_move_constructible<duration>::value, "");
static_assert(std::is_move_assignable<duration>::value, "");
#endif

//! Specifies the type of a <tt>\<Body></tt> element
enum class body_type
{
    //! \brief The response will return the richest available content.
    //!
    //! This is useful if it is unknown whether the content is text or HTML.
    //! The returned body will be text if the stored body is plain-text.
    //! Otherwise, the response will return HTML if the stored body is in
    //! either HTML or RTF format. This is usually the default value.
    best,

    //! The response will return an item body as plain-text
    plain_text,

    //! The response will return an item body as HTML
    html
};

namespace internal
{
    inline std::string body_type_str(body_type type)
    {
        switch (type)
        {
        case body_type::best:
            return "Best";
        case body_type::plain_text:
            return "Text";
        case body_type::html:
            return "HTML";
        default:
            throw exception("Bad enum value");
        }
    }
}

//! \brief Represents the actual content of a message.
//!
//! A \<Body/> element can be of type Best, HTML, or plain-text.
class body final
{
public:
    //! Creates an empty body element; body_type is plain-text
    body() : content_(), type_(body_type::plain_text), is_truncated_(false) {}

    //! Creates a new body element with given content and type
    explicit body(std::string content, body_type type = body_type::plain_text)
        : content_(std::move(content)), type_(type), is_truncated_(false)
    {
    }

    body_type type() const EWS_NOEXCEPT { return type_; }

    void set_type(body_type type) { type_ = type; }

    bool is_truncated() const EWS_NOEXCEPT { return is_truncated_; }

    void set_truncated(bool truncated) { is_truncated_ = truncated; }

    const std::string& content() const EWS_NOEXCEPT { return content_; }

    void set_content(std::string content) { content_ = std::move(content); }

    std::string to_xml() const
    {
        // FIXME: what about IsTruncated attribute?
        static const auto cdata_beg = std::string("<![CDATA[");
        static const auto cdata_end = std::string("]]>");

        std::stringstream sstr;
        sstr << "<t:Body BodyType=\"" << internal::body_type_str(type());
        sstr << "\">";
        if (type() == body_type::html &&
            !(content_.compare(0, cdata_beg.length(), cdata_beg) == 0))
        {
            sstr << cdata_beg << content_ << cdata_end;
        }
        else
        {
            sstr << content_;
        }
        sstr << "</t:Body>";
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

//! \brief Represents an item's <tt>\<MimeContent CharacterSet="" /></tt>
//! element.
//!
//! Contains the ASCII MIME stream of an object that is represented in
//! base64Binary format (as in RFC 2045).
class mime_content final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    mime_content() = default;
#else
    mime_content() {}
#endif

    //! Copies \p len bytes from \p ptr into an internal buffer.
    mime_content(std::string charset, const char* const ptr, std::size_t len)
        : charset_(std::move(charset)), bytearray_(ptr, ptr + len)
    {
    }

    //! Returns how the string is encoded, e.g., "UTF-8"
    const std::string& character_set() const EWS_NOEXCEPT { return charset_; }

    //! Note: the pointer to the data is not 0-terminated
    const char* bytes() const EWS_NOEXCEPT { return bytearray_.data(); }

    std::size_t len_bytes() const EWS_NOEXCEPT { return bytearray_.size(); }

    //! Returns true if no MIME content is available.
    //!
    //! Note that a \<MimeContent> property is only included in a GetItem
    //! response when explicitly requested using additional properties. This
    //! function lets you test whether MIME content is available.
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

//! \brief Represents a SMTP mailbox.
//!
//! Identifies a fully resolved email address. Usually represents a
//! contact's email address, a message recipient, or the organizer of a
//! meeting.
class mailbox final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    //! \brief Creates a new undefined mailbox.
    //!
    //! Only useful as return value to indicate that no mailbox is set or
    //! available. (Good candidate for {boost,std}::optional)
    mailbox() = default;
#else
    mailbox() {}
#endif

    explicit mailbox(item_id id)
        : id_(std::move(id)), value_(), name_(), routing_type_(),
          mailbox_type_()
    {
    }

    explicit mailbox(std::string value, std::string name = std::string(),
                     std::string routing_type = std::string(),
                     std::string mailbox_type = std::string())
        : id_(), value_(std::move(value)), name_(std::move(name)),
          routing_type_(std::move(routing_type)),
          mailbox_type_(std::move(mailbox_type))
    {
    }

    //! True if this mailbox is undefined
    bool none() const EWS_NOEXCEPT { return value_.empty() && !id_.valid(); }

    // TODO: rename
    const item_id& id() const EWS_NOEXCEPT { return id_; }

    //! Returns the email address
    const std::string& value() const EWS_NOEXCEPT { return value_; }

    //! \brief Returns the name of the mailbox user.
    //!
    //! This attribute is optional.
    const std::string& name() const EWS_NOEXCEPT { return name_; }

    //! \brief Returns the routing type.
    //
    //! This attribute is optional. Default is SMTP
    const std::string& routing_type() const EWS_NOEXCEPT
    {
        return routing_type_;
    }

    //! \brief Returns the mailbox type.
    //!
    //! This attribute is optional.
    const std::string& mailbox_type() const EWS_NOEXCEPT
    {
        return mailbox_type_;
    }

    //! Returns the XML serialized string version of this mailbox instance
    std::string to_xml() const
    {
        std::stringstream sstr;
        sstr << "<t:Mailbox>";
        if (id().valid())
        {
            sstr << id().to_xml();
        }
        else
        {
            sstr << "<t:EmailAddress>" << value() << "</t:EmailAddress>";

            if (!name().empty())
            {
                sstr << "<t:Name>" << name() << "</t:Name>";
            }

            if (!routing_type().empty())
            {
                sstr << "<t:RoutingType>" << routing_type()
                     << "</t:RoutingType>";
            }

            if (!mailbox_type().empty())
            {
                sstr << "<t:MailboxType>" << mailbox_type()
                     << "</t:MailboxType>";
            }
        }
        sstr << "</t:Mailbox>";
        return sstr.str();
    }

    //! \brief Creates a new \<Mailbox> XML element and appends it to given
    //! parent node.
    //!
    //! Returns a reference to the newly created element.
    rapidxml::xml_node<>& to_xml_element(rapidxml::xml_node<>& parent) const
    {
        auto doc = parent.document();

        EWS_ASSERT(doc && "parent node needs to be somewhere in a document");

        using namespace internal;
        auto& mailbox_node = create_node(parent, "t:Mailbox");

        if (!id_.valid())
        {
            EWS_ASSERT(!value_.empty() &&
                       "Neither item_id nor value set in mailbox instance");

            create_node(mailbox_node, "t:EmailAddress", value_);

            if (!name_.empty())
            {
                create_node(mailbox_node, "t:Name", name_);
            }

            if (!routing_type_.empty())
            {
                create_node(mailbox_node, "t:RoutingType", routing_type_);
            }

            if (!mailbox_type_.empty())
            {
                create_node(mailbox_node, "t:MailboxType", mailbox_type_);
            }
        }
        else
        {
            auto item_id_node = &create_node(mailbox_node, "t:ItemId");

            auto ptr_to_key = doc->allocate_string("Id");
            auto ptr_to_value = doc->allocate_string(id_.id().c_str());
            item_id_node->append_attribute(
                doc->allocate_attribute(ptr_to_key, ptr_to_value));

            ptr_to_key = doc->allocate_string("ChangeKey");
            ptr_to_value = doc->allocate_string(id_.change_key().c_str());
            item_id_node->append_attribute(
                doc->allocate_attribute(ptr_to_key, ptr_to_value));
        }
        return mailbox_node;
    }

    //! Makes a mailbox instance from a \<Mailbox> XML element
    static mailbox from_xml_element(const rapidxml::xml_node<>& elem)
    {
        using rapidxml::internal::compare;

        //  <Mailbox>
        //      <Name/>
        //      <EmailAddress/>
        //      <RoutingType/>
        //      <MailboxType/>
        //      <ItemId/>
        //  </Mailbox>
        //
        // <EmailAddress> child element is required except when dealing
        // with a private distribution list or a contact from a user's
        // contacts folder, in which case the <ItemId> child element is
        // used instead

        std::string name;
        std::string address;
        std::string routing_type;
        std::string mailbox_type;
        item_id id;

        for (auto node = elem.first_node(); node; node = node->next_sibling())
        {
            if (compare(node->local_name(), node->local_name_size(), "Name",
                        std::strlen("Name")))
            {
                name = std::string(node->value(), node->value_size());
            }
            else if (compare(node->local_name(), node->local_name_size(),
                             "EmailAddress", std::strlen("EmailAddress")))
            {
                address = std::string(node->value(), node->value_size());
            }
            else if (compare(node->local_name(), node->local_name_size(),
                             "RoutingType", std::strlen("RoutingType")))
            {
                routing_type = std::string(node->value(), node->value_size());
            }
            else if (compare(node->local_name(), node->local_name_size(),
                             "MailboxType", std::strlen("MailboxType")))
            {
                mailbox_type = std::string(node->value(), node->value_size());
            }
            else if (compare(node->local_name(), node->local_name_size(),
                             "ItemId", std::strlen("ItemId")))
            {
                id = item_id::from_xml_element(*node);
            }
            else
            {
                throw exception("Unexpected child element in <Mailbox>");
            }
        }

        if (!id.valid())
        {
            EWS_ASSERT(!address.empty() &&
                       "<EmailAddress> element value can't be empty");

            return mailbox(std::move(address), std::move(name),
                           std::move(routing_type), std::move(mailbox_type));
        }

        return mailbox(std::move(id));
    }

private:
    item_id id_;
    std::string value_;
    std::string name_;
    std::string routing_type_;
    std::string mailbox_type_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(std::is_default_constructible<mailbox>::value, "");
static_assert(std::is_copy_constructible<mailbox>::value, "");
static_assert(std::is_copy_assignable<mailbox>::value, "");
static_assert(std::is_move_constructible<mailbox>::value, "");
static_assert(std::is_move_assignable<mailbox>::value, "");
#endif

//! \brief An attendee of a meeting or a meeting room.
//!
//! An attendee is just a mailbox for the most part. The other two
//! properties, ResponseType and LastResponseTime, are read-only properties
//! that usually get populated by the Exchange server and can be used to
//! track attendee responses.
class attendee final
{
public:
    attendee(mailbox address, response_type type, date_time last_response_time)
        : mailbox_(std::move(address)), response_type_(std::move(type)),
          last_response_time_(std::move(last_response_time))
    {
    }

    explicit attendee(mailbox address)
        : mailbox_(std::move(address)),
          response_type_(ews::response_type::unknown), last_response_time_()
    {
    }

    //! Returns this attendee's email address
    const mailbox& get_mailbox() const EWS_NOEXCEPT { return mailbox_; }

    //! \brief Returns this attendee's response
    //!
    //! This property is only relevant to a meeting organizer's calendar
    //! item.
    response_type get_response_type() const EWS_NOEXCEPT
    {
        return response_type_;
    }

    //! Returns the date and time of the latest response that was received
    const date_time& get_last_response_time() const EWS_NOEXCEPT
    {
        return last_response_time_;
    }

    //! Returns the XML serialized version of this attendee instance
    std::string to_xml() const
    {
        std::stringstream sstr;
        sstr << "<t:Attendee>";
        sstr << mailbox_.to_xml();
        sstr << "<t:ResponseType>" << internal::enum_to_str(response_type_)
             << "</t:ResponseType>";
        sstr << "<t:LastResponseTime>" << last_response_time_.to_string()
             << "</t:LastResponseTime>";
        sstr << "</t:Attendee>";
        return sstr.str();
    }

    //! \brief Creates a new \<Attendee> XML element and appends it to
    //! given parent node.
    //!
    //! Returns a reference to the newly created element.
    rapidxml::xml_node<>& to_xml_element(rapidxml::xml_node<>& parent) const
    {
        EWS_ASSERT(parent.document() &&
                   "parent node needs to be somewhere in a document");

        //  <Attendee>
        //    <Mailbox/>
        //    <ResponseType/>
        //    <LastResponseTime/>
        //  </Attendee>

        using namespace internal;

        auto& attendee_node = create_node(parent, "t:Attendee");
        mailbox_.to_xml_element(attendee_node);

        create_node(attendee_node, "t:ResponseType",
                    enum_to_str(response_type_));
        create_node(attendee_node, "t:LastResponseTime",
                    last_response_time_.to_string());

        return attendee_node;
    }

    //! Makes a attendee instance from an \<Attendee> XML element
    static attendee from_xml_element(const rapidxml::xml_node<>& elem)
    {
        using rapidxml::internal::compare;

        //  <Attendee>
        //    <Mailbox/>
        //    <ResponseType/>
        //    <LastResponseTime/>
        //  </Attendee>

        mailbox addr;
        auto resp_type = response_type::unknown;
        date_time last_resp_time("");

        for (auto node = elem.first_node(); node; node = node->next_sibling())
        {
            if (compare(node->local_name(), node->local_name_size(), "Mailbox",
                        std::strlen("Mailbox")))
            {
                addr = mailbox::from_xml_element(*node);
            }
            else if (compare(node->local_name(), node->local_name_size(),
                             "ResponseType", std::strlen("ResponseType")))
            {
                resp_type = internal::str_to_response_type(
                    std::string(node->value(), node->value_size()));
            }
            else if (compare(node->local_name(), node->local_name_size(),
                             "LastResponseTime",
                             std::strlen("LastResponseTime")))
            {
                last_resp_time = ews::date_time(
                    std::string(node->value(), node->value_size()));
            }
            else
            {
                throw exception("Unexpected child element in <Attendee>");
            }
        }

        return attendee(addr, resp_type, last_resp_time);
    }

private:
    mailbox mailbox_;
    response_type response_type_;
    date_time last_response_time_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<attendee>::value, "");
static_assert(std::is_copy_constructible<attendee>::value, "");
static_assert(std::is_copy_assignable<attendee>::value, "");
static_assert(std::is_move_constructible<attendee>::value, "");
static_assert(std::is_move_assignable<attendee>::value, "");
#endif

//! \brief Represents an <tt>\<InternetMessageHeader\></tt> property.
//!
//! An instance of this class describes a single name-value pair as it is
//! found in a message's header, essentially as defined in RFC 5322 and it's
//! former revisions.
//!
//! Most standard fields are already covered by EWS properties (e.g., the
//! destination address fields "To:", "Cc:", and "Bcc:"), however, because
//! users are allowed to define custom header fields as the seem fit, you
//! can directly access message headers.
//!
//! \sa message::get_internet_message_headers
class internet_message_header final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    internet_message_header() = delete;
#else
private:
    internet_message_header();

public:
#endif
    //! Constructs a header filed with given values
    internet_message_header(std::string name, std::string value)
        : header_name_(std::move(name)), header_value_(std::move(value))
    {
    }

    //! Returns the name of the header field
    const std::string& get_name() const EWS_NOEXCEPT { return header_name_; }

    //! Returns the value of the header field
    const std::string& get_value() const EWS_NOEXCEPT { return header_value_; }

private:
    std::string header_name_;
    std::string header_value_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<internet_message_header>::value,
              "");
static_assert(std::is_copy_constructible<internet_message_header>::value, "");
static_assert(std::is_copy_assignable<internet_message_header>::value, "");
static_assert(std::is_move_constructible<internet_message_header>::value, "");
static_assert(std::is_move_assignable<internet_message_header>::value, "");
#endif

namespace internal
{
    // Wraps a string so that it becomes its own type
    template <int tag> class str_wrapper final
    {
    public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
        str_wrapper() = default;
#else
        str_wrapper() : value_() {}
#endif

        explicit str_wrapper(std::string str) : value_(std::move(str)) {}

        const std::string& str() const EWS_NOEXCEPT { return value_; }

    private:
        std::string value_;
    };

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
    static_assert(std::is_default_constructible<str_wrapper<0>>::value, "");
    static_assert(std::is_copy_constructible<str_wrapper<0>>::value, "");
    static_assert(std::is_copy_assignable<str_wrapper<0>>::value, "");
    static_assert(std::is_move_constructible<str_wrapper<0>>::value, "");
    static_assert(std::is_move_assignable<str_wrapper<0>>::value, "");
#endif

    enum type_tags
    {
        distinguished_property_set_id,
        property_set_id,
        property_tag,
        property_name,
        property_id,
        property_type
    };
}

//! The ExtendedFieldURI element identifies an extended MAPI property
class extended_field_uri final
{
public:
#ifndef EWS_DOXYGEN_SHOULD_SKIP_THIS
#ifdef EWS_HAS_TYPE_ALIAS
    using distinguished_property_set_id =
        internal::str_wrapper<internal::distinguished_property_set_id>;
    using property_set_id = internal::str_wrapper<internal::property_set_id>;
    using property_tag = internal::str_wrapper<internal::property_tag>;
    using property_name = internal::str_wrapper<internal::property_name>;
    using property_id = internal::str_wrapper<internal::property_id>;
    using property_type = internal::str_wrapper<internal::property_type>;
#else
    typedef internal::str_wrapper<internal::distinguished_property_set_id>
        distinguished_property_set_id;
    typedef internal::str_wrapper<internal::property_set_id> property_set_id;
    typedef internal::str_wrapper<internal::property_tag> property_tag;
    typedef internal::str_wrapper<internal::property_name> property_name;
    typedef internal::str_wrapper<internal::property_id> property_id;
    typedef internal::str_wrapper<internal::property_type> property_type;
#endif
#endif

#ifdef EWS_HAS_DEFAULT_AND_DELETE
    extended_field_uri() = delete;
#else
private:
    extended_field_uri();

public:
#endif

    extended_field_uri(distinguished_property_set_id set_id, property_id id,
                       property_type type)
        : distinguished_set_id_(std::move(set_id)), id_(std::move(id)),
          type_(std::move(type))
    {
    }

    extended_field_uri(distinguished_property_set_id set_id, property_name name,
                       property_type type)
        : distinguished_set_id_(std::move(set_id)), name_(std::move(name)),
          type_(std::move(type))
    {
    }

    extended_field_uri(property_set_id set_id, property_id id,
                       property_type type)
        : set_id_(std::move(set_id)), id_(std::move(id)), type_(std::move(type))
    {
    }

    extended_field_uri(property_set_id set_id, property_name name,
                       property_type type)
        : set_id_(std::move(set_id)), name_(std::move(name)),
          type_(std::move(type))
    {
    }

    extended_field_uri(property_tag tag, property_type type)
        : tag_(std::move(tag)), type_(std::move(type))
    {
    }

    const std::string& get_distinguished_property_set_id() const EWS_NOEXCEPT
    {
        return distinguished_set_id_.str();
    }

    const std::string& get_property_set_id() const EWS_NOEXCEPT
    {
        return set_id_.str();
    }

    const std::string& get_property_tag() const EWS_NOEXCEPT
    {
        return tag_.str();
    }

    const std::string& get_property_name() const EWS_NOEXCEPT
    {
        return name_.str();
    }

    const std::string& get_property_id() const EWS_NOEXCEPT
    {
        return id_.str();
    }

    const std::string& get_property_type() const EWS_NOEXCEPT
    {
        return type_.str();
    }

    //! Returns a string representation of this extended_field_uri
    std::string to_xml()
    {
        std::stringstream sstr;

        sstr << "<t:ExtendedFieldURI ";

        if (!distinguished_set_id_.str().empty())
        {
            sstr << "DistinguishedPropertySetId=\""
                 << distinguished_set_id_.str() << "\" ";
        }
        if (!id_.str().empty())
        {
            sstr << "PropertyId=\"" << id_.str() << "\" ";
        }
        if (!set_id_.str().empty())
        {
            sstr << "PropertySetId=\"" << set_id_.str() << "\" ";
        }
        if (!tag_.str().empty())
        {
            sstr << "PropertyTag=\"" << tag_.str() << "\" ";
        }
        if (!name_.str().empty())
        {
            sstr << "PropertyName=\"" << name_.str() << "\" ";
        }
        if (!type_.str().empty())
        {
            sstr << "PropertyType=\"" << type_.str() << "\"/>";
        }
        return sstr.str();
    }

    //! Converts an xml string into a extended_field_uri property.
    static extended_field_uri from_xml_element(const rapidxml::xml_node<>& elem)
    {
        using rapidxml::internal::compare;

        EWS_ASSERT(compare(elem.name(), elem.name_size(), "t:ExtendedFieldURI",
                           std::strlen("t:ExtendedFieldURI")) &&
                   "Expected a <ExtendedFieldURI/>, got something else");

        std::string distinguished_set_id;
        std::string set_id;
        std::string tag;
        std::string name;
        std::string id;
        std::string type;

        for (auto attr = elem.first_attribute(); attr != nullptr;
             attr = attr->next_attribute())
        {
            if (compare(attr->name(), attr->name_size(),
                        "DistinguishedPropertySetId",
                        std::strlen("DistinguishedPropertySetId")))
            {
                distinguished_set_id =
                    std::string(attr->value(), attr->value_size());
            }
            else if (compare(attr->name(), attr->name_size(), "PropertySetId",
                             std::strlen("PropertySetId")))
            {
                set_id = std::string(attr->value(), attr->value_size());
            }
            else if (compare(attr->name(), attr->name_size(), "PropertyTag",
                             std::strlen("PropertyTag")))
            {
                tag = std::string(attr->value(), attr->value_size());
            }
            else if (compare(attr->name(), attr->name_size(), "PropertyName",
                             std::strlen("PropertyName")))
            {
                name = std::string(attr->value(), attr->value_size());
            }
            else if (compare(attr->name(), attr->name_size(), "PropertyId",
                             std::strlen("PropertyId")))
            {
                id = std::string(attr->value(), attr->value_size());
            }
            else if (compare(attr->name(), attr->name_size(), "PropertyType",
                             std::strlen("PropertyType")))
            {
                type = std::string(attr->value(), attr->value_size());
            }
            else
            {
                throw exception("Unexpected attribute in <ExtendedFieldURI>");
            }
        }

        EWS_ASSERT(!type.empty() && "Type attribute missing");

        if (!distinguished_set_id.empty())
        {
            if (!id.empty())
            {
                return extended_field_uri(
                    distinguished_property_set_id(distinguished_set_id),
                    property_id(id), property_type(type));
            }
            else if (!name.empty())
            {
                return extended_field_uri(
                    distinguished_property_set_id(distinguished_set_id),
                    property_name(name), property_type(type));
            }
        }
        else if (!set_id.empty())
        {
            if (!id.empty())
            {
                return extended_field_uri(property_set_id(set_id),
                                          property_id(id), property_type(type));
            }
            else if (!name.empty())
            {
                return extended_field_uri(property_set_id(set_id),
                                          property_name(name),
                                          property_type(type));
            }
        }
        else if (!tag.empty())
        {
            return extended_field_uri(property_tag(tag), property_type(type));
        }

        throw exception(
            "Unexpected combination of <ExtendedFieldURI/> attributes");
    }

    rapidxml::xml_node<>& to_xml_element(rapidxml::xml_node<>& parent) const
    {
        auto doc = parent.document();

        auto ptr_to_node = doc->allocate_string("t:ExtendedFieldURI");
        auto new_node = doc->allocate_node(rapidxml::node_element, ptr_to_node);
        new_node->namespace_uri(internal::uri<>::microsoft::types(),
                                internal::uri<>::microsoft::types_size);

        if (!get_distinguished_property_set_id().empty())
        {
            auto ptr_to_attr =
                doc->allocate_string("DistinguishedPropertySetId");
            auto ptr_to_property = doc->allocate_string(
                get_distinguished_property_set_id().c_str());
            auto attr_new =
                doc->allocate_attribute(ptr_to_attr, ptr_to_property);
            new_node->append_attribute(attr_new);
        }

        if (!get_property_set_id().empty())
        {
            auto ptr_to_attr = doc->allocate_string("PropertySetId");
            auto ptr_to_property =
                doc->allocate_string(get_property_set_id().c_str());
            auto attr_new =
                doc->allocate_attribute(ptr_to_attr, ptr_to_property);
            new_node->append_attribute(attr_new);
        }

        if (!get_property_tag().empty())
        {
            auto ptr_to_attr = doc->allocate_string("PropertyTag");
            auto ptr_to_property =
                doc->allocate_string(get_property_tag().c_str());
            auto attr_new =
                doc->allocate_attribute(ptr_to_attr, ptr_to_property);
            new_node->append_attribute(attr_new);
        }

        if (!get_property_name().empty())
        {
            auto ptr_to_attr = doc->allocate_string("PropertyName");
            auto ptr_to_property =
                doc->allocate_string(get_property_name().c_str());
            auto attr_new =
                doc->allocate_attribute(ptr_to_attr, ptr_to_property);
            new_node->append_attribute(attr_new);
        }

        if (!get_property_type().empty())
        {
            auto ptr_to_attr = doc->allocate_string("PropertyType");
            auto ptr_to_property =
                doc->allocate_string(get_property_type().c_str());
            auto attr_new =
                doc->allocate_attribute(ptr_to_attr, ptr_to_property);
            new_node->append_attribute(attr_new);
        }

        if (!get_property_id().empty())
        {
            auto ptr_to_attr = doc->allocate_string("PropertyId");
            auto ptr_to_property =
                doc->allocate_string(get_property_id().c_str());
            auto attr_new =
                doc->allocate_attribute(ptr_to_attr, ptr_to_property);
            new_node->append_attribute(attr_new);
        }

        parent.append_node(new_node);
        return *new_node;
    }

private:
    distinguished_property_set_id distinguished_set_id_;
    property_set_id set_id_;
    property_tag tag_;
    property_name name_;
    property_id id_;
    property_type type_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<extended_field_uri>::value, "");
static_assert(std::is_copy_constructible<extended_field_uri>::value, "");
static_assert(std::is_copy_assignable<extended_field_uri>::value, "");
static_assert(std::is_move_constructible<extended_field_uri>::value, "");
static_assert(std::is_move_assignable<extended_field_uri>::value, "");
#endif

//! \brief Represents an <tt>\<ExtendedProperty\><tt>
//
//! The ExtendedProperty element identifies extended MAPI properties on
//! folders and items. Extended properties enable Microsoft Exchange Server
//! clients to add customized properties to items and folders that are
//! stored in an Exchange mailbox. Custom properties can be used to store
//! data that is relevant to an object.
class extended_property final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    extended_property() = delete;
#else
private:
    extended_property();

public:
#endif
    //! \brief Constructor to initialize an <tt>\<ExtendedProperty\></tt>
    //! with the
    //! neccessary values
    extended_property(extended_field_uri ext_field_uri,
                      std::vector<std::string> values)
        : extended_field_uri_(std::move(ext_field_uri)),
          values_(std::move(values))
    {
    }

    //! Returns the the extended_field_uri element of this
    //! extended_property.
    const extended_field_uri& get_extended_field_uri() const EWS_NOEXCEPT
    {
        return extended_field_uri_;
    }

    //! \brief Returns the values of the extended_property as a vector
    //! even it is just one
    const std::vector<std::string>& get_values() const EWS_NOEXCEPT
    {
        return values_;
    }

private:
    extended_field_uri extended_field_uri_;
    std::vector<std::string> values_;
};
#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<extended_property>::value, "");
static_assert(std::is_copy_constructible<extended_property>::value, "");
static_assert(std::is_copy_assignable<extended_property>::value, "");
static_assert(std::is_move_constructible<extended_property>::value, "");
static_assert(std::is_move_assignable<extended_property>::value, "");
#endif

//! Represents a generic <tt>\<Item></tt> in the Exchange store
class item
{
public:
//! Constructs a new item
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    item() = default;
#else
    item() {}
#endif
    //! Constructs a new item with the given id
    explicit item(item_id id) : item_id_(std::move(id)), xml_subtree_() {}

#ifndef EWS_DOXYGEN_SHOULD_SKIP_THIS
    item(item_id&& id, internal::xml_subtree&& properties)
        : item_id_(std::move(id)), xml_subtree_(std::move(properties))
    {
    }
#endif

    //! Returns the id of an item
    const item_id& get_item_id() const EWS_NOEXCEPT { return item_id_; }

    //! Base64-encoded contents of the MIME stream of this item
    mime_content get_mime_content() const
    {
        const auto node = xml().get_node("MimeContent");
        if (!node)
        {
            return mime_content();
        }
        auto charset = node->first_attribute("CharacterSet");
        EWS_ASSERT(charset &&
                   "Expected <MimeContent> to have CharacterSet attribute");
        return mime_content(charset->value(), node->value(),
                            node->value_size());
    }

    //! \brief Returns a unique identifier for the folder that contains
    //! this item
    //!
    //! This is a read-only property.
    folder_id get_parent_folder_id() const
    {
        const auto node = xml().get_node("ParentFolderId");
        return node ? folder_id::from_xml_element(*node) : folder_id();
    }

    //! \brief Returns the PR_MESSAGE_CLASS MAPI property (the message
    //! class) for an item
    std::string get_item_class() const
    {
        return xml().get_value_as_string("ItemClass");
    }

    //! Sets this item's subject. Limited to 255 characters.
    void set_subject(const std::string& subject)
    {
        xml().set_or_update("Subject", subject);
    }

    //! Returns this item's subject
    std::string get_subject() const
    {
        return xml().get_value_as_string("Subject");
    }

    //! Returns the sensitivity level of this item
    sensitivity get_sensitivity() const
    {
        const auto val = xml().get_value_as_string("Sensitivity");
        return !val.empty() ? internal::str_to_sensitivity(val)
                            : sensitivity::normal;
    }

    //! Sets the sensitivity level of this item
    void set_sensitivity(sensitivity s)
    {
        xml().set_or_update("Sensitivity", internal::enum_to_str(s));
    }

    //! Sets the body of this item
    void set_body(const body& b)
    {
        auto doc = xml().document();

        auto body_node = xml().get_node("Body");
        if (body_node)
        {
            doc->remove_node(body_node);
        }

        body_node = &internal::create_node(*doc, "t:Body", b.content());

        auto ptr_to_key = doc->allocate_string("BodyType");
        auto ptr_to_value =
            doc->allocate_string(internal::body_type_str(b.type()).c_str());
        body_node->append_attribute(
            doc->allocate_attribute(ptr_to_key, ptr_to_value));

        bool truncated = b.is_truncated();
        if (truncated)
        {
            ptr_to_key = doc->allocate_string("IsTruncated");
            ptr_to_value = doc->allocate_string(truncated ? "true" : "false");
            body_node->append_attribute(
                doc->allocate_attribute(ptr_to_key, ptr_to_value));
        }
    }

    //! Returns the body of this item
    body get_body() const
    {
        using rapidxml::internal::compare;

        body b;
        auto body_node = xml().get_node("Body");
        if (body_node)
        {
            for (auto attr = body_node->first_attribute(); attr;
                 attr = attr->next_attribute())
            {
                if (compare(attr->name(), attr->name_size(), "BodyType",
                            std::strlen("BodyType")))
                {
                    if (compare(attr->value(), attr->value_size(), "HTML",
                                std::strlen("HTML")))
                    {
                        b.set_type(body_type::html);
                    }
                    else if (compare(attr->value(), attr->value_size(), "Text",
                                     std::strlen("Text")))
                    {
                        b.set_type(body_type::plain_text);
                    }
                    else if (compare(attr->value(), attr->value_size(), "Best",
                                     std::strlen("Best")))
                    {
                        b.set_type(body_type::best);
                    }
                    else
                    {
                        EWS_ASSERT(false &&
                                   "Unexpected attribute value for BodyType");
                    }
                }
                else if (compare(attr->name(), attr->name_size(), "IsTruncated",
                                 std::strlen("IsTruncated")))
                {
                    const auto val = compare(attr->value(), attr->value_size(),
                                             "true", std::strlen("true"));
                    b.set_truncated(val);
                }
                else
                {
                    EWS_ASSERT(false &&
                               "Unexpected attribute in <Body> element");
                }
            }

            auto content =
                std::string(body_node->value(), body_node->value_size());
            b.set_content(std::move(content));
        }
        return b;
    }

    //! Returns the items or files that are attached to this item
    std::vector<attachment> get_attachments() const
    {
        // TODO: support attachment hierarchies

        const auto attachments_node = xml().get_node("Attachments");
        if (!attachments_node)
        {
            return std::vector<attachment>();
        }

        std::vector<attachment> attachments;
        for (auto child = attachments_node->first_node(); child != nullptr;
             child = child->next_sibling())
        {
            attachments.emplace_back(attachment::from_xml_element(*child));
        }
        return attachments;
    }

    //! \brief Date/Time an item was received.
    //!
    //! This is a read-only property.
    date_time get_date_time_received() const
    {
        const auto val = xml().get_value_as_string("DateTimeReceived");
        return !val.empty() ? date_time(val) : date_time();
    }

    //! \brief Size in bytes of an item.
    //!
    //! This is a read-only property. Default: 0
    std::size_t get_size() const
    {
        const auto size = xml().get_value_as_string("Size");
        return size.empty() ? 0U : std::stoul(size);
    }

    //! \brief Sets this item's categories.
    //!
    //! A category is a short user-defined string that groups items with
    //! the same category together. An item can have none or multiple
    //! categories assigned. Think of tags or Google Mail labels.
    //!
    //! \sa item::get_categories
    void set_categories(const std::vector<std::string>& categories)
    {
        auto doc = xml().document();
        auto target_node = xml().get_node("Categories");
        if (!target_node)
        {
            target_node = &internal::create_node(*doc, "t:Categories");
        }

        for (const auto& category : categories)
        {
            internal::create_node(*target_node, "t:String", category);
        }
    }

    //! \brief Returns the categories associated with this item.
    //!
    //! \sa item::set_categories
    std::vector<std::string> get_categories() const
    {
        const auto categories_node = xml().get_node("Categories");
        if (!categories_node)
        {
            return std::vector<std::string>();
        }

        std::vector<std::string> categories;
        for (auto child = categories_node->first_node(); child != nullptr;
             child = child->next_sibling())
        {
            categories.emplace_back(
                std::string(child->value(), child->value_size()));
        }
        return categories;
    }

    //! Returns the importance of this item
    importance get_importance() const
    {
        const auto val = xml().get_value_as_string("Importance");
        return !val.empty() ? internal::str_to_importance(val)
                            : importance::normal;
    }

    //! \brief Returns the identifier of the item to which this item is a
    //! reply.
    //!
    //! This is a read-only property.
    std::string get_in_reply_to() const
    {
        return xml().get_value_as_string("InReplyTo");
    }

    //! \brief True if this item has been submitted for delivery.
    //!
    //! Default: false.
    bool is_submitted() const
    {
        return xml().get_value_as_string("isSubmitted") == "true";
    }

    //! \brief True if this item is a draft.
    //!
    //! Default: false
    bool is_draft() const
    {
        return xml().get_value_as_string("isDraft") == "true";
    }

    //! \brief True if this item is from you.
    //!
    //! Default: false.
    bool is_from_me() const
    {
        return xml().get_value_as_string("isFromMe") == "true";
    }

    //! \brief True if this item a re-send.
    //!
    //! Default: false
    bool is_resend() const
    {
        return xml().get_value_as_string("isResend") == "true";
    }

    //! \brief True if this item is unmodified.
    //!
    //! Default: false.
    bool is_unmodified() const
    {
        return xml().get_value_as_string("isUnmodified") == "true";
    }

    //! \brief Returns a collection of Internet message headers associated
    //! with this item.
    //!
    //! This is a read-only property.
    //!
    //! \sa internet_message_header
    std::vector<internet_message_header> get_internet_message_headers() const
    {
        const auto node = xml().get_node("InternetMessageHeaders");
        if (!node)
        {
            return std::vector<internet_message_header>();
        }

        std::vector<internet_message_header> headers;
        for (auto child = node->first_node(); child != nullptr;
             child = child->next_sibling())
        {
            auto field = std::string(child->first_attribute()->value(),
                                     child->first_attribute()->value_size());
            auto value = std::string(child->value(), child->value_size());

            headers.emplace_back(
                internet_message_header(std::move(field), std::move(value)));
        }

        return headers;
    }

    //! \brief Returns the date/time this item was sent.
    //!
    //! This is a read-only property.
    date_time get_date_time_sent() const
    {
        return date_time(xml().get_value_as_string("DateTimeSent"));
    }

    //! \brief Returns the date/time this item was created.
    //!
    //! This is a read-only property.
    date_time get_date_time_created() const
    {
        return date_time(xml().get_value_as_string("DateTimeCreated"));
    }

    // Applicable response actions for this item
    // (NonEmptyArrayOfResponseObjectsType)
    // TODO: get_response_objects

    //! \brief Sets the due date of this item.
    //!
    //! Used for reminders.
    void set_reminder_due_by(const date_time& due_by)
    {
        xml().set_or_update("ReminderDueBy", due_by.to_string());
    }

    //! \brief Returns the due date of this item
    //!
    //! \sa item::set_reminder_due_by
    date_time get_reminder_due_by() const
    {
        return date_time(xml().get_value_as_string("ReminderDueBy"));
    }

    //! Set a reminder on this item
    void set_reminder_enabled(bool enabled)
    {
        xml().set_or_update("ReminderIsSet", enabled ? "true" : "false");
    }

    //! True if a reminder has been enabled on this item
    bool is_reminder_enabled() const
    {
        return xml().get_value_as_string("ReminderIsSet") == "true";
    }

    //! \brief Sets the minutes before due date that a reminder should be
    //! shown to the user.
    void set_reminder_minutes_before_start(std::uint32_t minutes)
    {
        xml().set_or_update("ReminderMinutesBeforeStart",
                            std::to_string(minutes));
    }

    //! \brief Returns the number of minutes before due date that a
    //! reminder should be shown to the user.
    std::uint32_t get_reminder_minutes_before_start() const
    {
        std::string minutes =
            xml().get_value_as_string("ReminderMinutesBeforeStart");
        return minutes.empty() ? 0U : std::stoul(minutes);
    }

    //! \brief Returns a nice string containing all Cc: recipients of this
    //! item.
    //!
    //! The \<DisplayCc/> property is a concatenated string of the display
    //! names of the Cc: recipients of an item. Each recipient is separated
    //! by a semicolon. This is a read-only property.
    std::string get_display_cc() const
    {
        return xml().get_value_as_string("DisplayCc");
    }

    //! \brief Returns a nice string containing all To: recipients of this
    //! item.
    //!
    //! The \<DisplayTo/> property is a concatenated string of the display
    //! names of all the To: recipients of an item. Each recipient is
    //! separated by a semicolon. This is a read-only property.
    std::string get_display_to() const
    {
        return xml().get_value_as_string("DisplayTo");
    }

    //! \brief True if this item has non-hidden attachments.
    //!
    //! This is a read-only property.
    bool has_attachments() const
    {
        return xml().get_value_as_string("HasAttachments") == "true";
    }

    //! List of zero or more extended properties that are requested for
    //! an item
    std::vector<extended_property> get_extended_properties() const
    {
        using rapidxml::internal::compare;

        std::vector<extended_property> properties;

        for (auto top_node = xml().get_node("ExtendedProperty");
             top_node != nullptr; top_node = top_node->next_sibling())
        {
            if (!compare(top_node->name(), top_node->name_size(),
                         "t:ExtendedProperty",
                         std::strlen("t:ExtendedProperty")))
            {
                continue;
            }

            for (auto node = top_node->first_node(); node != nullptr;
                 node = node->next_sibling())
            {
                auto ext_field_uri =
                    extended_field_uri::from_xml_element(*node);

                std::vector<std::string> values;
                node = node->next_sibling(); // go to value(es)

                if (compare(node->name(), node->name_size(), "t:Value",
                            std::strlen("t:Value")))
                {
                    values.emplace_back(
                        std::string(node->value(), node->value_size()));
                }
                else if (compare(node->name(), node->name_size(), "t:Values",
                                 std::strlen("t:Values")))
                {
                    for (auto child = node->first_node(); child != nullptr;
                         child = child->next_sibling())
                    {
                        values.emplace_back(
                            std::string(child->value(), child->value_size()));
                    }
                }

                properties.emplace_back(extended_property(
                    std::move(ext_field_uri), std::move(values)));
            }
        }

        return properties;
    }

    //! Sets an extended property of an item
    void set_extended_property(const extended_property& extended_prop)
    {
        auto doc = xml().document();

        auto ptr_to_qname = doc->allocate_string("t:ExtendedProperty");
        auto top_node = doc->allocate_node(rapidxml::node_element);
        top_node->qname(ptr_to_qname, std::strlen("t:ExtendedProperty"),
                        ptr_to_qname + 2);
        top_node->namespace_uri(internal::uri<>::microsoft::types(),
                                internal::uri<>::microsoft::types_size);
        doc->append_node(top_node);

        extended_field_uri field_uri = extended_prop.get_extended_field_uri();
        field_uri.to_xml_element(*top_node);

        rapidxml::xml_node<>* cover_node = nullptr;
        if (extended_prop.get_values().size() > 1)
        {
            auto ptr_to_values = doc->allocate_string("t:Values");
            cover_node =
                doc->allocate_node(rapidxml::node_element, ptr_to_values);
            cover_node->namespace_uri(internal::uri<>::microsoft::types(),
                                      internal::uri<>::microsoft::types_size);
            top_node->append_node(cover_node);
        }

        for (auto str : extended_prop.get_values())
        {
            auto new_str = doc->allocate_string(str.c_str());
            auto ptr_to_value = doc->allocate_string("t:Value");
            auto cur_node =
                doc->allocate_node(rapidxml::node_element, ptr_to_value);
            cur_node->namespace_uri(internal::uri<>::microsoft::types(),
                                    internal::uri<>::microsoft::types_size);
            cur_node->value(new_str);

            cover_node == nullptr ? top_node->append_node(cur_node)
                                  : cover_node->append_node(cur_node);
        }
    }

    //! Sets the culture name associated with the body of this item
    void set_culture(const std::string& culture)
    {
        xml().set_or_update("Culture", culture);
    }

    //! Returns the culture name associated with the body of this item
    std::string get_culture() const
    {
        return xml().get_value_as_string("Culture");
    }

// Following properties are beyond 2007 scope:
//   <EffectiveRights/>
//   <LastModifiedName/>
//   <LastModifiedTime/>
//   <IsAssociated/>
//   <WebClientReadFormQueryString/>
//   <WebClientEditFormQueryString/>
//   <ConversationId/>
//   <UniqueBody/>

#ifndef EWS_DOXYGEN_SHOULD_SKIP_THIS
protected:
    internal::xml_subtree& xml() EWS_NOEXCEPT { return xml_subtree_; }

    const internal::xml_subtree& xml() const EWS_NOEXCEPT
    {
        return xml_subtree_;
    }
#endif

    void set_array_of_strings_helper(const std::vector<std::string>& strings,
                                     const char* name)
    {
        // TODO: this does not meet strong exception safety guarantees

        using namespace internal;

        auto outer_node = xml().get_node(name);
        if (outer_node)
        {
            auto doc = outer_node->document();
            doc->remove_node(outer_node);
        }

        if (strings.empty())
        {
            // Nothing to do
            return;
        }

        outer_node = &create_node(*xml().document(), std::string("t:") + name);
        for (const auto& element : strings)
        {
            create_node(*outer_node, "t:String", element);
        }
    }

    std::vector<std::string> get_array_of_strings_helper(const char* name) const
    {
        auto node = xml().get_node(name);
        if (!node)
        {
            return std::vector<std::string>();
        }
        auto res = std::vector<std::string>();
        for (auto entry = node->first_node(); entry;
             entry = entry->next_sibling())
        {
            res.emplace_back(std::string(entry->value(), entry->value_size()));
        }
        return res;
    }

private:
    friend class attachment;

    item_id item_id_;
    internal::xml_subtree xml_subtree_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(std::is_default_constructible<item>::value, "");
static_assert(std::is_copy_constructible<item>::value, "");
static_assert(std::is_copy_assignable<item>::value, "");
static_assert(std::is_move_constructible<item>::value, "");
static_assert(std::is_move_assignable<item>::value, "");
#endif

//! \brief Describes the state of a delegated task.
//!
//! Values indicate whether the delegated task was accepted or not.
enum class delegation_state
{
    //! \brief The task is not a delegated task, or the task request has
    //! been created but not sent
    no_match,

    //! \brief The task is new and the request has been sent, but the
    //! delegate has no yet responded to the task
    own_new,

    //! Should not be used
    owned,

    //! The task was accepted by the delegate
    accepted,

    //! The task was declined by the delegate
    declined,

    //! Should not be used
    max
};

namespace internal
{
    inline std::string enum_to_str(delegation_state state)
    {
        switch (state)
        {
        case delegation_state::no_match:
            return "NoMatch";
        case delegation_state::own_new:
            return "OwnNew";
        case delegation_state::owned:
            return "Owned";
        case delegation_state::accepted:
            return "Accepted";
        case delegation_state::declined:
            return "Declined";
        case delegation_state::max:
            return "Max";
        default:
            throw exception("Unexpected <DelegationState>");
        }
    }
}

//! Specifies the status of a task item
enum class status
{
    //! The task is not started
    not_started,

    //! The task is started and in progress
    in_progress,

    //! The task is completed
    completed,

    //! The task is waiting on other
    waiting_on_others,

    //! The task is deferred
    deferred
};

namespace internal
{
    inline std::string enum_to_str(status s)
    {
        switch (s)
        {
        case status::not_started:
            return "NotStarted";
        case status::in_progress:
            return "InProgress";
        case status::completed:
            return "Completed";
        case status::waiting_on_others:
            return "WaitingOnOthers";
        case status::deferred:
            return "Deferred";
        default:
            throw exception("Unexpected <Status>");
        }
    }
}

//! Describes the month when a yearly recurring item occurs
enum class month
{
    //! January
    jan,

    //! February
    feb,

    //! March
    mar,

    //! April
    apr,

    //! May
    may,

    //! June
    june,

    //! July
    july,

    //! August
    aug,

    //! September
    sept,

    //! October
    oct,

    //! November
    nov,

    //! December
    dec
};

namespace internal
{
    inline std::string enum_to_str(month m)
    {
        switch (m)
        {
        case month::jan:
            return "January";
        case month::feb:
            return "February";
        case month::mar:
            return "March";
        case month::apr:
            return "April";
        case month::may:
            return "May";
        case month::june:
            return "June";
        case month::july:
            return "July";
        case month::aug:
            return "August";
        case month::sept:
            return "September";
        case month::oct:
            return "October";
        case month::nov:
            return "November";
        case month::dec:
            return "December";
        default:
            throw exception("Unexpected <Month>");
        }
    }

    inline month str_to_month(const std::string& str)
    {
        month mon;
        if (str == "January")
        {
            mon = month::jan;
        }
        else if (str == "February")
        {
            mon = month::feb;
        }
        else if (str == "March")
        {
            mon = month::mar;
        }
        else if (str == "April")
        {
            mon = month::apr;
        }
        else if (str == "May")
        {
            mon = month::may;
        }
        else if (str == "June")
        {
            mon = month::june;
        }
        else if (str == "July")
        {
            mon = month::july;
        }
        else if (str == "August")
        {
            mon = month::aug;
        }
        else if (str == "September")
        {
            mon = month::sept;
        }
        else if (str == "October")
        {
            mon = month::oct;
        }
        else if (str == "November")
        {
            mon = month::nov;
        }
        else if (str == "December")
        {
            mon = month::dec;
        }
        else
        {
            throw exception("Unexpected <Month>");
        }
        return mon;
    }
}

//! Describes working days
enum class day_of_week
{
    //! Sunday
    sun,

    //! Monday
    mon,

    //! Tuesday
    tue,

    //! Wednesday
    wed,

    //! Thursday
    thu,

    //! Friday
    fri,

    //! Saturday
    sat,

    //! Any day
    day,

    //! A weekday
    weekday,

    //! A day weekend
    weekend_day
};

namespace internal
{
    inline std::string enum_to_str(day_of_week d)
    {
        switch (d)
        {
        case day_of_week::sun:
            return "Sunday";
        case day_of_week::mon:
            return "Monday";
        case day_of_week::tue:
            return "Tuesday";
        case day_of_week::wed:
            return "Wednesday";
        case day_of_week::thu:
            return "Thursday";
        case day_of_week::fri:
            return "Friday";
        case day_of_week::sat:
            return "Saturday";
        case day_of_week::day:
            return "Day";
        case day_of_week::weekday:
            return "Weekday";
        case day_of_week::weekend_day:
            return "WeekendDay";
        default:
            throw exception("Unexpected <DaysOfWeek>");
        }
    }

    inline day_of_week str_to_day_of_week(const std::string& str)
    {
        if (str == "Sunday")
        {
            return day_of_week::sun;
        }
        else if (str == "Monday")
        {
            return day_of_week::mon;
        }
        else if (str == "Tuesday")
        {
            return day_of_week::tue;
        }
        else if (str == "Wednesday")
        {
            return day_of_week::wed;
        }
        else if (str == "Thursday")
        {
            return day_of_week::thu;
        }
        else if (str == "Friday")
        {
            return day_of_week::fri;
        }
        else if (str == "Saturday")
        {
            return day_of_week::sat;
        }
        else if (str == "Day")
        {
            return day_of_week::day;
        }
        else if (str == "Weekday")
        {
            return day_of_week::weekday;
        }
        else if (str == "WeekendDay")
        {
            return day_of_week::weekend_day;
        }
        else
        {
            throw exception("Unexpected <DaysOfWeek>");
        }
    }
}

//! \brief This element describes which week in a month is used in a
//! relative recurrence pattern.
//!
//! For example, the second Monday of a month may occur in the third week
//! of that month. If a month starts on a Friday, the first week of the
//! month only contains a few days and does not contain a Monday.
//! Therefore, the first Monday would have to occur in the second week.
enum class day_of_week_index
{
    //! The first occurrence of a day within a month
    first,

    //! The second occurrence of a day within a month
    second,

    //! The third occurrence of a day within a month
    third,

    //! The fourth occurrence of a day within a month
    fourth,

    //! The last occurrence of a day within a month
    last
};

namespace internal
{
    inline std::string enum_to_str(day_of_week_index i)
    {
        switch (i)
        {
        case day_of_week_index::first:
            return "First";
        case day_of_week_index::second:
            return "Second";
        case day_of_week_index::third:
            return "Third";
        case day_of_week_index::fourth:
            return "Fourth";
        case day_of_week_index::last:
            return "Last";
        default:
            throw exception("Unexpected <DayOfWeekIndex>");
        }
    }

    inline day_of_week_index str_to_day_of_week_index(const std::string& str)
    {
        if (str == "First")
        {
            return day_of_week_index::first;
        }
        else if (str == "Second")
        {
            return day_of_week_index::second;
        }
        else if (str == "Third")
        {
            return day_of_week_index::third;
        }
        else if (str == "Fourth")
        {
            return day_of_week_index::fourth;
        }
        else if (str == "Last")
        {
            return day_of_week_index::last;
        }
        else
        {
            throw exception("Unexpected <DayOfWeekIndex>");
        }
    }
}

//! Represents a concrete task in the Exchange store
class task final : public item
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    task() = default;
#else
    task() {}
#endif

    //! Constructs a new task with the given item_id
    explicit task(item_id id) : item(std::move(id)) {}

#ifndef EWS_DOXYGEN_SHOULD_SKIP_THIS
    task(item_id&& id, internal::xml_subtree&& properties)
        : item(std::move(id), std::move(properties))
    {
    }
#endif

    //! \brief Returns the actual amount of work expended on the task.
    //!
    //! Measured in minutes.
    int get_actual_work() const
    {
        const auto val = xml().get_value_as_string("ActualWork");
        if (val.empty())
        {
            return 0;
        }
        return std::stoi(val);
    }

    //! \brief Sets the actual amount of work expended on the task.
    //!
    //! Measured in minutes.
    void set_actual_work(int actual_work)
    {
        xml().set_or_update("ActualWork", std::to_string(actual_work));
    }

    //! \brief Returns the time this task was assigned to the current
    //! owner.
    //!
    //! If this task is not a delegated task, this property is not set.
    //! This is a read-only property.
    date_time get_assigned_time() const
    {
        return date_time(xml().get_value_as_string("AssignedTime"));
    }

    //! Returns the billing information associated with this task
    std::string get_billing_information() const
    {
        return xml().get_value_as_string("BillingInformation");
    }

    //! Sets the billing information associated with this task
    void set_billing_information(const std::string& billing_info)
    {
        xml().set_or_update("BillingInformation", billing_info);
    }

    //! \brief Returns the change count of this task.
    //!
    //! How many times this task has been acted upon (sent, accepted,
    //! etc.). This is simply a way to resolve conflicts when the
    //! delegator sends multiple updates. Also known as TaskVersion
    //! Seems to be read-only.
    int get_change_count() const
    {
        const auto val = xml().get_value_as_string("ChangeCount");
        if (val.empty())
        {
            return 0;
        }
        return std::stoi(val);
    }

    //! \brief Returns the companies associated with this task.
    //!
    //! A list of company names associated with this task.
    //!
    //! Note: It seems that Exchange server accepts only one \<String>
    //! element here, although it is an ArrayOfStringsType.
    std::vector<std::string> get_companies() const
    {
        return get_array_of_strings_helper("Companies");
    }

    //! \brief Sets the companies associated with this task.
    //!
    //! Note: It seems that Exchange server accepts only one \<String>
    //! element here, although it is an ArrayOfStringsType.
    void set_companies(const std::vector<std::string>& companies)
    {
        set_array_of_strings_helper(companies, "Companies");
    }

    //! Returns the time the task was completed
    date_time get_complete_date() const
    {
        return date_time(xml().get_value_as_string("CompleteDate"));
    }

    //! Returns a list of contacts associated with this task
    std::vector<std::string> get_contacts() const
    {
        return get_array_of_strings_helper("Contacts");
    }

    //! Sets the contacts associated with this task to \p contacts.
    void set_contacts(const std::vector<std::string>& contacts)
    {
        set_array_of_strings_helper(contacts, "Contacts");
    }

    //! \brief Returns the delegation state of this task
    //!
    //! This is a read-only property.
    delegation_state get_delegation_state() const
    {
        const auto& val = xml().get_value_as_string("DelegationState");

        if (val.empty() || val == "NoMatch")
        {
            return delegation_state::no_match;
        }
        else if (val == "OwnNew")
        {
            return delegation_state::own_new;
        }
        else if (val == "Owned")
        {
            return delegation_state::owned;
        }
        else if (val == "Accepted")
        {
            return delegation_state::accepted;
        }
        else if (val == "Declined")
        {
            return delegation_state::declined;
        }
        else if (val == "Max")
        {
            return delegation_state::max;
        }
        else
        {
            throw exception("Unexpected <DelegationState>");
        }
    }

    //! Returns the name of the user that delegated the task
    std::string get_delegator() const
    {
        return xml().get_value_as_string("Delegator");
    }

    //! Sets the date that the task is due
    void set_due_date(const date_time& due_date)
    {
        xml().set_or_update("DueDate", due_date.to_string());
    }

    //! Returns the date that the task is due
    date_time get_due_date() const
    {
        return date_time(xml().get_value_as_string("DueDate"));
    }

// TODO
// This is a read-only property.
#if 0
        int is_assignment_editable() const
        {
            // Possible values:
            // 0 The default for all task items
            // 1 A task request
            // 2 A task acceptance from a recipient of a task request
            // 3 A task declination from a recipient of a task request
            // 4 An update to a previous task request
            // 5 Not used
        }
#endif

    //! \brief True if the task is marked as complete.
    //!
    //! This is a read-only property. See also
    //! task_property_path::percent_complete
    bool is_complete() const
    {
        return xml().get_value_as_string("IsComplete") == "true";
    }

    //! True if the task is recurring
    bool is_recurring() const
    {
        return xml().get_value_as_string("IsRecurring") == "true";
    }

    //! \brief True if the task is a team task.
    //!
    //! This is a read-only property.
    bool is_team_task() const
    {
        return xml().get_value_as_string("IsTeamTask") == "true";
    }

    //! \brief Returns the mileage associated with the task
    //!
    //! Potentially used for reimbursement purposes
    std::string get_mileage() const
    {
        return xml().get_value_as_string("Mileage");
    }

    //! Sets the mileage associated with the task
    void set_mileage(const std::string& mileage)
    {
        xml().set_or_update("Mileage", mileage);
    }

// TODO: Not in AllProperties shape in EWS 2013, investigate
// The name of the user who owns the task.
// This is a read-only property
#if 0
        std::string get_owner() const
        {
            return xml().get_value_as_string("Owner");
        }
#endif

    //! \brief Returns the percentage of the task that has been completed.
    //!
    //! Valid values are 0-100.
    int get_percent_complete() const
    {
        const auto val = xml().get_value_as_string("PercentComplete");
        if (val.empty())
        {
            return 0;
        }
        return std::stoi(val);
    }

    //! \brief Sets the percentage of the task that has been completed.
    //!
    //! Valid values are 0-100. Note that setting \<PercentComplete> to 100
    //! has the same effect as setting a \<CompleteDate> or \<Status> to
    //! ews::status::completed.
    //!
    //! See MSDN for more on this.
    void set_percent_complete(int value)
    {
        xml().set_or_update("PercentComplete", std::to_string(value));
    }

    // Used for recurring tasks
    // TODO: get_recurrence

    //! Set the date that work on the task should start
    void set_start_date(const date_time& start_date)
    {
        xml().set_or_update("StartDate", start_date.to_string());
    }

    //! Returns the date that work on the task should start
    date_time get_start_date() const
    {
        return date_time(xml().get_value_as_string("StartDate"));
    }

    //! Returns the status of the task
    status get_status() const
    {
        const auto& val = xml().get_value_as_string("Status");
        if (val == "NotStarted")
        {
            return status::not_started;
        }
        else if (val == "InProgress")
        {
            return status::in_progress;
        }
        else if (val == "Completed")
        {
            return status::completed;
        }
        else if (val == "WaitingOnOthers")
        {
            return status::waiting_on_others;
        }
        else if (val == "Deferred")
        {
            return status::deferred;
        }
        else
        {
            throw exception("Unexpected <Status>");
        }
    }

    //! Sets the status of the task to \p s.
    void set_status(status s)
    {
        xml().set_or_update("Status", internal::enum_to_str(s));
    }

    //! \brief Returns the status description.
    //!
    //! A localized string version of the status. Useful for display
    //! purposes. This is a read-only property.
    std::string get_status_description() const
    {
        return xml().get_value_as_string("StatusDescription");
    }

    //! Returns the total amount of work for this task
    int get_total_work() const
    {
        const auto val = xml().get_value_as_string("TotalWork");
        if (val.empty())
        {
            return 0;
        }
        return std::stoi(val);
    }

    //! Sets the total amount of work for this task
    void set_total_work(int total_work)
    {
        xml().set_or_update("TotalWork", std::to_string(total_work));
    }

    // Every property below is 2010 or 2013 dialect

    // TODO: add remaining properties

    //! Makes a task instance from a <tt>\<Task></tt> XML element
    static task from_xml_element(const rapidxml::xml_node<>& elem)
    {
        auto id_node =
            elem.first_node_ns(internal::uri<>::microsoft::types(), "ItemId");
        EWS_ASSERT(id_node && "Expected <ItemId>");
        return task(item_id::from_xml_element(*id_node),
                    internal::xml_subtree(elem));
    }

private:
    template <typename U> friend class basic_service;

    std::string create_item_request_string() const
    {
        std::stringstream sstr;
        sstr << "<m:CreateItem>"
                "<m:Items>"
                "<t:Task>"
             << xml().to_string() << "</t:Task>"
                                     "</m:Items>"
                                     "</m:CreateItem>";
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

class complete_name final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    complete_name() = default;
#else
    complete_name() {}
#endif

    complete_name(std::string title, std::string firstname,
                  std::string middlename, std::string lastname,
                  std::string suffix, std::string initials,
                  std::string fullname, std::string nickname)
        : title_(std::move(title)), firstname_(std::move(firstname)),
          middlename_(std::move(middlename)), lastname_(std::move(lastname)),
          suffix_(std::move(suffix)), initials_(std::move(initials)),
          fullname_(std::move(fullname)), nickname_(std::move(nickname))
    {
    }

    static complete_name from_xml_element(const rapidxml::xml_node<char>& node)
    {
        using namespace rapidxml::internal;
        std::string title;
        std::string firstname;
        std::string middlename;
        std::string lastname;
        std::string suffix;
        std::string initials;
        std::string fullname;
        std::string nickname;
        for (auto child = node.first_node(); child != nullptr;
             child = child->next_sibling())
        {
            if (compare("Title", std::strlen("Title"), child->local_name(),
                        child->local_name_size()))
            {
                title = std::string(child->value(), child->value_size());
            }
            else if (compare("FirstName", std::strlen("FirstName"),
                             child->local_name(), child->local_name_size()))
            {
                firstname = std::string(child->value(), child->value_size());
            }
            else if (compare("MiddleName", std::strlen("MiddleName"),
                             child->local_name(), child->local_name_size()))
            {
                middlename = std::string(child->value(), child->value_size());
            }
            else if (compare("LastName", std::strlen("LastName"),
                             child->local_name(), child->local_name_size()))
            {
                lastname = std::string(child->value(), child->value_size());
            }
            else if (compare("Suffix", std::strlen("Suffix"),
                             child->local_name(), child->local_name_size()))
            {
                suffix = std::string(child->value(), child->value_size());
            }
            else if (compare("Initials", std::strlen("Initials"),
                             child->local_name(), child->local_name_size()))
            {
                initials = std::string(child->value(), child->value_size());
            }
            else if (compare("FullName", std::strlen("FullName"),
                             child->local_name(), child->local_name_size()))
            {
                fullname = std::string(child->value(), child->value_size());
            }
            else if (compare("Nickname", std::strlen("Nickname"),
                             child->local_name(), child->local_name_size()))
            {
                nickname = std::string(child->value(), child->value_size());
            }
        }

        return complete_name(title, firstname, middlename, lastname, suffix,
                             initials, fullname, nickname);
    }

    const std::string& get_title() const EWS_NOEXCEPT { return title_; }
    const std::string& get_first_name() const EWS_NOEXCEPT
    {
        return firstname_;
    }
    const std::string& get_middle_name() const EWS_NOEXCEPT
    {
        return middlename_;
    }
    const std::string& get_last_name() const EWS_NOEXCEPT { return lastname_; }
    const std::string& get_suffix() const EWS_NOEXCEPT { return suffix_; }
    const std::string& get_initials() const EWS_NOEXCEPT { return initials_; }
    const std::string& get_full_name() const EWS_NOEXCEPT { return fullname_; }
    const std::string& get_nickname() const EWS_NOEXCEPT { return nickname_; }

private:
    std::string title_;
    std::string firstname_;
    std::string middlename_;
    std::string lastname_;
    std::string suffix_;
    std::string initials_;
    std::string fullname_;
    std::string nickname_;
};

class email_address final
{
public:
    enum class key
    {
        email_address_1,
        email_address_2,
        email_address_3
    };

    explicit email_address(key k, std::string value)
        : key_(std::move(k)), value_(std::move(value))
    {
    }

    static email_address
    from_xml_element(const rapidxml::xml_node<char>& node); // defined below
    std::string to_xml() const;
    key get_key() const EWS_NOEXCEPT { return key_; }
    const std::string& get_value() const EWS_NOEXCEPT { return value_; }

private:
    key key_;
    std::string value_;
    friend bool operator==(const email_address&, const email_address&);
};

namespace internal
{
    inline email_address::key
    str_to_email_address_key(const std::string& keystring)
    {
        email_address::key k;
        if (keystring == "EmailAddress1")
        {
            k = email_address::key::email_address_1;
        }
        else if (keystring == "EmailAddress2")
        {
            k = email_address::key::email_address_2;
        }
        else if (keystring == "EmailAddress3")
        {
            k = email_address::key::email_address_3;
        }
        else
        {
            throw exception("Unrecognized key: " + keystring);
        }
        return k;
    }

    inline std::string enum_to_str(email_address::key k)
    {
        switch (k)
        {
        case email_address::key::email_address_1:
            return "EmailAddress1";
        case email_address::key::email_address_2:
            return "EmailAddress2";
        case email_address::key::email_address_3:
            return "EmailAddress3";
        default:
            throw exception("Bad enum value");
        }
    }
}

inline email_address
email_address::from_xml_element(const rapidxml::xml_node<char>& node)
{
    using namespace internal;
    using rapidxml::internal::compare;

    // <t:EmailAddresses>
    //  <Entry Key="EmailAddress1">donald.duck@duckburg.de</Entry>
    //  <Entry Key="EmailAddress3">dragonmaster1999@yahoo.com</Entry>
    // </t:EmailAddresses>

    EWS_ASSERT(compare(node.local_name(), node.local_name_size(), "Entry",
                       std::strlen("Entry")));
    auto key = node.first_attribute("Key");
    EWS_ASSERT(key && "Expected attribute Key");
    return email_address(
        str_to_email_address_key(std::string(key->value(), key->value_size())),
        std::string(node.value(), node.value_size()));
}

inline std::string email_address::to_xml() const
{
    std::stringstream sstr;
    sstr << " <t:"
         << "EmailAddresses"
         << ">";
    sstr << " <t:Entry Key=";
    sstr << "\"" << internal::enum_to_str(get_key());
    sstr << "\">";
    sstr << get_value();
    sstr << "</t:Entry>";
    sstr << " </t:"
         << "EmailAddresses"
         << ">";
    return sstr.str();
}

inline bool operator==(const email_address& lhs, const email_address& rhs)
{
    return (lhs.key_ == rhs.key_) && (lhs.value_ == rhs.value_);
}

class physical_address final
{
public:
    enum class key
    {
        home,
        business,
        other
    };

    physical_address(key k, std::string street, std::string city,
                     std::string state, std::string cor,
                     std::string postal_code)
        : key_(std::move(k)), street_(std::move(street)),
          city_(std::move(city)), state_(std::move(state)),
          country_or_region_(std::move(cor)),
          postal_code_(std::move(postal_code))
    {
    }

    static physical_address
    from_xml_element(const rapidxml::xml_node<char>& node); // defined below
    std::string to_xml() const;
    key get_key() const EWS_NOEXCEPT { return key_; }
    const std::string& street() const EWS_NOEXCEPT { return street_; }
    const std::string& city() const EWS_NOEXCEPT { return city_; }
    const std::string& state() const EWS_NOEXCEPT { return state_; }
    const std::string& country_or_region() const EWS_NOEXCEPT
    {
        return country_or_region_;
    }
    const std::string& postal_code() const EWS_NOEXCEPT { return postal_code_; }

private:
    key key_;
    std::string street_;
    std::string city_;
    std::string state_;
    std::string country_or_region_;
    std::string postal_code_;

    friend bool operator==(const physical_address&, const physical_address&);
};

inline bool operator==(const physical_address& lhs, const physical_address& rhs)
{
    return (lhs.key_ == rhs.key_) && (lhs.street_ == rhs.street_) &&
           (lhs.city_ == rhs.city_) && (lhs.state_ == rhs.state_) &&
           (lhs.country_or_region_ == rhs.country_or_region_) &&
           (lhs.postal_code_ == rhs.postal_code_);
}

namespace internal
{
    inline physical_address::key
    string_to_physical_address_key(const std::string& keystring)
    {
        physical_address::key k;
        if (keystring == "Home")
        {
            k = physical_address::key::home;
        }
        else if (keystring == "Business")
        {
            k = physical_address::key::business;
        }
        else if (keystring == "Other")
        {
            k = physical_address::key::other;
        }
        else
        {
            throw exception("Unrecognized key: " + keystring);
        }
        return k;
    }

    inline std::string enum_to_str(physical_address::key k)
    {
        switch (k)
        {
        case physical_address::key::home:
            return "Home";
        case physical_address::key::business:
            return "Business";
        case physical_address::key::other:
            return "Other";
        default:
            throw exception("Bad enum value");
        }
    }
}

inline physical_address
physical_address::from_xml_element(const rapidxml::xml_node<char>& node)
{
    using namespace internal;
    using rapidxml::internal::compare;

    EWS_ASSERT(compare(node.local_name(), node.local_name_size(), "Entry",
                       std::strlen("Entry")));

    // <Entry Key="Home">
    //      <Street>
    //      <City>
    //      <State>
    //      <CountryOrRegion>
    //      <PostalCode>
    // </Entry>

    auto key_attr = node.first_attribute();
    EWS_ASSERT(key_attr);
    EWS_ASSERT(compare(key_attr->name(), key_attr->name_size(), "Key", 3));
    const key key = string_to_physical_address_key(key_attr->value());

    std::string street;
    std::string city;
    std::string state;
    std::string cor;
    std::string postal_code;

    for (auto child = node.first_node(); child != nullptr;
         child = child->next_sibling())
    {
        if (compare("Street", std::strlen("Street"), child->local_name(),
                    child->local_name_size()))
        {
            street = std::string(child->value(), child->value_size());
        }
        if (compare("City", std::strlen("City"), child->local_name(),
                    child->local_name_size()))
        {
            city = std::string(child->value(), child->value_size());
        }
        if (compare("State", std::strlen("State"), child->local_name(),
                    child->local_name_size()))
        {
            state = std::string(child->value(), child->value_size());
        }
        if (compare("CountryOrRegion", std::strlen("CountryOrRegion"),
                    child->local_name(), child->local_name_size()))
        {
            cor = std::string(child->value(), child->value_size());
        }
        if (compare("PostalCode", std::strlen("PostalCode"),
                    child->local_name(), child->local_name_size()))
        {
            postal_code = std::string(child->value(), child->value_size());
        }
    }
    return physical_address(key, street, city, state, cor, postal_code);
}

inline std::string physical_address::to_xml() const
{
    std::stringstream sstr;
    sstr << " <t:"
         << "PhysicalAddresses"
         << ">";
    sstr << " <t:Entry Key=";
    sstr << "\"" << internal::enum_to_str(get_key());
    sstr << "\">";
    if (!street().empty())
    {
        sstr << "<t:Street>";
        sstr << street();
        sstr << "</t:Street>";
    }
    if (!city().empty())
    {
        sstr << "<t:City>";
        sstr << city();
        sstr << "</t:City>";
    }
    if (!state().empty())
    {
        sstr << "<t:State>";
        sstr << state();
        sstr << "</t:State>";
    }
    if (!country_or_region().empty())
    {
        sstr << "<t:CountryOrRegion>";
        sstr << country_or_region();
        sstr << "</t:CountryOrRegion>";
    }
    if (!postal_code().empty())
    {
        sstr << "<t:PostalCode>";
        sstr << postal_code();
        sstr << "</t:PostalCode>";
    }
    sstr << "</t:Entry>";
    sstr << " </t:"
         << "PhysicalAddresses"
         << ">";
    return sstr.str();
}

namespace internal
{
    enum file_as_mapping
    {
        none,
        last_comma_first,
        first_space_last,
        company,
        last_comma_first_company,
        company_last_first,
        last_first,
        last_first_company,
        company_last_comma_first,
        last_first_suffix,
        last_space_first_company,
        company_last_space_first,
        last_space_first
    };

    inline file_as_mapping str_to_map(const std::string& maptype)
    {
        file_as_mapping map;

        if (maptype == "LastCommaFirst")
        {
            map = file_as_mapping::last_comma_first;
        }
        else if (maptype == "FirstSpaceLast")
        {
            map = file_as_mapping::first_space_last;
        }
        else if (maptype == "Company")
        {
            map = file_as_mapping::company;
        }
        else if (maptype == "LastCommaFirstCompany")
        {
            map = file_as_mapping::last_comma_first_company;
        }
        else if (maptype == "CompanyLastFirst")
        {
            map = file_as_mapping::company_last_first;
        }
        else if (maptype == "LastFirst")
        {
            map = file_as_mapping::last_first;
        }
        else if (maptype == "LastFirstCompany")
        {
            map = file_as_mapping::last_first_company;
        }
        else if (maptype == "CompanyLastCommaFirst")
        {
            map = file_as_mapping::company_last_comma_first;
        }
        else if (maptype == "LastFirstSuffix")
        {
            map = file_as_mapping::last_first_suffix;
        }
        else if (maptype == "LastSpaceFirstCompany")
        {
            map = file_as_mapping::last_space_first_company;
        }
        else if (maptype == "CompanyLastSpaceFirst")
        {
            map = file_as_mapping::company_last_space_first;
        }
        else if (maptype == "LastSpaceFirst")
        {
            map = file_as_mapping::last_space_first;
        }
        else if (maptype == "None" || maptype == "")
        {
            map = file_as_mapping::none;
        }
        else
        {
            throw exception(std::string("Unrecognized FileAsMapping Type: ") +
                            maptype);
        }
        return map;
    }

    inline std::string enum_to_str(internal::file_as_mapping maptype)
    {
        std::string mappingtype;
        switch (maptype)
        {
        case file_as_mapping::none:
            mappingtype = "None";
            break;
        case file_as_mapping::last_comma_first:
            mappingtype = "LastCommaFirst";
            break;
        case file_as_mapping::first_space_last:
            mappingtype = "FirstSpaceLast";
            break;
        case file_as_mapping::company:
            mappingtype = "Company";
            break;
        case file_as_mapping::last_comma_first_company:
            mappingtype = "LastCommaFirstCompany";
            break;
        case file_as_mapping::company_last_first:
            mappingtype = "CompanyLastFirst";
            break;
        case file_as_mapping::last_first:
            mappingtype = "LastFirst";
            break;
        case file_as_mapping::last_first_company:
            mappingtype = "LastFirstCompany";
            break;
        case file_as_mapping::company_last_comma_first:
            mappingtype = "CompanyLastCommaFirst";
            break;
        case file_as_mapping::last_first_suffix:
            mappingtype = "LastFirstSuffix";
            break;
        case file_as_mapping::last_space_first_company:
            mappingtype = "LastSpaceFirstCompany";
            break;
        case file_as_mapping::company_last_space_first:
            mappingtype = "CompanyLastSpaceFirst";
            break;
        case file_as_mapping::last_space_first:
            mappingtype = "LastSpaceFirst";
            break;
        default:
            break;
        }
        return mappingtype;
    }
}

class im_address final
{
public:
    enum class key
    {
        imaddress1,
        imaddress2,
        imaddress3
    };

    im_address(key k, std::string value)
        : key_(std::move(k)), value_(std::move(value))
    {
    }

    static im_address from_xml_element(const rapidxml::xml_node<char>& node);
    std::string to_xml() const;

    key get_key() const EWS_NOEXCEPT { return key_; }
    const std::string& get_value() const EWS_NOEXCEPT { return value_; }

private:
    key key_;
    std::string value_;
    friend bool operator==(const im_address&, const im_address&);
};

namespace internal
{
    inline std::string enum_to_str(im_address::key k)
    {
        switch (k)
        {
        case im_address::key::imaddress1:
            return "ImAddress1";
        case im_address::key::imaddress2:
            return "ImAddress2";
        case im_address::key::imaddress3:
            return "ImAddress3";
        default:
            throw exception("Bad enum value");
        }
    }

    inline im_address::key str_to_im_address_key(const std::string& keystring)
    {
        im_address::key k;
        if (keystring == "ImAddress1")
        {
            k = im_address::key::imaddress1;
        }
        else if (keystring == "ImAddress2")
        {
            k = im_address::key::imaddress2;
        }
        else if (keystring == "ImAddress3")
        {
            k = im_address::key::imaddress3;
        }
        else
        {
            throw exception("Unrecognized key: " + keystring);
        }
        return k;
    }
}

inline im_address
im_address::from_xml_element(const rapidxml::xml_node<char>& node)
{
    using namespace internal;
    using rapidxml::internal::compare;

    // <t:ImAddresses>
    //  <Entry Key="ImAddress1">WOWMLGPRO</Entry>
    //  <Entry Key="ImAddress2">xXSwaggerBoiXx</Entry>
    // </t:ImAddresses>

    EWS_ASSERT(compare(node.local_name(), node.local_name_size(), "Entry",
                       std::strlen("Entry")));
    auto key = node.first_attribute("Key");
    EWS_ASSERT(key && "Expected attribute Key");
    return im_address(
        str_to_im_address_key(std::string(key->value(), key->value_size())),
        std::string(node.value(), node.value_size()));
}

inline std::string im_address::to_xml() const
{
    std::stringstream sstr;
    sstr << " <t:"
         << "ImAddresses"
         << ">";
    sstr << " <t:Entry Key=";
    sstr << "\"" << internal::enum_to_str(key_);
    sstr << "\">";
    sstr << get_value();
    sstr << "</t:Entry>";
    sstr << " </t:"
         << "ImAddresses"
         << ">";
    return sstr.str();
}

inline bool operator==(const im_address& lhs, const im_address& rhs)
{
    return (lhs.key_ == rhs.key_) && (lhs.value_ == rhs.value_);
}

class phone_number final
{
public:
    enum class key
    {
        assistant_phone,
        business_fax,
        business_phone,
        business_phone_2,
        callback,
        car_phone,
        company_main_phone,
        home_fax,
        home_phone,
        home_phone_2,
        isdn,
        mobile_phone,
        other_fax,
        other_telephone,
        pager,
        primary_phone,
        radio_phone,
        telex,
        ttytdd_phone
    };

    phone_number(key k, std::string val)
        : key_(std::move(k)), value_(std::move(val))
    {
    }

    static phone_number from_xml_element(const rapidxml::xml_node<char>& node);
    std::string to_xml() const;
    key get_key() const EWS_NOEXCEPT { return key_; }
    const std::string& get_value() const EWS_NOEXCEPT { return value_; }

private:
    key key_;
    std::string value_;
    friend bool operator==(const phone_number&, const phone_number&);
};

namespace internal
{
    inline phone_number::key
    str_to_phone_number_key(const std::string& keystring)
    {
        phone_number::key k;
        if (keystring == "AssistantPhone")
        {
            k = phone_number::key::assistant_phone;
        }
        else if (keystring == "BusinessFax")
        {
            k = phone_number::key::business_fax;
        }
        else if (keystring == "BusinessPhone")
        {
            k = phone_number::key::business_phone;
        }
        else if (keystring == "BusinessPhone2")
        {
            k = phone_number::key::business_phone_2;
        }
        else if (keystring == "Callback")
        {
            k = phone_number::key::callback;
        }
        else if (keystring == "CarPhone")
        {
            k = phone_number::key::car_phone;
        }
        else if (keystring == "CompanyMainPhone")
        {
            k = phone_number::key::company_main_phone;
        }
        else if (keystring == "HomeFax")
        {
            k = phone_number::key::home_fax;
        }
        else if (keystring == "HomePhone")
        {
            k = phone_number::key::home_phone;
        }
        else if (keystring == "HomePhone2")
        {
            k = phone_number::key::home_phone_2;
        }
        else if (keystring == "Isdn")
        {
            k = phone_number::key::isdn;
        }
        else if (keystring == "MobilePhone")
        {
            k = phone_number::key::mobile_phone;
        }
        else if (keystring == "OtherFax")
        {
            k = phone_number::key::other_fax;
        }
        else if (keystring == "OtherTelephone")
        {
            k = phone_number::key::other_telephone;
        }
        else if (keystring == "Pager")
        {
            k = phone_number::key::pager;
        }
        else if (keystring == "PrimaryPhone")
        {
            k = phone_number::key::primary_phone;
        }
        else if (keystring == "RadioPhone")
        {
            k = phone_number::key::radio_phone;
        }
        else if (keystring == "Telex")
        {
            k = phone_number::key::telex;
        }
        else if (keystring == "TtyTddPhone")
        {
            k = phone_number::key::ttytdd_phone;
        }
        else
        {
            throw exception("Unrecognized key: " + keystring);
        }
        return k;
    }

    inline std::string enum_to_str(phone_number::key k)
    {
        switch (k)
        {
        case phone_number::key::assistant_phone:
            return "AssistantPhone";
        case phone_number::key::business_fax:
            return "BusinessFax";
        case phone_number::key::business_phone:
            return "BusinessPhone";
        case phone_number::key::business_phone_2:
            return "BusinessPhone2";
        case phone_number::key::callback:
            return "Callback";
        case phone_number::key::car_phone:
            return "CarPhone";
        case phone_number::key::company_main_phone:
            return "CompanyMainPhone";
        case phone_number::key::home_fax:
            return "HomeFax";
        case phone_number::key::home_phone:
            return "HomePhone";
        case phone_number::key::home_phone_2:
            return "HomePhone2";
        case phone_number::key::isdn:
            return "Isdn";
        case phone_number::key::mobile_phone:
            return "MobilePhone";
        case phone_number::key::other_fax:
            return "OtherFax";
        case phone_number::key::other_telephone:
            return "OtherTelephone";
        case phone_number::key::pager:
            return "Pager";
        case phone_number::key::primary_phone:
            return "PrimaryPhone";
        case phone_number::key::radio_phone:
            return "RadioPhone";
        case phone_number::key::telex:
            return "Telex";
        case phone_number::key::ttytdd_phone:
            return "TtyTddPhone";
        default:
            throw exception("Bad enum value");
        }
    }
}

inline phone_number
phone_number::from_xml_element(const rapidxml::xml_node<char>& node)
{
    using namespace internal;
    using rapidxml::internal::compare;

    // <t:PhoneNumbers>
    //  <Entry Key="AssistantPhone">0123456789</Entry>
    //  <Entry Key="BusinessFax">9876543210</Entry>
    // </t:PhoneNumbers>

    EWS_ASSERT(compare(node.local_name(), node.local_name_size(), "Entry",
                       std::strlen("Entry")));
    auto key = node.first_attribute("Key");
    EWS_ASSERT(key && "Expected attribute Key");
    return phone_number(
        str_to_phone_number_key(std::string(key->value(), key->value_size())),
        std::string(node.value(), node.value_size()));
}

inline std::string phone_number::to_xml() const
{
    std::stringstream sstr;
    sstr << " <t:"
         << "PhoneNumbers"
         << ">";
    sstr << " <t:Entry Key=";
    sstr << "\"" << internal::enum_to_str(key_);
    sstr << "\">";
    sstr << get_value();
    sstr << "</t:Entry>";
    sstr << " </t:"
         << "PhoneNumbers"
         << ">";
    return sstr.str();
}

inline bool operator==(const phone_number& lhs, const phone_number& rhs)
{
    return (lhs.key_ == rhs.key_) && (lhs.value_ == rhs.value_);
}

//! A contact item in the Exchange store.
class contact final : public item
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    contact() = default;
#else
    contact() {}
#endif

    explicit contact(item_id id) : item(id) {}

#ifndef EWS_DOXYGEN_SHOULD_SKIP_THIS
    contact(item_id&& id, internal::xml_subtree&& properties)
        : item(std::move(id), std::move(properties))
    {
    }
#endif

    //! How the name should be filed for display/sorting purposes
    void set_file_as(const std::string& fileas)
    {
        xml().set_or_update("FileAs", fileas);
    }

    std::string get_file_as() { return xml().get_value_as_string("FileAs"); }

    //! How the various parts of a contact's information interact to form
    //! the FileAs property value
    //! Overrides previously made FileAs settings
    void set_file_as_mapping(internal::file_as_mapping maptype)
    {
        auto mapping = internal::enum_to_str(maptype);
        xml().set_or_update("FileAsMapping", mapping);
    }

    internal::file_as_mapping get_file_as_mapping()
    {
        return internal::str_to_map(xml().get_value_as_string("FileAsMapping"));
    }

    //! Sets the name to display for a contact
    void set_display_name(const std::string& display_name)
    {
        xml().set_or_update("DisplayName", display_name);
    }

    //! Returns the displayed name of the contact
    std::string get_display_name() const
    {
        return xml().get_value_as_string("DisplayName");
    }

    //! Sets the name by which a person is known to `given_name`; often
    //! referred to as a person's first name
    void set_given_name(const std::string& given_name)
    {
        xml().set_or_update("GivenName", given_name);
    }

    //! Returns the person's first name
    std::string get_given_name() const
    {
        return xml().get_value_as_string("GivenName");
    }

    //! Set the Initials for the contact
    void set_initials(const std::string& initials)
    {
        xml().set_or_update("Initials", initials);
    }

    //! Returns the person's initials
    std::string get_initials() const
    {
        return xml().get_value_as_string("Initials");
    }

    //! Set the middle name for the contact
    void set_middle_name(const std::string& middle_name)
    {
        xml().set_or_update("MiddleName", middle_name);
    }

    //! Returns the middle name of the contact
    std::string get_middle_name() const
    {
        return xml().get_value_as_string("MiddleName");
    }

    //! Sets another name by which the contact is known
    void set_nickname(const std::string& nickname)
    {
        xml().set_or_update("Nickname", nickname);
    }

    //! Returns the nickname of the contact
    std::string get_nickname() const
    {
        return xml().get_value_as_string("Nickname");
    }

    //! A combination of several name fields in one convenient place
    complete_name get_complete_name() const
    {
        auto node = xml().get_node("CompleteName");
        if (node == nullptr)
        {
            return complete_name();
        }
        return complete_name::from_xml_element(*node);
    }

    //! Sets the company that the contact is affiliated with
    void set_company_name(const std::string& company_name)
    {
        xml().set_or_update("CompanyName", company_name);
    }

    //! Returns the comany of the contact
    std::string get_company_name() const
    {
        return xml().get_value_as_string("CompanyName");
    }

    //! A collection of email addresses for the contact
    std::vector<email_address> get_email_addresses() const
    {
        const auto addresses = xml().get_node("EmailAddresses");
        if (!addresses)
        {
            return std::vector<email_address>();
        }
        std::vector<email_address> result;
        for (auto entry = addresses->first_node(); entry != nullptr;
             entry = entry->next_sibling())
        {
            result.push_back(email_address::from_xml_element(*entry));
        }
        return result;
    }

    void set_email_address(const email_address& address)
    {
        using rapidxml::internal::compare;
        using internal::create_node;
        auto doc = xml().document();
        auto email_addresses = xml().get_node("EmailAddresses");

        if (email_addresses)
        {
            bool entry_exists = false;
            auto entry = email_addresses->first_node();
            for (; entry != nullptr; entry = entry->next_sibling())
            {
                auto key_attr = entry->first_attribute();
                EWS_ASSERT(key_attr);
                EWS_ASSERT(
                    compare(key_attr->name(), key_attr->name_size(), "Key", 3));
                const auto key = internal::enum_to_str(address.get_key());
                if (compare(key_attr->value(), key_attr->value_size(),
                            key.c_str(), key.size()))
                {
                    entry_exists = true;
                    break;
                }
            }
            if (entry_exists)
            {
                email_addresses->remove_node(entry);
            }
        }
        else
        {
            email_addresses = &create_node(*doc, "t:EmailAddresses");
        }

        // create entry & key
        const auto value = address.get_value();
        auto new_entry = &create_node(*email_addresses, "t:Entry", value);
        auto ptr_to_key = doc->allocate_string("Key");
        const auto key = internal::enum_to_str(address.get_key());
        auto ptr_to_value = doc->allocate_string(key.c_str());
        new_entry->append_attribute(
            doc->allocate_attribute(ptr_to_key, ptr_to_value));
    }

    //! A collection of mailing addresses for the contact
    std::vector<physical_address> get_physical_addresses() const
    {
        const auto addresses = xml().get_node("PhysicalAddresses");
        if (!addresses)
        {
            return std::vector<physical_address>();
        }
        std::vector<physical_address> result;
        for (auto entry = addresses->first_node(); entry != nullptr;
             entry = entry->next_sibling())
        {
            result.push_back(physical_address::from_xml_element(*entry));
        }
        return result;
    }

    void set_physical_address(const physical_address& address)
    {
        using rapidxml::internal::compare;
        using internal::create_node;
        auto doc = xml().document();
        auto addresses = xml().get_node("PhysicalAddresses");
        // <PhysicalAddresses>
        //   <Entry Key="Home">
        //     <Street>
        //     <City>
        //     <State>
        //     <CountryOrRegion>
        //     <PostalCode>
        //   <Entry/>
        //   <Entry Key="Business">
        //     <Street>
        //     <City>
        //     <State>
        //     <CountryOrRegion>
        //     <PostalCode>
        //   <Entry/>
        // <PhysicalAddresses/>

        if (addresses)
        {
            bool entry_exists = false;
            auto entry = addresses->first_node();
            for (; entry != nullptr; entry = entry->next_sibling())
            {
                auto key_attr = entry->first_attribute();
                EWS_ASSERT(key_attr);
                EWS_ASSERT(
                    compare(key_attr->name(), key_attr->name_size(), "Key", 3));
                const auto key = internal::enum_to_str(address.get_key());
                if (compare(key_attr->value(), key_attr->value_size(),
                            key.c_str(), key.size()))
                {
                    entry_exists = true;
                    break;
                }
            }
            if (entry_exists)
            {
                addresses->remove_node(entry);
            }
        }
        else
        {
            addresses = &create_node(*doc, "t:PhysicalAddresses");
        }

        // create entry & key
        auto new_entry = &create_node(*addresses, "t:Entry");

        auto ptr_to_key = doc->allocate_string("Key");
        auto keystr = internal::enum_to_str(address.get_key());
        auto ptr_to_value = doc->allocate_string(keystr.c_str());
        new_entry->append_attribute(
            doc->allocate_attribute(ptr_to_key, ptr_to_value));

        if (!address.street().empty())
        {
            create_node(*new_entry, "t:Street", address.street());
        }
        if (!address.city().empty())
        {
            create_node(*new_entry, "t:City", address.city());
        }
        if (!address.state().empty())
        {
            create_node(*new_entry, "t:State", address.state());
        }
        if (!address.country_or_region().empty())
        {
            create_node(*new_entry, "t:CountryOrRegion",
                        address.country_or_region());
        }
        if (!address.postal_code().empty())
        {
            create_node(*new_entry, "t:PostalCode", address.postal_code());
        }
    }

    // A collection of phone numbers for the contact
    void set_phone_number(const phone_number& number)
    {
        using rapidxml::internal::compare;
        using internal::create_node;
        auto doc = xml().document();
        auto phone_numbers = xml().get_node("PhoneNumbers");

        if (phone_numbers)
        {
            bool entry_exists = false;
            auto entry = phone_numbers->first_node();
            for (; entry != nullptr; entry = entry->next_sibling())
            {
                auto key_attr = entry->first_attribute();
                EWS_ASSERT(key_attr);
                EWS_ASSERT(
                    compare(key_attr->name(), key_attr->name_size(), "Key", 3));
                const auto key = internal::enum_to_str(number.get_key());
                if (compare(key_attr->value(), key_attr->value_size(),
                            key.c_str(), key.size()))
                {
                    entry_exists = true;
                    break;
                }
            }
            if (entry_exists)
            {
                phone_numbers->remove_node(entry);
            }
        }
        else
        {
            phone_numbers = &create_node(*doc, "t:PhoneNumbers");
        }

        // create entry & key
        const auto value = number.get_value();
        auto new_entry = &create_node(*phone_numbers, "t:Entry", value);
        auto ptr_to_key = doc->allocate_string("Key");
        const auto key = internal::enum_to_str(number.get_key());
        auto ptr_to_value = doc->allocate_string(key.c_str());
        new_entry->append_attribute(
            doc->allocate_attribute(ptr_to_key, ptr_to_value));
    }

    std::vector<phone_number> get_phone_numbers() const
    {
        const auto numbers = xml().get_node("PhoneNumbers");
        if (!numbers)
        {
            return std::vector<phone_number>();
        }
        std::vector<phone_number> result;
        for (auto entry = numbers->first_node(); entry != nullptr;
             entry = entry->next_sibling())
        {
            result.push_back(phone_number::from_xml_element(*entry));
        }
        return result;
    }

    //! Sets the name of the contact's assistant
    void set_assistant_name(const std::string& assistant_name)
    {
        xml().set_or_update("AssistantName", assistant_name);
    }

    //! Returns the contact's assistant's name
    std::string get_assistant_name() const
    {
        return xml().get_value_as_string("AssistantName");
    }

    //! The contact's birthday
    // Be careful with the formating of the date string
    // It has to be in the format YYYY-MM-DD(THH:MM:SSZ) <- can be left out
    // if the time of the day isn't important, will automatically be set to
    // YYYY-MM-DDT00:00:00Z
    //
    // This also applies to any other contact property with a date type
    // string
    void set_birthday(const std::string& birthday)
    {
        xml().set_or_update("Birthday", birthday);
    }

    std::string get_birthday() const
    {
        return xml().get_value_as_string("Birthday");
    }

    //! Sets the web page for the contact's business; typically a URL
    void set_business_homepage(const std::string& business_homepage)
    {
        xml().set_or_update("BusinessHomePage", business_homepage);
    }

    //! Returns the URL of the contact
    std::string get_business_homepage() const
    {
        return xml().get_value_as_string("BusinessHomePage");
    }

    //! A collection of children's names associated with the contact
    void set_children(const std::vector<std::string>& children)
    {
        set_array_of_strings_helper(children, "Children");
    }

    std::vector<std::string> get_children() const
    {
        return get_array_of_strings_helper("Children");
    }

    //! A collection of companies a contact is associated with
    void set_companies(const std::vector<std::string>& companies)
    {
        set_array_of_strings_helper(companies, "Companies");
    }

    std::vector<std::string> get_companies() const
    {
        return get_array_of_strings_helper("Companies");
    }

    //! \brief Indicates whether this is a directory or a store contact.
    //!
    //! This is a read-only property.
    std::string get_contact_source() const
    {
        return xml().get_value_as_string("ContactSource");
    }

    //! Set the department name that the contact is in
    void set_department(const std::string& department)
    {
        xml().set_or_update("Department", department);
    }

    //! Return the department name of the contact
    std::string get_department() const
    {
        return xml().get_value_as_string("Department");
    }

    //! Sets the generation of the contact
    //! e.g.: Sr, Jr, I, II, III, and so on
    void set_generation(const std::string& generation)
    {
        xml().set_or_update("Generation", generation);
    }

    //! Returns the generation of the contact
    std::string get_generation() const
    {
        return xml().get_value_as_string("Generation");
    }

    //! A collection of instant messaging addresses for the contact

    void set_im_address(const im_address& im_address)
    {
        using rapidxml::internal::compare;
        using internal::create_node;
        auto doc = xml().document();
        auto im_addresses = xml().get_node("ImAddresses");

        if (im_addresses)
        {
            bool entry_exists = false;
            auto entry = im_addresses->first_node();
            for (; entry != nullptr; entry = entry->next_sibling())
            {
                auto key_attr = entry->first_attribute();
                EWS_ASSERT(key_attr);
                EWS_ASSERT(
                    compare(key_attr->name(), key_attr->name_size(), "Key", 3));
                const auto key = internal::enum_to_str(im_address.get_key());
                if (compare(key_attr->value(), key_attr->value_size(),
                            key.c_str(), key.size()))
                {
                    entry_exists = true;
                    break;
                }
            }
            if (entry_exists)
            {
                im_addresses->remove_node(entry);
            }
        }
        else
        {
            im_addresses = &create_node(*doc, "t:ImAddresses");
        }

        // create entry & key
        const auto value = im_address.get_value();
        auto new_entry = &create_node(*im_addresses, "t:Entry", value);
        auto ptr_to_key = doc->allocate_string("Key");
        const auto key = internal::enum_to_str(im_address.get_key());
        auto ptr_to_value = doc->allocate_string(key.c_str());
        new_entry->append_attribute(
            doc->allocate_attribute(ptr_to_key, ptr_to_value));
    }

    std::vector<im_address> get_im_addresses() const
    {
        const auto addresses = xml().get_node("ImAddresses");
        if (!addresses)
        {
            return std::vector<im_address>();
        }
        std::vector<im_address> result;
        for (auto entry = addresses->first_node(); entry != nullptr;
             entry = entry->next_sibling())
        {
            result.push_back(im_address::from_xml_element(*entry));
        }
        return result;
    }

    //! Sets this contact's job title.
    void set_job_title(const std::string& title)
    {
        xml().set_or_update("JobTitle", title);
    }

    //! Returns the job title for the contact
    std::string get_job_title() const
    {
        return xml().get_value_as_string("JobTitle");
    }

    //! Sets the name of the contact's manager
    void set_manager(const std::string& manager)
    {
        xml().set_or_update("Manager", manager);
    }

    //! Returns the name of the contact's manager
    std::string get_manager() const
    {
        return xml().get_value_as_string("Manager");
    }

    //! Sets the distance that the contact resides
    //! from some reference point
    void set_mileage(const std::string& mileage)
    {
        xml().set_or_update("Mileage", mileage);
    }

    //! Returns the distance to the reference point
    std::string get_mileage() const
    {
        return xml().get_value_as_string("Mileage");
    }

    //! Sets the location of the contact's office
    void set_office_location(const std::string& office_location)
    {
        xml().set_or_update("OfficeLocation", office_location);
    }

    //! Returns the location of the contact's office
    std::string get_office_location() const
    {
        return xml().get_value_as_string("OfficeLocation");
    }

    // The physical addresses in the PhysicalAddresses collection that
    // represents the mailing address for the contact
    // TODO: get_postal_address_index

    //! Sets the occupation or discipline of the contact
    void set_profession(const std::string& profession)
    {
        xml().set_or_update("Profession", profession);
    }

    //! Returns the occupation of the contact
    std::string get_profession() const
    {
        return xml().get_value_as_string("Profession");
    }

    //! Set name of the contact's significant other
    void set_spouse_name(const std::string& spouse_name)
    {
        xml().set_or_update("SpouseName", spouse_name);
    }

    //! Get name of the contact's significant other
    std::string get_spouse_name() const
    {
        return xml().get_value_as_string("SpouseName");
    }

    //! Sets the family name of the contact; usually considered the last
    //! name
    void set_surname(const std::string& surname)
    {
        xml().set_or_update("Surname", surname);
    }

    //! Returns the family name of the contact; usually considered the last
    //! name
    std::string get_surname() const
    {
        return xml().get_value_as_string("Surname");
    }

    //! Date that the contact was married
    void set_wedding_anniversary(const std::string& anniversary)
    {
        xml().set_or_update("WeddingAnniversary", anniversary);
    }

    std::string get_wedding_anniversary() const
    {
        return xml().get_value_as_string("WeddingAnniversary");
    }

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

    //! Makes a contact instance from a \<Contact> XML element
    static contact from_xml_element(const rapidxml::xml_node<>& elem)
    {
        auto id_node =
            elem.first_node_ns(internal::uri<>::microsoft::types(), "ItemId");
        EWS_ASSERT(id_node && "Expected <ItemId>");
        return contact(item_id::from_xml_element(*id_node),
                       internal::xml_subtree(elem));
    }

private:
    template <typename U> friend class basic_service;
    std::string create_item_request_string() const
    {
        std::stringstream sstr;
        sstr << "<m:CreateItem>"
                "<m:Items>"
                "<t:Contact>";
        sstr << xml().to_string();
        sstr << "</t:Contact>"
                "</m:Items>"
                "</m:CreateItem>";
        return sstr.str();
    }

    // Helper function for get_email_address_{1,2,3}
    std::string get_email_address_by_key(const char* key) const
    {
        using rapidxml::internal::compare;

        // <Entry Key="" Name="" RoutingType="" MailboxType="" />
        const auto addresses = xml().get_node("EmailAddresses");
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
                if (compare(attr->name(), attr->name_size(), "Key", 3) &&
                    compare(attr->value(), attr->value_size(), key,
                            std::strlen(key)))
                {
                    return std::string(entry->value(), entry->value_size());
                }
            }
        }
        // None with such key
        return "";
    }

    // Helper function for set_email_address_{1,2,3}
    void set_email_address_by_key(const char* key, mailbox&& mail)
    {
        using rapidxml::internal::compare;
        using internal::create_node;

        auto doc = xml().document();
        auto addresses = xml().get_node("EmailAddresses");
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
                    if (compare(attr->name(), attr->name_size(), "Key", 3) &&
                        compare(attr->value(), attr->value_size(), key,
                                std::strlen(key)))
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
            addresses = &create_node(*doc, "t:EmailAddresses");
        }

        // <Entry Key="" Name="" RoutingType="" MailboxType="" />
        auto new_entry = &create_node(*addresses, "t:Entry", mail.value());
        auto ptr_to_key = doc->allocate_string("Key");
        auto ptr_to_value = doc->allocate_string(key);
        new_entry->append_attribute(
            doc->allocate_attribute(ptr_to_key, ptr_to_value));
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
            ptr_to_value = doc->allocate_string(mail.routing_type().c_str());
            new_entry->append_attribute(
                doc->allocate_attribute(ptr_to_key, ptr_to_value));
        }
        if (!mail.mailbox_type().empty())
        {
            ptr_to_key = doc->allocate_string("MailboxType");
            ptr_to_value = doc->allocate_string(mail.mailbox_type().c_str());
            new_entry->append_attribute(
                doc->allocate_attribute(ptr_to_key, ptr_to_value));
        }
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(std::is_default_constructible<contact>::value, "");
static_assert(std::is_copy_constructible<contact>::value, "");
static_assert(std::is_copy_assignable<contact>::value, "");
static_assert(std::is_move_constructible<contact>::value, "");
static_assert(std::is_move_assignable<contact>::value, "");
#endif

//! \brief Holds a subset of properties from an existing calendar item.
//!
//! \sa calendar_item::get_first_occurrence
//! \sa calendar_item::get_last_occurrence
//! \sa calendar_item::get_modified_occurrences
//! \sa calendar_item::get_deleted_occurrences
class occurrence_info final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    occurrence_info() = default;
#else
    occurrence_info() {}
#endif

    occurrence_info(item_id id, date_time start, date_time end,
                    date_time original_start)
        : item_id_(std::move(id)), start_(std::move(start)),
          end_(std::move(end)), original_start_(std::move(original_start))
    {
    }

    //! True if this occurrence_info is undefined
    bool none() const EWS_NOEXCEPT { return !item_id_.valid(); }

    const item_id& get_item_id() const EWS_NOEXCEPT { return item_id_; }

    const date_time& get_start() const EWS_NOEXCEPT { return start_; }

    const date_time& get_end() const EWS_NOEXCEPT { return end_; }

    const date_time& get_original_start() const EWS_NOEXCEPT
    {
        return original_start_;
    }

    //! \brief Makes a occurrence_info instance from a \<FirstOccurrence>,
    //! \<LastOccurrence>, or \<Occurrence> XML element.
    static occurrence_info from_xml_element(const rapidxml::xml_node<>& elem)
    {
        using rapidxml::internal::compare;

        date_time original_start;
        date_time end;
        date_time start;
        item_id id;

        for (auto node = elem.first_node(); node; node = node->next_sibling())
        {
            if (compare(node->local_name(), node->local_name_size(),
                        "OriginalStart", std::strlen("OriginalStart")))
            {
                original_start =
                    date_time(std::string(node->value(), node->value_size()));
            }
            else if (compare(node->local_name(), node->local_name_size(), "End",
                             std::strlen("End")))
            {
                end = date_time(std::string(node->value(), node->value_size()));
            }
            else if (compare(node->local_name(), node->local_name_size(),
                             "Start", std::strlen("Start")))
            {
                start =
                    date_time(std::string(node->value(), node->value_size()));
            }
            else if (compare(node->local_name(), node->local_name_size(),
                             "ItemId", std::strlen("ItemId")))
            {
                id = item_id::from_xml_element(*node);
            }
            else
            {
                throw exception("Unexpected child element in <Mailbox>");
            }
        }

        return occurrence_info(std::move(id), std::move(start), std::move(end),
                               std::move(original_start));
    }

private:
    item_id item_id_;
    date_time start_;
    date_time end_;
    date_time original_start_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(std::is_default_constructible<occurrence_info>::value, "");
static_assert(std::is_copy_constructible<occurrence_info>::value, "");
static_assert(std::is_copy_assignable<occurrence_info>::value, "");
static_assert(std::is_move_constructible<occurrence_info>::value, "");
static_assert(std::is_move_assignable<occurrence_info>::value, "");
#endif

//! Abstract base class for all recurrence patterns.
class recurrence_pattern
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    virtual ~recurrence_pattern() = default;

    recurrence_pattern(const recurrence_pattern&) = delete;
    recurrence_pattern& operator=(const recurrence_pattern&) = delete;
#else
    virtual ~recurrence_pattern() {}

private:
    recurrence_pattern(const recurrence_pattern&);
    recurrence_pattern& operator=(const recurrence_pattern&);

public:
#endif

    std::string to_xml() const { return this->to_xml_impl(); }

    //! \brief Creates a new XML element for this recurrence pattern and
    //! appends it to given parent node.
    //!
    //! Returns a reference to the newly created element.
    rapidxml::xml_node<>& to_xml_element(rapidxml::xml_node<>& parent) const
    {
        return this->to_xml_element_impl(parent);
    }

    //! \brief Makes a recurrence_pattern instance from a \<Recurrence> XML
    //! element.
    static std::unique_ptr<recurrence_pattern>
    from_xml_element(const rapidxml::xml_node<>& elem); // Defined below

protected:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    recurrence_pattern() = default;
#else
    recurrence_pattern() {}
#endif

private:
    virtual std::string to_xml_impl() const = 0;

    virtual rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>&) const = 0;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<recurrence_pattern>::value, "");
static_assert(!std::is_copy_constructible<recurrence_pattern>::value, "");
static_assert(!std::is_copy_assignable<recurrence_pattern>::value, "");
static_assert(!std::is_move_constructible<recurrence_pattern>::value, "");
static_assert(!std::is_move_assignable<recurrence_pattern>::value, "");
#endif

//! \brief An event that occurrs annually relative to a month, week, and
//! day.
//!
//! Describes an annual relative recurrence, e.g., every third Monday in
//! April.
class relative_yearly_recurrence final : public recurrence_pattern
{
public:
    relative_yearly_recurrence(day_of_week days_of_week,
                               day_of_week_index index, month m)
        : days_of_week_(days_of_week), index_(index), month_(m)
    {
    }

    day_of_week get_days_of_week() const EWS_NOEXCEPT { return days_of_week_; }

    day_of_week_index get_day_of_week_index() const EWS_NOEXCEPT
    {
        return index_;
    }

    month get_month() const EWS_NOEXCEPT { return month_; }

private:
    day_of_week days_of_week_;
    day_of_week_index index_;
    month month_;

    std::string to_xml_impl() const override
    {
        using namespace internal;

        std::stringstream sstr;
        sstr << "<t:RelativeYearlyRecurrence>"
             << "<t:DaysOfWeek>" << enum_to_str(days_of_week_)
             << "</t:DaysOfWeek>"
             << "<t:DayOfWeekIndex>" << enum_to_str(index_)
             << "</t:DayOfWeekIndex>"
             << "<t:Month>" << enum_to_str(month_) << "</t:Month>"
             << "</t:RelativeYearlyRecurrence>";
        return sstr.str();
    }

    rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>& parent) const override
    {
        EWS_ASSERT(parent.document() &&
                   "parent node needs to be part of a document");

        using namespace internal;
        auto& pattern_node = create_node(parent, "t:RelativeYearlyRecurrence");
        create_node(pattern_node, "t:DaysOfWeek", enum_to_str(days_of_week_));
        create_node(pattern_node, "t:DayOfWeekIndex", enum_to_str(index_));
        create_node(pattern_node, "t:Month", enum_to_str(month_));
        return pattern_node;
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<relative_yearly_recurrence>::value,
              "");
static_assert(!std::is_copy_constructible<relative_yearly_recurrence>::value,
              "");
static_assert(!std::is_copy_assignable<relative_yearly_recurrence>::value, "");
static_assert(!std::is_move_constructible<relative_yearly_recurrence>::value,
              "");
static_assert(!std::is_move_assignable<relative_yearly_recurrence>::value, "");
#endif

//! A yearly recurrence pattern, e.g., a birthday.
class absolute_yearly_recurrence final : public recurrence_pattern
{
public:
    absolute_yearly_recurrence(std::uint32_t day_of_month, month m)
        : day_of_month_(day_of_month), month_(m)
    {
    }

    std::uint32_t get_day_of_month() const EWS_NOEXCEPT
    {
        return day_of_month_;
    }

    month get_month() const EWS_NOEXCEPT { return month_; }

private:
    std::uint32_t day_of_month_;
    month month_;

    std::string to_xml_impl() const override
    {
        using namespace internal;

        std::stringstream sstr;
        sstr << "<t:AbsoluteYearlyRecurrence>"
             << "<t:DayOfMonth>" << day_of_month_ << "</t:DayOfMonth>"
             << "<t:Month>" << enum_to_str(month_) << "</t:Month>"
             << "</t:AbsoluteYearlyRecurrence>";
        return sstr.str();
    }

    rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>& parent) const override
    {
        using namespace internal;
        EWS_ASSERT(parent.document() &&
                   "parent node needs to be part of a document");

        auto& pattern_node = create_node(parent, "t:AbsoluteYearlyRecurrence");
        create_node(pattern_node, "t:DayOfMonth",
                    std::to_string(day_of_month_));
        create_node(pattern_node, "t:Month", internal::enum_to_str(month_));
        return pattern_node;
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<absolute_yearly_recurrence>::value,
              "");
static_assert(!std::is_copy_constructible<absolute_yearly_recurrence>::value,
              "");
static_assert(!std::is_copy_assignable<absolute_yearly_recurrence>::value, "");
static_assert(!std::is_move_constructible<absolute_yearly_recurrence>::value,
              "");
static_assert(!std::is_move_assignable<absolute_yearly_recurrence>::value, "");
#endif

//! \brief An event that occurrs on the same day each month or monthly
//! interval.
//!
//! A good example is payment of a rent that is due on the second of each
//! month.
//!
//! \code
//! auto rent_is_due = ews::absolute_monthly_recurrence(1, 2);
//! \endcode
//!
//! The \p interval parameter indicates the interval in months between
//! each time zone. For example, an \p interval value of 1 would yield an
//! appointment occurring twelve times a year, a value of 6 would produce
//! two occurrences a year and so on.
class absolute_monthly_recurrence final : public recurrence_pattern
{
public:
    absolute_monthly_recurrence(std::uint32_t interval,
                                std::uint32_t day_of_month)
        : interval_(interval), day_of_month_(day_of_month)
    {
    }

    std::uint32_t get_interval() const EWS_NOEXCEPT { return interval_; }

    std::uint32_t get_days_of_month() const EWS_NOEXCEPT
    {
        return day_of_month_;
    }

private:
    std::uint32_t interval_;
    std::uint32_t day_of_month_;

    std::string to_xml_impl() const override
    {
        using namespace internal;

        std::stringstream sstr;
        sstr << "<t:AbsoluteMonthlyRecurrence>"
             << "<t:Interval>" << interval_ << "</t:Interval>"
             << "<t:DayOfMonth>" << day_of_month_ << "</t:DayOfMonth>"
             << "</t:AbsoluteMonthlyRecurrence>";
        return sstr.str();
    }

    rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>& parent) const override
    {
        using internal::create_node;
        EWS_ASSERT(parent.document() &&
                   "parent node needs to be part of a document");

        auto& pattern_node = create_node(parent, "t:AbsoluteMonthlyRecurrence");
        create_node(pattern_node, "t:Interval", std::to_string(interval_));
        create_node(pattern_node, "t:DayOfMonth",
                    std::to_string(day_of_month_));
        return pattern_node;
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(
    !std::is_default_constructible<absolute_monthly_recurrence>::value, "");
static_assert(!std::is_copy_constructible<absolute_monthly_recurrence>::value,
              "");
static_assert(!std::is_copy_assignable<absolute_monthly_recurrence>::value, "");
static_assert(!std::is_move_constructible<absolute_monthly_recurrence>::value,
              "");
static_assert(!std::is_move_assignable<absolute_monthly_recurrence>::value, "");
#endif

//! \brief An event that occurrs annually relative to a month, week, and
//! day.
//!
//! For example, if you are a member of a C++ user group that decides to
//! meet on the third Thursday of every month you would write
//!
//! \code
//! auto meetup =
//!        ews::relative_monthly_recurrence(1,
//!                                         ews::day_of_week::thu,
//!                                         ews::day_of_week_index::third);
//! \endcode
//!
//! The \p interval parameter indicates the interval in months between
//! each time zone. For example, an \p interval value of 1 would yield an
//! appointment occurring twelve times a year, a value of 6 would produce
//! two occurrences a year and so on.
class relative_monthly_recurrence final : public recurrence_pattern
{
public:
    relative_monthly_recurrence(std::uint32_t interval,
                                day_of_week days_of_week,
                                day_of_week_index index)
        : interval_(interval), days_of_week_(days_of_week), index_(index)
    {
    }

    std::uint32_t get_interval() const EWS_NOEXCEPT { return interval_; }

    day_of_week get_days_of_week() const EWS_NOEXCEPT { return days_of_week_; }

    day_of_week_index get_day_of_week_index() const EWS_NOEXCEPT
    {
        return index_;
    }

private:
    std::uint32_t interval_;
    day_of_week days_of_week_;
    day_of_week_index index_;

    std::string to_xml_impl() const override
    {
        using namespace internal;

        std::stringstream sstr;
        sstr << "<t:RelativeMonthlyRecurrence>"
             << "<t:Interval>" << interval_ << "</t:Interval>"
             << "<t:DaysOfWeek>" << enum_to_str(days_of_week_)
             << "</t:DaysOfWeek>"
             << "<t:DayOfWeekIndex>" << enum_to_str(index_)
             << "</t:DayOfWeekIndex>"
             << "</t:RelativeMonthlyRecurrence>";
        return sstr.str();
    }

    rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>& parent) const override
    {
        using internal::create_node;
        auto& pattern_node = create_node(parent, "t:RelativeMonthlyRecurrence");
        create_node(pattern_node, "t:Interval", std::to_string(interval_));
        create_node(pattern_node, "t:DaysOfWeek",
                    internal::enum_to_str(days_of_week_));
        create_node(pattern_node, "t:DayOfWeekIndex",
                    internal::enum_to_str(index_));
        return pattern_node;
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(
    !std::is_default_constructible<relative_monthly_recurrence>::value, "");
static_assert(!std::is_copy_constructible<relative_monthly_recurrence>::value,
              "");
static_assert(!std::is_copy_assignable<relative_monthly_recurrence>::value, "");
static_assert(!std::is_move_constructible<relative_monthly_recurrence>::value,
              "");
static_assert(!std::is_move_assignable<relative_monthly_recurrence>::value, "");
#endif

//! \brief A weekly recurrence
//!
//! An example for a weekly recurrence is a regular meeting on a specific
//! day each week.
//!
//! \code
//! auto standup =
//!        ews::weekly_recurrence(1, ews::day_of_week::mon);
//! \endcode
class weekly_recurrence final : public recurrence_pattern
{
public:
    weekly_recurrence(std::uint32_t interval, day_of_week days_of_week,
                      day_of_week first_day_of_week = day_of_week::mon)
        : interval_(interval), days_of_week_(),
          first_day_of_week_(first_day_of_week)
    {
        days_of_week_.push_back(days_of_week);
    }

    weekly_recurrence(std::uint32_t interval,
                      std::vector<day_of_week> days_of_week,
                      day_of_week first_day_of_week = day_of_week::mon)
        : interval_(interval), days_of_week_(std::move(days_of_week)),
          first_day_of_week_(first_day_of_week)
    {
    }

    std::uint32_t get_interval() const EWS_NOEXCEPT { return interval_; }

    const std::vector<day_of_week>& get_days_of_week() const EWS_NOEXCEPT
    {
        return days_of_week_;
    }

    day_of_week get_first_day_of_week() const EWS_NOEXCEPT
    {
        return first_day_of_week_;
    }

private:
    std::uint32_t interval_;
    std::vector<day_of_week> days_of_week_;
    day_of_week first_day_of_week_;

    std::string to_xml_impl() const override
    {
        using namespace internal;

        std::string value;
        for (const auto& day : days_of_week_)
        {
            value += enum_to_str(day) + " ";
        }
        value.resize(value.size() - 1);
        std::stringstream sstr;
        sstr << "<t:WeeklyRecurrence>"
             << "<t:Interval>" << interval_ << "</t:Interval>"
             << "<t:DaysOfWeek>" << value << "</t:DaysOfWeek>"
             << "<t:FirstDayOfWeek>" << enum_to_str(first_day_of_week_)
             << "</t:FirstDayOfWeek>"
             << "</t:WeeklyRecurrence>";
        return sstr.str();
    }

    rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>& parent) const override
    {
        using namespace internal;
        EWS_ASSERT(parent.document() &&
                   "parent node needs to be part of a document");

        auto& pattern_node = create_node(parent, "t:WeeklyRecurrence");
        create_node(pattern_node, "t:Interval", std::to_string(interval_));
        std::string value;
        for (const auto& day : days_of_week_)
        {
            value += enum_to_str(day) + " ";
        }
        value.pop_back();
        create_node(pattern_node, "t:DaysOfWeek", value);
        create_node(pattern_node, "t:FirstDayOfWeek",
                    enum_to_str(first_day_of_week_));
        return pattern_node;
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<weekly_recurrence>::value, "");
static_assert(!std::is_copy_constructible<weekly_recurrence>::value, "");
static_assert(!std::is_copy_assignable<weekly_recurrence>::value, "");
static_assert(!std::is_move_constructible<weekly_recurrence>::value, "");
static_assert(!std::is_move_assignable<weekly_recurrence>::value, "");
#endif

//! \brief Describes a daily recurring event
class daily_recurrence final : public recurrence_pattern
{
public:
    explicit daily_recurrence(std::uint32_t interval) : interval_(interval) {}

    std::uint32_t get_interval() const EWS_NOEXCEPT { return interval_; }

private:
    std::uint32_t interval_;

    std::string to_xml_impl() const override
    {
        using namespace internal;

        std::stringstream sstr;
        sstr << "<t:DailyRecurrence>"
             << "<t:Interval>" << interval_ << "</t:Interval>"
             << "</t:DailyRecurrence>";
        return sstr.str();
    }

    rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>& parent) const override
    {
        using internal::create_node;
        EWS_ASSERT(parent.document() &&
                   "parent node needs to be part of a document");

        auto& pattern_node = create_node(parent, "t:DailyRecurrence");
        create_node(pattern_node, "t:Interval", std::to_string(interval_));
        return pattern_node;
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<daily_recurrence>::value, "");
static_assert(!std::is_copy_constructible<daily_recurrence>::value, "");
static_assert(!std::is_copy_assignable<daily_recurrence>::value, "");
static_assert(!std::is_move_constructible<daily_recurrence>::value, "");
static_assert(!std::is_move_assignable<daily_recurrence>::value, "");
#endif

//! Abstract base class for all recurrence ranges.
class recurrence_range
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    virtual ~recurrence_range() = default;

    recurrence_range(const recurrence_range&) = delete;
    recurrence_range& operator=(const recurrence_range&) = delete;
#else
    virtual ~recurrence_range() {}

private:
    recurrence_range(const recurrence_range&);
    recurrence_range& operator=(const recurrence_range&);

public:
#endif
    std::string to_xml() const { return this->to_xml_impl(); }

    //! \brief Creates a new XML element for this recurrence range and
    //! appends it to given parent node.
    //!
    //! Returns a reference to the newly created element.
    rapidxml::xml_node<>& to_xml_element(rapidxml::xml_node<>& parent) const
    {
        return this->to_xml_element_impl(parent);
    }

    //! Makes a recurrence_range instance from a \<Recurrence> XML element.
    static std::unique_ptr<recurrence_range>
    from_xml_element(const rapidxml::xml_node<>& elem); // Defined below

protected:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    recurrence_range() = default;
#else
    recurrence_range() {}
#endif

private:
    virtual std::string to_xml_impl() const = 0;

    virtual rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>&) const = 0;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<recurrence_range>::value, "");
static_assert(!std::is_copy_constructible<recurrence_range>::value, "");
static_assert(!std::is_copy_assignable<recurrence_range>::value, "");
static_assert(!std::is_move_constructible<recurrence_range>::value, "");
static_assert(!std::is_move_assignable<recurrence_range>::value, "");
#endif

//! Represents recurrence range with no end date
class no_end_recurrence_range final : public recurrence_range
{
public:
    explicit no_end_recurrence_range(date start_date)
        : start_date_(std::move(start_date))
    {
    }

    const date_time& get_start_date() const EWS_NOEXCEPT { return start_date_; }

private:
    date start_date_;

    std::string to_xml_impl() const override
    {
        using namespace internal;

        std::stringstream sstr;
        sstr << "<t:NoEndRecurrence>"
             << "<t:StartDate>" << start_date_.to_string() << "</t:StartDate>"
             << "</t:NoEndRecurrence>";
        return sstr.str();
    }

    rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>& parent) const override
    {
        using namespace internal;
        EWS_ASSERT(parent.document() &&
                   "parent node needs to be part of a document");

        auto& pattern_node = create_node(parent, "t:NoEndRecurrence");
        create_node(pattern_node, "t:StartDate", start_date_.to_string());
        return pattern_node;
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<no_end_recurrence_range>::value,
              "");
static_assert(!std::is_copy_constructible<no_end_recurrence_range>::value, "");
static_assert(!std::is_copy_assignable<no_end_recurrence_range>::value, "");
static_assert(!std::is_move_constructible<no_end_recurrence_range>::value, "");
static_assert(!std::is_move_assignable<no_end_recurrence_range>::value, "");
#endif

//! Represents recurrence range with end date
class end_date_recurrence_range final : public recurrence_range
{
public:
    end_date_recurrence_range(date start_date, date end_date)
        : start_date_(std::move(start_date)), end_date_(std::move(end_date))
    {
    }

    const date_time& get_start_date() const EWS_NOEXCEPT { return start_date_; }

    const date_time& get_end_date() const EWS_NOEXCEPT { return end_date_; }

private:
    date start_date_;
    date end_date_;

    std::string to_xml_impl() const override
    {
        using namespace internal;

        std::stringstream sstr;
        sstr << "<t:EndDateRecurrence>"
             << "<t:StartDate>" << start_date_.to_string() << "</t:StartDate>"
             << "<t:EndDate>" << end_date_.to_string() << "</t:EndDate>"
             << "</t:EndDateRecurrence>";
        return sstr.str();
    }

    rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>& parent) const override
    {
        using internal::create_node;
        EWS_ASSERT(parent.document() &&
                   "parent node needs to be part of a document");

        auto& pattern_node = create_node(parent, "t:EndDateRecurrence");
        create_node(pattern_node, "t:StartDate", start_date_.to_string());
        create_node(pattern_node, "t:EndDate", end_date_.to_string());
        return pattern_node;
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<end_date_recurrence_range>::value,
              "");
static_assert(!std::is_copy_constructible<end_date_recurrence_range>::value,
              "");
static_assert(!std::is_copy_assignable<end_date_recurrence_range>::value, "");
static_assert(!std::is_move_constructible<end_date_recurrence_range>::value,
              "");
static_assert(!std::is_move_assignable<end_date_recurrence_range>::value, "");
#endif

//! Represents a numbered recurrence range
class numbered_recurrence_range final : public recurrence_range
{
public:
    numbered_recurrence_range(date start_date, std::uint32_t no_of_occurrences)
        : start_date_(std::move(start_date)),
          no_of_occurrences_(no_of_occurrences)
    {
    }

    const date_time& get_start_date() const EWS_NOEXCEPT { return start_date_; }

    std::uint32_t get_number_of_occurrences() const EWS_NOEXCEPT
    {
        return no_of_occurrences_;
    }

private:
    date start_date_;
    std::uint32_t no_of_occurrences_;

    std::string to_xml_impl() const override
    {
        using namespace internal;

        std::stringstream sstr;
        sstr << "<t:NumberedRecurrence>"
             << "<t:StartDate>" << start_date_.to_string() << "</t:StartDate>"
             << "<t:NumberOfOccurrences>" << no_of_occurrences_
             << "</t:NumberOfOccurrences>"
             << "</t:NumberedRecurrence>";
        return sstr.str();
    }

    rapidxml::xml_node<>&
    to_xml_element_impl(rapidxml::xml_node<>& parent) const override
    {
        using namespace internal;
        EWS_ASSERT(parent.document() &&
                   "parent node needs to be part of a document");

        auto& pattern_node = create_node(parent, "t:NumberedRecurrence");
        create_node(pattern_node, "t:StartDate", start_date_.to_string());
        create_node(pattern_node, "t:NumberOfOccurrences",
                    std::to_string(no_of_occurrences_));
        return pattern_node;
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<numbered_recurrence_range>::value,
              "");
static_assert(!std::is_copy_constructible<numbered_recurrence_range>::value,
              "");
static_assert(!std::is_copy_assignable<numbered_recurrence_range>::value, "");
static_assert(!std::is_move_constructible<numbered_recurrence_range>::value,
              "");
static_assert(!std::is_move_assignable<numbered_recurrence_range>::value, "");
#endif

//! Represents a calendar item in the Exchange store
class calendar_item final : public item
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    calendar_item() = default;
#else
    calendar_item() {}
#endif

    //! Creates a <tt>\<CalendarItem\></tt> with given id
    explicit calendar_item(item_id id) : item(id) {}

#ifndef EWS_DOXYGEN_SHOULD_SKIP_THIS
    calendar_item(item_id&& id, internal::xml_subtree&& properties)
        : item(std::move(id), std::move(properties))
    {
    }
#endif

    //! Returns the starting date and time for this calendar item
    date_time get_start() const
    {
        return date_time(xml().get_value_as_string("Start"));
    }

    //! Sets the starting date and time for this calendar item
    void set_start(const date_time& datetime)
    {
        xml().set_or_update("Start", datetime.to_string());
    }

    //! Returns the ending date and time for this calendar item
    date_time get_end() const
    {
        return date_time(xml().get_value_as_string("End"));
    }

    //! Sets the ending date and time for this calendar item
    void set_end(const date_time& datetime)
    {
        xml().set_or_update("End", datetime.to_string());
    }

    //! \brief The original start time of a calendar item
    //!
    //! This is a read-only property.
    date_time get_original_start() const
    {
        return date_time(xml().get_value_as_string("OriginalStart"));
    }

    //! True if this calendar item is lasting all day
    bool is_all_day_event() const
    {
        return xml().get_value_as_string("IsAllDayEvent") == "true";
    }

    //! Makes this calendar item an all day event or not
    void set_all_day_event_enabled(bool enabled)
    {
        xml().set_or_update("IsAllDayEvent", enabled ? "true" : "false");
    }

    //! Returns the free/busy status of this calendar item
    free_busy_status get_legacy_free_busy_status() const
    {
        const auto val = xml().get_value_as_string("LegacyFreeBusyStatus");
        // Default seems to be 'Busy' if not explicitly set
        if (val.empty() || val == "Busy")
        {
            return free_busy_status::busy;
        }
        if (val == "Tentative")
        {
            return free_busy_status::tentative;
        }
        if (val == "Free")
        {
            return free_busy_status::free;
        }
        if (val == "OOF")
        {
            return free_busy_status::out_of_office;
        }
        if (val == "NoData")
        {
            return free_busy_status::no_data;
        }
        throw exception("Unexpected <LegacyFreeBusyStatus>");
    }

    //! Sets the free/busy status of this calendar item
    void set_legacy_free_busy_status(free_busy_status status)
    {
        xml().set_or_update("LegacyFreeBusyStatus",
                            internal::enum_to_str(status));
    }

    //! \brief Returns the location where a meeting or event is supposed
    //! to take place.
    std::string get_location() const
    {
        return xml().get_value_as_string("Location");
    }

    //! \brief Sets the location where a meeting or event is supposed to
    //! take place.
    void set_location(const std::string& location)
    {
        xml().set_or_update("Location", location);
    }

    //! Returns a description of when this calendar item occurs
    std::string get_when() const { return xml().get_value_as_string("When"); }

    //! Sets a description of when this calendar item occurs
    void set_when(const std::string& desc)
    {
        xml().set_or_update("When", desc);
    }

    //! \brief Indicates whether this calendar item is a meeting
    //!
    //! This is a read-only property.
    bool is_meeting() const
    {
        return xml().get_value_as_string("IsMeeting") == "true";
    }

    //! \brief Indicates whether this calendar item has been cancelled by
    //! the organizer.
    //!
    //! This is a read-only property. It is only meaningful for meetings.
    bool is_cancelled() const
    {
        return xml().get_value_as_string("IsCancelled") == "true";
    }

    //! \brief True if a calendar item is part of a recurring series.
    //!
    //! This is a read-only property. Note that a recurring master is not
    //! considered part of a recurring series, even though it holds the
    //! recurrence information.
    bool is_recurring() const
    {
        return xml().get_value_as_string("IsRecurring") == "true";
    }

    //! \brief True if a meeting request for this calendar item has been
    //! sent to all attendees.
    //!
    //! This is a read-only property.
    bool meeting_request_was_sent() const
    {
        return xml().get_value_as_string("MeetingRequestWasSent") == "true";
    }

    //! \brief Indicates whether a response to a calendar item is needed.
    //!
    //! This is a read-only property.
    bool is_response_requested() const
    {
        return xml().get_value_as_string("IsResponseRequested") == "true";
    }

    //! \brief Returns the type of this calendar item
    //!
    //! This is a read-only property.
    calendar_item_type get_calendar_item_type() const
    {
        const auto val = xml().get_value_as_string("CalendarItemType");
        // By default, newly created calendar items are of type 'Single'
        if (val.empty() || val == "Single")
        {
            return calendar_item_type::single;
        }
        if (val == "Occurrence")
        {
            return calendar_item_type::occurrence;
        }
        if (val == "Exception")
        {
            return calendar_item_type::exception;
        }
        if (val == "RecurringMaster")
        {
            return calendar_item_type::recurring_master;
        }
        throw exception("Unexpected <CalendarItemType>");
    }

    //! \brief Returns the response of this calendar item's owner to the
    //! meeting
    //!
    //! This is a read-only property.
    response_type get_my_response_type() const
    {
        const auto val = xml().get_value_as_string("MyResponseType");
        if (val.empty())
        {
            return response_type::unknown;
        }
        return internal::str_to_response_type(val);
    }

    //! \brief Returns the organizer of this calendar item
    //!
    //! For meetings, the party responsible for coordinating attendance.
    //! This is a read-only property.
    mailbox get_organizer() const
    {
        const auto organizer = xml().get_node("Organizer");
        if (!organizer)
        {
            return mailbox(); // None
        }
        return mailbox::from_xml_element(*(organizer->first_node()));
    }

    //! Returns all attendees required to attend this meeting
    std::vector<attendee> get_required_attendees() const
    {
        return get_attendees_helper("RequiredAttendees");
    }

    //! Sets the attendees required to attend this meeting
    void set_required_attendees(const std::vector<attendee>& attendees) const
    {
        set_attendees_helper("RequiredAttendees", attendees);
    }

    //! Returns all attendees not required to attend this meeting
    std::vector<attendee> get_optional_attendees() const
    {
        return get_attendees_helper("OptionalAttendees");
    }

    //! Sets the attendees not required to attend this meeting
    void set_optional_attendees(const std::vector<attendee>& attendees) const
    {
        set_attendees_helper("OptionalAttendees", attendees);
    }

    //! Returns all scheduled resources of this meeting
    std::vector<attendee> get_resources() const
    {
        return get_attendees_helper("Resources");
    }

    //! Sets the scheduled resources of this meeting
    void set_resources(const std::vector<attendee>& resources) const
    {
        set_attendees_helper("Resources", resources);
    }

    //! \brief Returns the number of meetings that are in conflict with
    //! this meeting
    //!
    //! Note: this property is only included in \<GetItem/> response when
    //! 'calendar:ConflictingMeetingCount' is passed in
    //! \<AdditionalProperties/>. This is a read-only property.
    //! \sa service::get_conflicting_meetings
    int get_conflicting_meeting_count() const
    {
        const auto val = xml().get_value_as_string("ConflictingMeetingCount");
        return val.empty() ? 0 : std::stoi(val);
    }

    //! \brief Returns the number of meetings that are adjacent to this
    //! meeting
    //!
    //! Note: this property is only included in \<GetItem/> response when
    //! 'calendar:AdjacentMeetingCount' is passed in
    //! \<AdditionalProperties/>. This is a read-only property.
    //! \sa service::get_adjacent_meetings
    int get_adjacent_meeting_count() const
    {
        const auto val = xml().get_value_as_string("AdjacentMeetingCount");
        return val.empty() ? 0 : std::stoi(val);
    }

    // TODO: issue #19
    // <ConflictingMeetings/>
    // <AdjacentMeetings/>

    //! \brief Returns the duration of this meeting.
    //!
    //! This is a read-only property.
    duration get_duration() const
    {
        return duration(xml().get_value_as_string("Duration"));
    }

    //! \brief Returns the display name for the time zone associated with
    //! this calendar item.
    //!
    //! Provides a text-only description of the time zone for a calendar
    //! item.
    //!
    //! Note: This is a read-only property. EWS does not allow you to
    //! specify the name of a time zone for a calendar item with
    //! \<CreateItem/> and \<UpdateItem/>. Hence, any calendar items that
    //! is fetched with EWS that were also created with EWS will always
    //! have an empty \<TimeZone/> property.
    //!
    //! \sa get_meeting_timezone, set_meeting_timezone
    std::string get_timezone() const
    {
        return xml().get_value_as_string("TimeZone");
    }

    //! \brief Returns the date and time when when this meeting was
    //! responded to.
    //!
    //! Note: Applicable to meetings only. This is a read-only property
    date_time get_appointment_reply_time() const
    {
        return date_time(xml().get_value_as_string("AppointmentReplyTime"));
    }

    //! \brief Returns the sequence number of this meetings version
    //!
    //! Note: Applicable to meetings only. This is a read-only property
    int get_appointment_sequence_number() const
    {
        const auto val = xml().get_value_as_string("AppointmentSequenceNumber");
        return val.empty() ? 0 : std::stoi(val);
    }

    //! \brief Returns the status of this meeting.
    //!
    //! The returned integer is a bitmask.
    //!
    //! \li <tt>0x0000</tt> - No flags have been set. This is only used
    //!                       for a calendar item that does not include
    //!                       attendees
    //! \li <tt>0x0001</tt> - Appointment is a meeting
    //! \li <tt>0x0002</tt> - Appointment has been received
    //! \li <tt>0x0004</tt> - Appointment has been canceled
    //!
    //! Note: Applicable to meetings only and is only included in a
    //! meeting response. This is a read-only property
    int get_appointment_state() const
    {
        const auto val = xml().get_value_as_string("AppointmentState");
        return val.empty() ? 0 : std::stoi(val);
    }

    //! \brief Returns the recurrence pattern for calendar items and
    //! meeting requests.
    //!
    //! The returned pointers are NULL if this calendar item is not a
    //! recurring master.
    std::pair<std::unique_ptr<recurrence_pattern>,
              std::unique_ptr<recurrence_range>>
    get_recurrence() const
    {
        typedef std::pair<std::unique_ptr<recurrence_pattern>,
                          std::unique_ptr<recurrence_range>>
            return_type;
        auto node = xml().get_node("Recurrence");
        if (!node)
        {
            return return_type();
        }
        return std::make_pair(recurrence_pattern::from_xml_element(*node),
                              recurrence_range::from_xml_element(*node));
    }

    //! Sets the recurrence pattern for calendar items and meeting requests
    void set_recurrence(const recurrence_pattern& pattern,
                        const recurrence_range& range)
    {
        auto doc = xml().document();
        auto recurrence_node = xml().get_node("Recurrence");
        if (recurrence_node)
        {
            // Remove existing node first
            doc->remove_node(recurrence_node);
        }

        recurrence_node = &internal::create_node(*doc, "t:Recurrence");

        pattern.to_xml_element(*recurrence_node);
        range.to_xml_element(*recurrence_node);
    }

    //! Returns the first occurrence
    occurrence_info get_first_occurrence() const
    {
        auto node = xml().get_node("FirstOccurrence");
        if (!node)
        {
            return occurrence_info();
        }
        return occurrence_info::from_xml_element(*node);
    }

    //! Returns the last occurrence
    occurrence_info get_last_occurrence() const
    {
        auto node = xml().get_node("LastOccurrence");
        if (!node)
        {
            return occurrence_info();
        }
        return occurrence_info::from_xml_element(*node);
    }

    //! Returns the modified occurrences
    std::vector<occurrence_info> get_modified_occurrences() const
    {
        auto node = xml().get_node("ModifiedOccurrences");
        if (!node)
        {
            return std::vector<occurrence_info>();
        }

        auto occurrences = std::vector<occurrence_info>();
        for (auto occurrence = node->first_node(); occurrence;
             occurrence = occurrence->next_sibling())
        {
            occurrences.emplace_back(
                occurrence_info::from_xml_element(*occurrence));
        }
        return occurrences;
    }

    //! Returns the deleted occurrences
    std::vector<occurrence_info> get_deleted_occurrences() const
    {
        auto node = xml().get_node("DeletedOccurrences");
        if (!node)
        {
            return std::vector<occurrence_info>();
        }

        auto occurrences = std::vector<occurrence_info>();
        for (auto occurrence = node->first_node(); occurrence;
             occurrence = occurrence->next_sibling())
        {
            occurrences.emplace_back(
                occurrence_info::from_xml_element(*occurrence));
        }
        return occurrences;
    }

    // TODO: issue #23
    // <MeetingTimeZone/>
    // <StartTimeZone/>
    // <EndTimeZone/>

    //! \brief Returns the type of conferencing that is performed with this
    //! calendar item.
    //!
    //! Possible values:
    //!
    //! \li 0 - NetMeeting
    //! \li 1 - NetShow
    //! \li 2 - Chat
    int get_conference_type() const
    {
        const auto val = xml().get_value_as_string("ConferenceType");
        return val.empty() ? 0 : std::stoi(val);
    }

    //! \brief Sets the type of conferencing that is performed with this
    //! calendar item.
    //!
    //! \sa get_conference_type
    void set_conference_type(int value)
    {
        xml().set_or_update("ConferenceType", std::to_string(value));
    }

    //! \brief Returns true if attendees are allowed to respond to the
    //! organizer with new time suggestions.
    bool is_new_time_proposal_allowed() const
    {
        return xml().get_value_as_string("AllowNewTimeProposal") == "true";
    }

    //! \brief If set to true, allows attendees to respond to the organizer
    //! with new time suggestions.
    //!
    //! Note: This property is read-writable for the organizer's calendar
    //! item. For meeting requests and for attendees' calendar items this
    //! is read-only .
    void set_new_time_proposal_allowed(bool allowed)
    {
        xml().set_or_update("AllowNewTimeProposal", allowed ? "true" : "false");
    }

    //! \brief Returns whether this meeting is held online.
    bool is_online_meeting() const
    {
        return xml().get_value_as_string("IsOnlineMeeting") == "true";
    }

    //! \brief If set to true, this meeting is supposed to be held online.
    //!
    //! Note: This property is read-writable for the organizer's calendar
    //! item. For meeting requests and for attendees' calendar items this
    //! is read-only .
    void set_online_meeting_enabled(bool enabled)
    {
        xml().set_or_update("IsOnlineMeeting", enabled ? "true" : "false");
    }

    //! Returns the URL for a meeting workspace
    std::string get_meeting_workspace_url() const
    {
        return xml().get_value_as_string("MeetingWorkspaceUrl");
    }

    //! \brief Sets the URL for a meeting workspace.
    //!
    //! Note: This property is read-writable for the organizer's calendar
    //! item. For meeting requests and for attendees' calendar items this
    //! is read-only .
    void set_meeting_workspace_url(const std::string& url)
    {
        xml().set_or_update("MeetingWorkspaceUrl", url);
    }

    //! Returns a URL for Microsoft NetShow online meeting
    std::string get_net_show_url() const
    {
        return xml().get_value_as_string("NetShowUrl");
    }

    //! \brief Sets the URL for Microsoft NetShow online meeting.
    //!
    //! Note: This property is read-writable for the organizer's calendar
    //! item. For meeting requests and for attendees' calendar items this
    //! is read-only .
    void set_net_show_url(const std::string& url)
    {
        xml().set_or_update("NetShowUrl", url);
    }

    //! Makes a calendar item instance from a \<CalendarItem> XML element
    static calendar_item from_xml_element(const rapidxml::xml_node<>& elem)
    {
        auto id_node =
            elem.first_node_ns(internal::uri<>::microsoft::types(), "ItemId");
        EWS_ASSERT(id_node && "Expected <ItemId>");
        return calendar_item(item_id::from_xml_element(*id_node),
                             internal::xml_subtree(elem));
    }

private:
    inline std::vector<attendee>
    get_attendees_helper(const char* node_name) const
    {
        const auto attendees = xml().get_node(node_name);
        if (!attendees)
        {
            return std::vector<attendee>();
        }

        std::vector<attendee> result;
        for (auto attendee_node = attendees->first_node(); attendee_node;
             attendee_node = attendee_node->next_sibling())
        {
            result.emplace_back(attendee::from_xml_element(*attendee_node));
        }
        return result;
    }

    inline void
    set_attendees_helper(const char* node_name,
                         const std::vector<attendee>& attendees) const
    {
        auto doc = xml().document();

        auto attendees_node = xml().get_node(node_name);
        if (attendees_node)
        {
            doc->remove_node(attendees_node);
        }

        attendees_node =
            &internal::create_node(*doc, "t:" + std::string(node_name));

        for (const auto& a : attendees)
        {
            a.to_xml_element(*attendees_node);
        }
    }

    template <typename U> friend class basic_service;

    std::string
    create_item_request_string(send_meeting_invitations meeting_invitations =
                                   send_meeting_invitations::send_to_none) const
    {
        std::stringstream sstr;
        sstr << "<m:CreateItem SendMeetingInvitations=\"" +
                    internal::enum_to_str(meeting_invitations) +
                    "\">"
                    "<m:Items>"
                    "<t:CalendarItem>";
        sstr << xml().to_string();
        sstr << "</t:CalendarItem>"
                "</m:Items>"
                "</m:CreateItem>";
        return sstr.str();
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(std::is_default_constructible<calendar_item>::value, "");
static_assert(std::is_copy_constructible<calendar_item>::value, "");
static_assert(std::is_copy_assignable<calendar_item>::value, "");
static_assert(std::is_move_constructible<calendar_item>::value, "");
static_assert(std::is_move_assignable<calendar_item>::value, "");
#endif

//! A message item in the Exchange store
class message final : public item
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    message() = default;
#else
    message() {}
#endif

    explicit message(item_id id) : item(id) {}

#ifndef EWS_DOXYGEN_SHOULD_SKIP_THIS
    message(item_id&& id, internal::xml_subtree&& properties)
        : item(std::move(id), std::move(properties))
    {
    }
#endif

    // <Sender/>

    void set_to_recipients(const std::vector<mailbox>& recipients)
    {
        auto doc = xml().document();

        auto to_recipients_node = xml().get_node("ToRecipients");
        if (to_recipients_node)
        {
            doc->remove_node(to_recipients_node);
        }

        to_recipients_node = &internal::create_node(*doc, "t:ToRecipients");

        for (const auto& recipient : recipients)
        {
            recipient.to_xml_element(*to_recipients_node);
        }
    }

    std::vector<mailbox> get_to_recipients() const
    {
        const auto recipients = xml().get_node("ToRecipients");
        if (!recipients)
        {
            return std::vector<mailbox>();
        }
        std::vector<mailbox> result;
        for (auto mailbox_node = recipients->first_node(); mailbox_node;
             mailbox_node = mailbox_node->next_sibling())
        {
            result.emplace_back(mailbox::from_xml_element(*mailbox_node));
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

    //! Returns whether this message has been read
    bool is_read() const
    {
        return xml().get_value_as_string("IsRead") == "true";
    }

    //! \brief Sets whether this message has been read.
    //!
    //! If is_read_receipt_requested() evaluates to true, updating this
    //! property to true sends a read receipt.
    void set_is_read(bool value)
    {
        xml().set_or_update("IsRead", value ? "true" : "false");
    }

    // <IsResponseRequested/>
    // <References/>
    // <ReplyTo/>

    //! Makes a message instance from a \<Message> XML element
    static message from_xml_element(const rapidxml::xml_node<>& elem)
    {
        auto id_node =
            elem.first_node_ns(internal::uri<>::microsoft::types(), "ItemId");
        EWS_ASSERT(id_node && "Expected <ItemId>");
        return message(item_id::from_xml_element(*id_node),
                       internal::xml_subtree(elem));
    }

private:
    template <typename U> friend class basic_service;

    std::string
    create_item_request_string(ews::message_disposition disposition) const
    {
        std::stringstream sstr;
        sstr << "<m:CreateItem MessageDisposition=\""
             << internal::enum_to_str(disposition) << "\">"
                                                      "<m:Items>"
                                                      "<t:Message>";
        sstr << xml().to_string();
        sstr << "</t:Message>"
                "</m:Items>"
                "</m:CreateItem>";
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

//! Identifies frequently referenced properties by an URI
class property_path
{
public:
    // Intentionally not explicit
    property_path(const char* uri) : uri_(uri) { class_name(); }

#ifdef EWS_HAS_DEFAULT_AND_DELETE
    virtual ~property_path() = default;
    property_path(const property_path&) = default;
    property_path& operator=(const property_path&) = default;
#else
    virtual ~property_path() {}
#endif

    std::string to_xml() const { return this->to_xml_impl(); }

    std::string to_xml(const std::string& value) const
    {
        return this->to_xml_impl(value);
    }

    //! Returns the value of the \<FieldURI> element
    const std::string& field_uri() const EWS_NOEXCEPT { return uri_; }

protected:
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
            return "CalendarItem";
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

    virtual std::string to_xml_impl() const
    {
        std::stringstream sstr;
        sstr << "<t:FieldURI FieldURI=\"";
        sstr << uri_ << "\"/>";
        return sstr.str();
    }

    virtual std::string to_xml_impl(const std::string& value) const
    {
        std::stringstream sstr;
        sstr << "<t:FieldURI FieldURI=\"";
        sstr << uri_ << "\"/>";
        sstr << "<t:" << class_name() << ">";
        sstr << "<t:" << property_name() << ">";
        sstr << value;
        sstr << "</t:" << property_name() << ">";
        sstr << "</t:" << class_name() << ">";
        return sstr.str();
    }

private:
    std::string property_name() const
    {
        const auto n = uri_.rfind(':');
        EWS_ASSERT(n != std::string::npos);
        return uri_.substr(n + 1);
    }

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

//! Identifies individual members of a dictionary property by an URI and index
class indexed_property_path final : public property_path
{
public:
    indexed_property_path(const char* uri, const char* index)
        : property_path(uri), index_(index)
    {
    }

private:
    std::string to_xml_impl() const override
    {
        std::stringstream sstr;
        sstr << "<t:IndexedFieldURI FieldURI=";
        sstr << "\"" << field_uri() << "\"";
        sstr << " FieldIndex=";
        sstr << "\"";
        sstr << index_;
        sstr << "\"";
        sstr << "/>";
        return sstr.str();
    }

    std::string to_xml_impl(const std::string& value) const override
    {
        std::stringstream sstr;
        sstr << "<t:IndexedFieldURI FieldURI=";
        sstr << "\"" << field_uri() << "\"";
        sstr << " FieldIndex=";
        sstr << "\"";
        sstr << index_;
        sstr << "\"";
        sstr << "/>";
        sstr << "<t:" << class_name() << ">";
        sstr << value;
        sstr << " </t:" << class_name() << ">";
        return sstr.str();
    }

    std::string index_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<indexed_property_path>::value, "");
static_assert(std::is_copy_constructible<indexed_property_path>::value, "");
static_assert(std::is_copy_assignable<indexed_property_path>::value, "");
static_assert(std::is_move_constructible<indexed_property_path>::value, "");
static_assert(std::is_move_assignable<indexed_property_path>::value, "");
#endif

// TODO: extended_property_path missing?

namespace folder_property_path
{
    static const property_path folder_id = "folder:FolderId";
    static const property_path parent_folder_id = "folder:ParentFolderId";
    static const property_path display_name = "folder:DisplayName";
    static const property_path unread_count = "folder:UnreadCount";
    static const property_path total_count = "folder:TotalCount";
    static const property_path child_folder_count = "folder:ChildFolderCount";
    static const property_path folder_class = "folder:FolderClass";
    static const property_path search_parameters = "folder:SearchParameters";
    static const property_path managed_folder_information =
        "folder:ManagedFolderInformation";
    static const property_path permission_set = "folder:PermissionSet";
    static const property_path effective_rights = "folder:EffectiveRights";
    static const property_path sharing_effective_rights =
        "folder:SharingEffectiveRights";
}

namespace item_property_path
{
    static const property_path item_id = "item:ItemId";
    static const property_path parent_folder_id = "item:ParentFolderId";
    static const property_path item_class = "item:ItemClass";
    static const property_path mime_content = "item:MimeContent";
    static const property_path attachment = "item:Attachments";
    static const property_path subject = "item:Subject";
    static const property_path date_time_received = "item:DateTimeReceived";
    static const property_path size = "item:Size";
    static const property_path categories = "item:Categories";
    static const property_path has_attachments = "item:HasAttachments";
    static const property_path importance = "item:Importance";
    static const property_path in_reply_to = "item:InReplyTo";
    static const property_path internet_message_headers =
        "item:InternetMessageHeaders";
    static const property_path is_associated = "item:IsAssociated";
    static const property_path is_draft = "item:IsDraft";
    static const property_path is_from_me = "item:IsFromMe";
    static const property_path is_resend = "item:IsResend";
    static const property_path is_submitted = "item:IsSubmitted";
    static const property_path is_unmodified = "item:IsUnmodified";
    static const property_path date_time_sent = "item:DateTimeSent";
    static const property_path date_time_created = "item:DateTimeCreated";
    static const property_path body = "item:Body";
    static const property_path response_objects = "item:ResponseObjects";
    static const property_path sensitivity = "item:Sensitivity";
    static const property_path reminder_due_by = "item:ReminderDueBy";
    static const property_path reminder_is_set = "item:ReminderIsSet";
    static const property_path reminder_next_time = "item:ReminderNextTime";
    static const property_path reminder_minutes_before_start =
        "item:ReminderMinutesBeforeStart";
    static const property_path display_to = "item:DisplayTo";
    static const property_path display_cc = "item:DisplayCc";
    static const property_path culture = "item:Culture";
    static const property_path effective_rights = "item:EffectiveRights";
    static const property_path last_modified_name = "item:LastModifiedName";
    static const property_path last_modified_time = "item:LastModifiedTime";
    static const property_path conversation_id = "item:ConversationId";
    static const property_path unique_body = "item:UniqueBody";
    static const property_path flag = "item:Flag";
    static const property_path store_entry_id = "item:StoreEntryId";
    static const property_path instance_key = "item:InstanceKey";
    static const property_path normalized_body = "item:NormalizedBody";
    static const property_path entity_extraction_result =
        "item:EntityExtractionResult";
    static const property_path policy_tag = "item:PolicyTag";
    static const property_path archive_tag = "item:ArchiveTag";
    static const property_path retention_date = "item:RetentionDate";
    static const property_path preview = "item:Preview";
    static const property_path next_predicted_action =
        "item:NextPredictedAction";
    static const property_path grouping_action = "item:GroupingAction";
    static const property_path predicted_action_reasons =
        "item:PredictedActionReasons";
    static const property_path is_clutter = "item:IsClutter";
    static const property_path rights_management_license_data =
        "item:RightsManagementLicenseData";
    static const property_path block_status = "item:BlockStatus";
    static const property_path has_blocked_images = "item:HasBlockedImages";
    static const property_path web_client_read_from_query_string =
        "item:WebClientReadFormQueryString";
    static const property_path web_client_edit_from_query_string =
        "item:WebClientEditFormQueryString";
    static const property_path text_body = "item:TextBody";
    static const property_path icon_index = "item:IconIndex";
    static const property_path mime_content_utf8 = "item:MimeContentUTF8";
}

namespace message_property_path
{
    static const property_path conversation_index = "message:ConversationIndex";
    static const property_path conversation_topic = "message:ConversationTopic";
    static const property_path internet_message_id =
        "message:InternetMessageId";
    static const property_path is_read = "message:IsRead";
    static const property_path is_response_requested =
        "message:IsResponseRequested";
    static const property_path is_read_receipt_requested =
        "message:IsReadReceiptRequested";
    static const property_path is_delivery_receipt_requested =
        "message:IsDeliveryReceiptRequested";
    static const property_path received_by = "message:ReceivedBy";
    static const property_path received_representing =
        "message:ReceivedRepresenting";
    static const property_path references = "message:References";
    static const property_path reply_to = "message:ReplyTo";
    static const property_path from = "message:From";
    static const property_path sender = "message:Sender";
    static const property_path to_recipients = "message:ToRecipients";
    static const property_path cc_recipients = "message:CcRecipients";
    static const property_path bcc_recipients = "message:BccRecipients";
    static const property_path approval_request_data =
        "message:ApprovalRequestData";
    static const property_path voting_information = "message:VotingInformation";
    static const property_path reminder_message_data =
        "message:ReminderMessageData";
}

namespace meeting_property_path
{
    static const property_path associated_calendar_item_id =
        "meeting:AssociatedCalendarItemId";
    static const property_path is_delegated = "meeting:IsDelegated";
    static const property_path is_out_of_date = "meeting:IsOutOfDate";
    static const property_path has_been_processed = "meeting:HasBeenProcessed";
    static const property_path response_type = "meeting:ResponseType";
    static const property_path proposed_start = "meeting:ProposedStart";
    static const property_path proposed_end = "meeting:PropsedEnd";
}

namespace meeting_request_property_path
{
    static const property_path meeting_request_type =
        "meetingRequest:MeetingRequestType";
    static const property_path intended_free_busy_status =
        "meetingRequest:IntendedFreeBusyStatus";
    static const property_path change_highlights =
        "meetingRequest:ChangeHighlights";
}

namespace calendar_property_path
{
    static const property_path start = "calendar:Start";
    static const property_path end = "calendar:End";
    static const property_path original_start = "calendar:OriginalStart";
    static const property_path start_wall_clock = "calendar:StartWallClock";
    static const property_path end_wall_clock = "calendar:EndWallClock";
    static const property_path start_time_zone_id = "calendar:StartTimeZoneId";
    static const property_path end_time_zone_id = "calendar:EndTimeZoneId";
    static const property_path is_all_day_event = "calendar:IsAllDayEvent";
    static const property_path legacy_free_busy_status =
        "calendar:LegacyFreeBusyStatus";
    static const property_path location = "calendar:Location";
    static const property_path when = "calendar:When";
    static const property_path is_meeting = "calendar:IsMeeting";
    static const property_path is_cancelled = "calendar:IsCancelled";
    static const property_path is_recurring = "calendar:IsRecurring";
    static const property_path meeting_request_was_sent =
        "calendar:MeetingRequestWasSent";
    static const property_path is_response_requested =
        "calendar:IsResponseRequested";
    static const property_path calendar_item_type = "calendar:CalendarItemType";
    static const property_path my_response_type = "calendar:MyResponseType";
    static const property_path organizer = "calendar:Organizer";
    static const property_path required_attendees =
        "calendar:RequiredAttendees";
    static const property_path optional_attendees =
        "calendar:OptionalAttendees";
    static const property_path resources = "calendar:Resources";
    static const property_path conflicting_meeting_count =
        "calendar:ConflictingMeetingCount";
    static const property_path adjacent_meeting_count =
        "calendar:AdjacentMeetingCount";
    static const property_path conflicting_meetings =
        "calendar:ConflictingMeetings";
    static const property_path adjacent_meetings = "calendar:AdjacentMeetings";
    static const property_path duration = "calendar:Duration";
    static const property_path time_zone = "calendar:TimeZone";
    static const property_path appointment_reply_time =
        "calendar:AppointmentReplyTime";
    static const property_path appointment_sequence_number =
        "calendar:AppointmentSequenceNumber";
    static const property_path appointment_state = "calendar:AppointmentState";
    static const property_path recurrence = "calendar:Recurrence";
    static const property_path first_occurrence = "calendar:FirstOccurrence";
    static const property_path last_occurrence = "calendar:LastOccurrence";
    static const property_path modified_occurrences =
        "calendar:ModifiedOccurrences";
    static const property_path deleted_occurrences =
        "calendar:DeletedOccurrences";
    static const property_path meeting_time_zone = "calendar:MeetingTimeZone";
    static const property_path conference_type = "calendar:ConferenceType";
    static const property_path allow_new_time_proposal =
        "calendar:AllowNewTimeProposal";
    static const property_path is_online_meeting = "calendar:IsOnlineMeeting";
    static const property_path meeting_workspace_url =
        "calendar:MeetingWorkspaceUrl";
    static const property_path net_show_url = "calendar:NetShowUrl";
    static const property_path uid = "calendar:UID";
    static const property_path recurrence_id = "calendar:RecurrenceId";
    static const property_path date_time_stamp = "calendar:DateTimeStamp";
    static const property_path start_time_zone = "calendar:StartTimeZone";
    static const property_path end_time_zone = "calendar:EndTimeZone";
    static const property_path join_online_meeting_url =
        "calendar:JoinOnlineMeetingUrl";
    static const property_path online_meeting_settings =
        "calendar:OnlineMeetingSettings";
    static const property_path is_organizer = "calendar:IsOrganizer";
}

namespace task_property_path
{
    static const property_path actual_work = "task:ActualWork";
    static const property_path assigned_time = "task:AssignedTime";
    static const property_path billing_information = "task:BillingInformation";
    static const property_path change_count = "task:ChangeCount";
    static const property_path companies = "task:Companies";
    static const property_path complete_date = "task:CompleteDate";
    static const property_path contacts = "task:Contacts";
    static const property_path delegation_state = "task:DelegationState";
    static const property_path delegator = "task:Delegator";
    static const property_path due_date = "task:DueDate";
    static const property_path is_assignment_editable =
        "task:IsAssignmentEditable";
    static const property_path is_complete = "task:IsComplete";
    static const property_path is_recurring = "task:IsRecurring";
    static const property_path is_team_task = "task:IsTeamTask";
    static const property_path mileage = "task:Mileage";
    static const property_path owner = "task:Owner";
    static const property_path percent_complete = "task:PercentComplete";
    static const property_path recurrence = "task:Recurrence";
    static const property_path start_date = "task:StartDate";
    static const property_path status = "task:Status";
    static const property_path status_description = "task:StatusDescription";
    static const property_path total_work = "task:TotalWork";
}

namespace contact_property_path
{
    static const property_path alias = "contacts:Alias";
    static const property_path assistant_name = "contacts:AssistantName";
    static const property_path birthday = "contacts:Birthday";
    static const property_path business_home_page = "contacts:BusinessHomePage";
    static const property_path children = "contacts:Children";
    static const property_path companies = "contacts:Companies";
    static const property_path company_name = "contacts:CompanyName";
    static const property_path complete_name = "contacts:CompleteName";
    static const property_path contact_source = "contacts:ContactSource";
    static const property_path culture = "contacts:Culture";
    static const property_path department = "contacts:Department";
    static const property_path display_name = "contacts:DisplayName";
    static const property_path directory_id = "contacts:DirectoryId";
    static const property_path direct_reports = "contacts:DirectReports";
    static const property_path email_addresses = "contacts:EmailAddresses";
    static const property_path email_address = "contacts:EmailAddress";
    static const indexed_property_path email_address_1("contacts:EmailAddress",
                                                       "EmailAddress1");
    static const indexed_property_path email_address_2("contacts:EmailAddress",
                                                       "EmailAddress2");
    static const indexed_property_path email_address_3("contacts:EmailAddress",
                                                       "EmailAddress3");
    static const property_path file_as = "contacts:FileAs";
    static const property_path file_as_mapping = "contacts:FileAsMapping";
    static const property_path generation = "contacts:Generation";
    static const property_path given_name = "contacts:GivenName";
    static const property_path im_addresses = "contacts:ImAddresses";
    static const property_path im_address = "contacts:ImAddress";
    static const indexed_property_path im_address_1("contacts:ImAddress",
                                                    "ImAddress1");
    static const indexed_property_path im_address_2("contacts:ImAddress",
                                                    "ImAddress2");
    static const indexed_property_path im_address_3("contacts:ImAddress",
                                                    "ImAddress3");
    static const property_path initials = "contacts:Initials";
    static const property_path job_title = "contacts:JobTitle";
    static const property_path manager = "contacts:Manager";
    static const property_path manager_mailbox = "contacts:ManagerMailbox";
    static const property_path middle_name = "contacts:MiddleName";
    static const property_path mileage = "contacts:Mileage";
    static const property_path ms_exchange_certificate =
        "contacts:MSExchangeCertificate";
    static const property_path nickname = "contacts:Nickname";
    static const property_path notes = "contacts:Notes";
    static const property_path office_location = "contacts:OfficeLocation";
    static const property_path phone_numbers = "contacts:PhoneNumbers";

    namespace phone_number
    {
        static const indexed_property_path home_phone("contacts:PhoneNumber",
                                                      "HomePhone");
        static const indexed_property_path pager("contacts:PhoneNumber",
                                                 "Pager");
        static const indexed_property_path
            business_phone("contacts:PhoneNumber", "BusinessPhone");
    }

    static const property_path phonetic_full_name = "contacts:PhoneticFullName";
    static const property_path phonetic_first_name =
        "contacts:PhoneticFirstName";
    static const property_path phonetic_last_name = "contacts:PhoneticLastName";
    static const property_path photo = "contacts:Photo";
    static const property_path physical_addresses =
        "contacts:PhysicalAddresses";

    namespace physical_address
    {
        static const indexed_property_path street("contacts:PhysicalAddress",
                                                  "Street");
        static const indexed_property_path city("contacts:PhysicalAddress:City",
                                                "Home");
        static const indexed_property_path state("contacts:PhysicalAddress",
                                                 "State");
        static const indexed_property_path
            country_or_region("contacts:PhysicalAddress", "CountryOrRegion");
        static const indexed_property_path
            postal_code("contacts:PhysicalAddress", "PostalCode");
    }

    static const property_path postal_adress_index =
        "contacts:PostalAddressIndex";
    static const property_path profession = "contacts:Profession";
    static const property_path spouse_name = "contacts:SpouseName";
    static const property_path surname = "contacts:Surname";
    static const property_path wedding_anniversary =
        "contacts:WeddingAnniversary";
    static const property_path smime_certificate =
        "contacts:UserSMIMECertificate";
    static const property_path has_picture = "contacts:HasPicture";
}

namespace distribution_list_property_path
{
    static const property_path members = "distributionlist:Members";
}

namespace post_item_property_path
{
    static const property_path posted_time = "postitem:PostedTime";
}

namespace conversation_property_path
{
    static const property_path conversation_id = "conversation:ConversationId";
    static const property_path conversation_topic =
        "conversation:ConversationTopic";
    static const property_path unique_recipients =
        "conversation:UniqueRecipients";
    static const property_path global_unique_recipients =
        "conversation:GlobalUniqueRecipients";
    static const property_path unique_unread_senders =
        "conversation:UniqueUnreadSenders";
    static const property_path global_unique_unread_readers =
        "conversation:GlobalUniqueUnreadSenders";
    static const property_path unique_senders = "conversation:UniqueSenders";
    static const property_path global_unique_senders =
        "conversation:GlobalUniqueSenders";
    static const property_path last_delivery_time =
        "conversation:LastDeliveryTime";
    static const property_path global_last_delivery_time =
        "conversation:GlobalLastDeliveryTime";
    static const property_path categories = "conversation:Categories";
    static const property_path global_categories =
        "conversation:GlobalCategories";
    static const property_path flag_status = "conversation:FlagStatus";
    static const property_path global_flag_status =
        "conversation:GlobalFlagStatus";
    static const property_path has_attachments = "conversation:HasAttachments";
    static const property_path global_has_attachments =
        "conversation:GlobalHasAttachments";
    static const property_path has_irm = "conversation:HasIrm";
    static const property_path global_has_irm = "conversation:GlobalHasIrm";
    static const property_path message_count = "conversation:MessageCount";
    static const property_path global_message_count =
        "conversation:GlobalMessageCount";
    static const property_path unread_count = "conversation:UnreadCount";
    static const property_path global_unread_count =
        "conversation:GlobalUnreadCount";
    static const property_path size = "conversation:Size";
    static const property_path global_size = "conversation:GlobalSize";
    static const property_path item_classes = "conversation:ItemClasses";
    static const property_path global_item_classes =
        "conversation:GlobalItemClasses";
    static const property_path importance = "conversation:Importance";
    static const property_path global_importance =
        "conversation:GlobalImportance";
    static const property_path item_ids = "conversation:ItemIds";
    static const property_path global_item_ids = "conversation:GlobalItemIds";
    static const property_path last_modified_time =
        "conversation:LastModifiedTime";
    static const property_path instance_key = "conversation:InstanceKey";
    static const property_path preview = "conversation:Preview";
    static const property_path global_parent_folder_id =
        "conversation:GlobalParentFolderId";
    static const property_path next_predicted_action =
        "conversation:NextPredictedAction";
    static const property_path grouping_action = "conversation:GroupingAction";
    static const property_path icon_index = "conversation:IconIndex";
    static const property_path global_icon_index =
        "conversation:GlobalIconIndex";
    static const property_path draft_item_ids = "conversation:DraftItemIds";
    static const property_path has_clutter = "conversation:HasClutter";
}

//! \brief Represents a single property
//!
//! Used in ews::service::update_item method call
class property final
{
public:
    //! Use this constructor if you want to delete a property from an item
    explicit property(property_path path) : value_(path.to_xml()) {}

    // Use this constructor (and following overloads) whenever you want to
    // set or update an item's property
    property(property_path path, std::string value) : value_(path.to_xml(value))
    {
    }

    property(property_path path, const char* value)
        : value_(path.to_xml(std::string(value)))
    {
    }

    property(property_path path, int value)
        : value_(path.to_xml(std::to_string(value)))
    {
    }

    property(property_path path, long value)
        : value_(path.to_xml(std::to_string(value)))
    {
    }

    property(property_path path, long long value)
        : value_(path.to_xml(std::to_string(value)))
    {
    }

    property(property_path path, unsigned value)
        : value_(path.to_xml(std::to_string(value)))
    {
    }

    property(property_path path, unsigned long value)
        : value_(path.to_xml(std::to_string(value)))
    {
    }

    property(property_path path, unsigned long long value)
        : value_(path.to_xml(std::to_string(value)))
    {
    }

    property(property_path path, float value)
        : value_(path.to_xml(std::to_string(value)))
    {
    }

    property(property_path path, double value)
        : value_(path.to_xml(std::to_string(value)))
    {
    }

    property(property_path path, long double value)
        : value_(path.to_xml(std::to_string(value)))
    {
    }

    property(property_path path, bool value)
        : value_(path.to_xml((value ? "true" : "false")))
    {
    }

#ifdef EWS_HAS_DEFAULT_TEMPLATE_ARGS_FOR_FUNCTIONS
    template <typename T,
              typename = typename std::enable_if<std::is_enum<T>::value>::type>
    property(property_path path, T enum_value)
        : value_(path.to_xml(internal::enum_to_str(enum_value)))
    {
    }
#else
    property(property_path path, ews::free_busy_status enum_value)
        : value_(path.to_xml(internal::enum_to_str(enum_value)))
    {
    }

    property(property_path path, ews::sensitivity enum_value)
        : value_(path.to_xml(internal::enum_to_str(enum_value)))
    {
    }
#endif

    property(property_path path, const body& value)
        : value_(path.to_xml(value.to_xml()))
    {
    }

    property(property_path path, const date_time& value)
        : value_(path.to_xml(value.to_string()))
    {
    }

    property(property_path path, const recurrence_pattern& pattern,
             const recurrence_range& range)
        : value_()
    {
        std::stringstream sstr;
        sstr << "<t:Recurrence>" << pattern.to_xml() << range.to_xml()
             << "</t:Recurrence>";
        value_ = path.to_xml(sstr.str());
    }

    template <typename T>
    property(property_path path, const std::vector<T>& value) : value_()
    {
        std::stringstream sstr;
        for (const auto& elem : value)
        {
            sstr << elem.to_xml();
        }
        value_ = path.to_xml(sstr.str());
    }

    property(property_path path, const std::vector<std::string>& value)
        : value_()
    {
        std::stringstream sstr;
        for (const auto& str : value)
        {
            sstr << "<t:String>" << str << "</t:String>";
        }
        value_ = path.to_xml(sstr.str());
    }

    property(const indexed_property_path& path, const physical_address& address)
        : value_(path.to_xml(address.to_xml()))
    {
    }

    property(const indexed_property_path& path, const im_address& address)
        : value_(path.to_xml(address.to_xml()))
    {
    }

    property(const indexed_property_path& path, const email_address& address)
        : value_(path.to_xml(address.to_xml()))
    {
    }

    property(const indexed_property_path& path, const phone_number& number)
        : value_(path.to_xml(number.to_xml()))
    {
    }

    const std::string& to_xml() const EWS_NOEXCEPT { return value_; }

private:
    std::string value_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<property>::value, "");
static_assert(std::is_copy_constructible<property>::value, "");
static_assert(std::is_copy_assignable<property>::value, "");
static_assert(std::is_move_constructible<property>::value, "");
static_assert(std::is_move_assignable<property>::value, "");
#endif

//! \brief Base-class for all search expressions.
//!
//! Search expressions are used to restrict the result set of a
//! \<FindItem/> operation.
//!
//! E.g.
//!
//!   - exists
//!   - excludes
//!   - is_equal_to
//!   - is_not_equal_to
//!   - is_greater_than
//!   - is_greater_than_or_equal_to
//!   - is_less_than
//!   - is_less_than_or_equal_to
//!   - contains
//!   - not
//!   - and
//!   - or
class search_expression
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    search_expression() = delete;
#endif

    std::string to_xml() const { return func_(); }

protected:
    explicit search_expression(std::function<std::string()>&& func)
        : func_(std::move(func))
    {
    }

    // Used by sub-classes to share common code
    search_expression(const char* term, property_path path, bool b)
        : func_([=]() -> std::string {
              std::stringstream sstr;

              sstr << "<t:" << term << ">";
              sstr << path.to_xml();
              sstr << "<t:FieldURIOrConstant>";
              sstr << "<t:Constant Value=\"";
              sstr << std::boolalpha << b;
              sstr << "\"/></t:FieldURIOrConstant></t:";
              sstr << term << ">";
              return sstr.str();
          })
    {
    }

    search_expression(const char* term, property_path path, int i)
        : func_([=]() -> std::string {
              std::stringstream sstr;

              sstr << "<t:" << term << ">";
              sstr << path.to_xml();
              sstr << "<t:"
                   << "FieldURIOrConstant>";
              sstr << "<t:Constant Value=\"";
              sstr << std::to_string(i);
              sstr << "\"/></t:FieldURIOrConstant></t:";
              sstr << term << ">";
              return sstr.str();
          })
    {
    }

    search_expression(const char* term, property_path path, const char* str)
        : func_([=]() -> std::string {
              std::stringstream sstr;

              sstr << "<t:" << term << ">";
              sstr << path.to_xml();
              sstr << "<t:"
                   << "FieldURIOrConstant>";
              sstr << "<t:Constant Value=\"";
              sstr << str;
              sstr << "\"/></t:FieldURIOrConstant></t:";
              sstr << term << ">";
              return sstr.str();
          })
    {
    }

    search_expression(const char* term, indexed_property_path path,
                      const char* str)
        : func_([=]() -> std::string {
              std::stringstream sstr;

              sstr << "<t:" << term << ">";
              sstr << path.to_xml();
              sstr << "<t:FieldURIOrConstant>";
              sstr << "<t:Constant Value=\"";
              sstr << str;
              sstr << "\"/></t:FieldURIOrConstant></t:";
              sstr << term << ">";
              return sstr.str();
          })
    {
    }

    search_expression(const char* term, property_path path, date_time when)
        : func_([=]() -> std::string {
              std::stringstream sstr;

              sstr << "<t:" << term << ">";
              sstr << path.to_xml();
              sstr << "<t:FieldURIOrConstant>";
              sstr << "<t:Constant Value=\"";
              sstr << when.to_string();
              sstr << "\"/></t:FieldURIOrConstant></t:";
              sstr << term << ">";
              return sstr.str();
          })
    {
    }

private:
    std::function<std::string()> func_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<search_expression>::value, "");
static_assert(std::is_copy_constructible<search_expression>::value, "");
static_assert(std::is_copy_assignable<search_expression>::value, "");
static_assert(std::is_move_constructible<search_expression>::value, "");
static_assert(std::is_move_assignable<search_expression>::value, "");
#endif

//! \brief Compare a property with a constant or another property
//!
//! A search expression that compares a property with either a constant
//! value or another property and evaluates to true if they are equal.
class is_equal_to final : public search_expression
{
public:
    is_equal_to(property_path path, bool b)
        : search_expression("IsEqualTo", std::move(path), b)
    {
    }

    is_equal_to(property_path path, int i)
        : search_expression("IsEqualTo", std::move(path), i)
    {
    }

    is_equal_to(property_path path, const char* str)
        : search_expression("IsEqualTo", std::move(path), str)
    {
    }

    is_equal_to(indexed_property_path path, const char* str)
        : search_expression("IsEqualTo", std::move(path), str)
    {
    }

    is_equal_to(property_path path, date_time when)
        : search_expression("IsEqualTo", std::move(path), std::move(when))
    {
    }

    // TODO: is_equal_to(property_path, property_path) {}
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<is_equal_to>::value, "");
static_assert(std::is_copy_constructible<is_equal_to>::value, "");
static_assert(std::is_copy_assignable<is_equal_to>::value, "");
static_assert(std::is_move_constructible<is_equal_to>::value, "");
static_assert(std::is_move_assignable<is_equal_to>::value, "");
#endif

//! \brief Compare a property with a constant or another property.
//!
//! Compares a property with either a constant value or another property
//! and returns true if the values are not the same.
class is_not_equal_to final : public search_expression
{
public:
    is_not_equal_to(property_path path, bool b)
        : search_expression("IsNotEqualTo", std::move(path), b)
    {
    }

    is_not_equal_to(property_path path, int i)
        : search_expression("IsNotEqualTo", std::move(path), i)
    {
    }

    is_not_equal_to(property_path path, const char* str)
        : search_expression("IsNotEqualTo", std::move(path), str)
    {
    }

    is_not_equal_to(indexed_property_path path, const char* str)
        : search_expression("IsNotEqualTo", std::move(path), str)
    {
    }

    is_not_equal_to(property_path path, date_time when)
        : search_expression("IsNotEqualTo", std::move(path), std::move(when))
    {
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<is_not_equal_to>::value, "");
static_assert(std::is_copy_constructible<is_not_equal_to>::value, "");
static_assert(std::is_copy_assignable<is_not_equal_to>::value, "");
static_assert(std::is_move_constructible<is_not_equal_to>::value, "");
static_assert(std::is_move_assignable<is_not_equal_to>::value, "");
#endif

//! \brief Compare a property with a constant or another property
class is_greater_than final : public search_expression
{
public:
    is_greater_than(property_path path, bool b)
        : search_expression("IsGreaterThan", std::move(path), b)
    {
    }

    is_greater_than(property_path path, int i)
        : search_expression("IsGreaterThan", std::move(path), i)
    {
    }

    is_greater_than(property_path path, const char* str)
        : search_expression("IsGreaterThan", std::move(path), str)
    {
    }

    is_greater_than(indexed_property_path path, const char* str)
        : search_expression("IsGreaterThan", std::move(path), str)
    {
    }

    is_greater_than(property_path path, date_time when)
        : search_expression("IsGreaterThan", std::move(path), std::move(when))
    {
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<is_greater_than>::value, "");
static_assert(std::is_copy_constructible<is_greater_than>::value, "");
static_assert(std::is_copy_assignable<is_greater_than>::value, "");
static_assert(std::is_move_constructible<is_greater_than>::value, "");
static_assert(std::is_move_assignable<is_greater_than>::value, "");
#endif

//! \brief Compare a property with a constant or another property
class is_greater_than_or_equal_to final : public search_expression
{
public:
    is_greater_than_or_equal_to(property_path path, bool b)
        : search_expression("IsGreaterThanOrEqualTo", std::move(path), b)
    {
    }

    is_greater_than_or_equal_to(property_path path, int i)
        : search_expression("IsGreaterThanOrEqualTo", std::move(path), i)
    {
    }

    is_greater_than_or_equal_to(property_path path, const char* str)
        : search_expression("IsGreaterThanOrEqualTo", std::move(path), str)
    {
    }

    is_greater_than_or_equal_to(indexed_property_path path, const char* str)
        : search_expression("IsGreaterThanOrEqualTo", std::move(path), str)
    {
    }

    is_greater_than_or_equal_to(property_path path, date_time when)
        : search_expression("IsGreaterThanOrEqualTo", std::move(path),
                            std::move(when))
    {
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(
    !std::is_default_constructible<is_greater_than_or_equal_to>::value, "");
static_assert(std::is_copy_constructible<is_greater_than_or_equal_to>::value,
              "");
static_assert(std::is_copy_assignable<is_greater_than_or_equal_to>::value, "");
static_assert(std::is_move_constructible<is_greater_than_or_equal_to>::value,
              "");
static_assert(std::is_move_assignable<is_greater_than_or_equal_to>::value, "");
#endif

//! \brief Compare a property with a constant or another property
class is_less_than final : public search_expression
{
public:
    is_less_than(property_path path, bool b)
        : search_expression("IsLessThan", std::move(path), b)
    {
    }

    is_less_than(property_path path, int i)
        : search_expression("IsLessThan", std::move(path), i)
    {
    }

    is_less_than(property_path path, const char* str)
        : search_expression("IsLessThan", std::move(path), str)
    {
    }

    is_less_than(indexed_property_path path, const char* str)
        : search_expression("IsLessThan", std::move(path), str)
    {
    }

    is_less_than(property_path path, date_time when)
        : search_expression("IsLessThan", std::move(path), std::move(when))
    {
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<is_less_than>::value, "");
static_assert(std::is_copy_constructible<is_less_than>::value, "");
static_assert(std::is_copy_assignable<is_less_than>::value, "");
static_assert(std::is_move_constructible<is_less_than>::value, "");
static_assert(std::is_move_assignable<is_less_than>::value, "");
#endif

//! \brief Compare a property with a constant or another property
class is_less_than_or_equal_to final : public search_expression
{
public:
    is_less_than_or_equal_to(property_path path, bool b)
        : search_expression("IsLessThanOrEqualTo", std::move(path), b)
    {
    }

    is_less_than_or_equal_to(property_path path, int i)
        : search_expression("IsLessThanOrEqualTo", std::move(path), i)
    {
    }

    is_less_than_or_equal_to(property_path path, const char* str)
        : search_expression("IsLessThanOrEqualTo", std::move(path), str)
    {
    }

    is_less_than_or_equal_to(indexed_property_path path, const char* str)
        : search_expression("IsLessThanOrEqualTo", std::move(path), str)
    {
    }

    is_less_than_or_equal_to(property_path path, date_time when)
        : search_expression("IsLessThanOrEqualTo", std::move(path),
                            std::move(when))
    {
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<is_less_than_or_equal_to>::value,
              "");
static_assert(std::is_copy_constructible<is_less_than_or_equal_to>::value, "");
static_assert(std::is_copy_assignable<is_less_than_or_equal_to>::value, "");
static_assert(std::is_move_constructible<is_less_than_or_equal_to>::value, "");
static_assert(std::is_move_assignable<is_less_than_or_equal_to>::value, "");
#endif

//! \brief Allows you to express a boolean And operation between two search
//! expressions
class and_ final : public search_expression
{
public:
    and_(const search_expression& first, const search_expression& second)
        : search_expression([=]() -> std::string {
              std::stringstream sstr;

              sstr << "<t:And"
                   << ">";
              sstr << first.to_xml();
              sstr << second.to_xml();
              sstr << "</t:And>";
              return sstr.str();
          })
    {
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<and_>::value, "");
static_assert(std::is_copy_constructible<and_>::value, "");
static_assert(std::is_copy_assignable<and_>::value, "");
static_assert(std::is_move_constructible<and_>::value, "");
static_assert(std::is_move_assignable<and_>::value, "");
#endif

//! \brief Allows you to express a logical Or operation between two search
//! expressions
class or_ final : public search_expression
{
public:
    or_(const search_expression& first, const search_expression& second)
        : search_expression([=]() -> std::string {
              std::stringstream sstr;

              sstr << "<t:Or"
                   << ">";
              sstr << first.to_xml();
              sstr << second.to_xml();
              sstr << "</t:Or>";
              return sstr.str();
          })
    {
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<or_>::value, "");
static_assert(std::is_copy_constructible<or_>::value, "");
static_assert(std::is_copy_assignable<or_>::value, "");
static_assert(std::is_move_constructible<or_>::value, "");
static_assert(std::is_move_assignable<or_>::value, "");
#endif

//! Negates the boolean value of the search expression it contains
class not_ final : public search_expression
{
public:
    not_(const search_expression& expr)
        : search_expression([=]() -> std::string {
              std::stringstream sstr;

              sstr << "<t:Not"
                   << ">";
              sstr << expr.to_xml();
              sstr << "</t:Not>";
              return sstr.str();
          })
    {
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<not_>::value, "");
static_assert(std::is_copy_constructible<not_>::value, "");
static_assert(std::is_copy_assignable<not_>::value, "");
static_assert(std::is_move_constructible<not_>::value, "");
static_assert(std::is_move_assignable<not_>::value, "");
#endif

//! \brief Specifies which parts of a text value are compared to a
//! supplied constant value.
//!
//! \sa contains
enum class containment_mode
{
    //! \brief The comparison is between the full string and the constant.
    //!
    //! The property value and the supplied constant are exactly the same.
    full_string,

    //! The comparison is between the string prefix and the constant
    prefixed,

    //! \brief The comparison is between a sub-string of the string and the
    //! constant.
    substring,

    //! \brief The comparison is between a prefix on individual words in
    //! the string and the constant.
    prefix_on_words,

    //! \brief The comparison is between an exact phrase in the string and
    //! the constant.
    exact_phrase
};

namespace internal
{
    inline std::string enum_to_str(containment_mode val)
    {
        switch (val)
        {
        case containment_mode::full_string:
            return "FullString";
        case containment_mode::prefixed:
            return "Prefixed";
        case containment_mode::substring:
            return "Substring";
        case containment_mode::prefix_on_words:
            return "PrefixOnWords";
        case containment_mode::exact_phrase:
            return "ExactPhrase";
        default:
            throw exception("Bad enum value");
        }
    }
}

//! \brief This enumeration determines how case and non-spacing characters
//! are considered when evaluating a text search expression.
//!
//! \sa contains
enum class containment_comparison
{
    //! The strings must exactly be the same
    exact,

    //! The comparison is case-insensitive
    ignore_case,

    //! Non-spacing characters will be ignored during comparison
    ignore_non_spacing_characters,

    //! This is \ref containment_comparison::ignore_case and \ref
    //! containment_comparison::ignore_non_spacing_characters
    loose,
};

namespace internal
{
    inline std::string enum_to_str(containment_comparison val)
    {
        switch (val)
        {
        case containment_comparison::exact:
            return "Exact";
        case containment_comparison::ignore_case:
            return "IgnoreCase";
        case containment_comparison::ignore_non_spacing_characters:
            return "IgnoreNonSpacingCharacters";
        case containment_comparison::loose:
            return "Loose";
        default:
            throw exception("Bad enum value");
        }
    }
}

//! \brief Check if a text property contains a sub-string
//!
//! A search filter that allows you to do text searches on string
//! properties.
class contains final : public search_expression
{
public:
    contains(property_path path, const char* str,
             containment_mode mode = containment_mode::substring,
             containment_comparison comparison = containment_comparison::loose)
        : search_expression([=]() -> std::string {
              std::stringstream sstr;

              sstr << "<t:Contains ";
              sstr << "ContainmentMode=\"" << internal::enum_to_str(mode)
                   << "\" ";
              sstr << "ContainmentComparison=\""
                   << internal::enum_to_str(comparison) << "\">";
              sstr << path.to_xml();
              sstr << "<t:Constant Value=\"";
              sstr << str;
              sstr << "\"/></t:Contains>";
              return sstr.str();
          })
    {
    }
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<contains>::value, "");
static_assert(std::is_copy_constructible<contains>::value, "");
static_assert(std::is_copy_assignable<contains>::value, "");
static_assert(std::is_move_constructible<contains>::value, "");
static_assert(std::is_move_assignable<contains>::value, "");
#endif

//! \brief A range view of appointments in a calendar.
//!
//! Represents a date range view of appointments in calendar folder search
//! operations.
class calendar_view
{
public:
    calendar_view(date_time start_date, date_time end_date)
        : start_date_(std::move(start_date)), end_date_(std::move(end_date)),
          max_entries_returned_(0U), max_entries_set_(false)
    {
    }

    calendar_view(date_time start_date, date_time end_date,
                  std::uint32_t max_entries_returned)
        : start_date_(std::move(start_date)), end_date_(std::move(end_date)),
          max_entries_returned_(max_entries_returned), max_entries_set_(true)
    {
    }

    std::uint32_t get_max_entries_returned() const EWS_NOEXCEPT
    {
        return max_entries_returned_;
    }

    const date_time& get_start_date() const EWS_NOEXCEPT { return start_date_; }

    const date_time& get_end_date() const EWS_NOEXCEPT { return end_date_; }

    std::string to_xml() const
    {
        std::stringstream sstr;

        sstr << "<m:CalendarView ";
        if (max_entries_set_)
        {
            sstr << "MaxEntriesReturned=\""
                 << std::to_string(max_entries_returned_) << "\" ";
        }
        sstr << "StartDate=\"" << start_date_.to_string() << "\" EndDate=\""
             << end_date_.to_string() << "\" />";
        return sstr.str();
    }

private:
    date_time start_date_;
    date_time end_date_;
    std::uint32_t max_entries_returned_;
    bool max_entries_set_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<calendar_view>::value, "");
static_assert(std::is_copy_constructible<calendar_view>::value, "");
static_assert(std::is_copy_assignable<calendar_view>::value, "");
static_assert(std::is_move_constructible<calendar_view>::value, "");
static_assert(std::is_move_assignable<calendar_view>::value, "");
#endif

//! \brief An update to a single property of an item.
//!
//! Represents either a \<SetItemField>, an \<AppendToItemField>, or a
//! \<DeleteItemField> operation.
class update final
{
public:
    enum class operation
    {
        //! \brief Replaces or creates a property.
        //!
        //! Replaces data for a property if the property already exists,
        //! otherwise creates the property and sets its value. The
        //! operation is only applicable to read-write properties.
        set_item_field,

        //! \brief Adds data to an existing property.
        //!
        //! This works only on some properties, such as
        //! \li calendar:OptionalAttendees
        //! \li calendar:RequiredAttendees
        //! \li calendar:Resources
        //! \li item:Body
        //! \li message:ToRecipients
        //! \li message:CcRecipients
        //! \li message:BccRecipients
        //! \li message:ReplyTo
        append_to_item_field,

        //! \brief Removes a property from an item.
        //!
        //! Only applicable to read-write properties.
        delete_item_field,
    };

#ifdef EWS_HAS_DEFAULT_AND_DELETE
    update() = delete;
#endif

    // intentionally not explicit
    update(property prop, operation action = operation::set_item_field)
        : prop_(std::move(prop)), op_(std::move(action))
    {
    }

    //! Serializes this update instance to an XML string
    std::string to_xml() const
    {
        std::string action = "SetItemField";
        if (op_ == operation::append_to_item_field)
        {
            action = "AppendToItemField";
        }
        else if (op_ == operation::delete_item_field)
        {
            action = "DeleteItemField";
        }
        std::stringstream sstr;
        sstr << "<t:" << action << ">";
        sstr << prop_.to_xml();
        sstr << "</t:" << action << ">";
        return sstr.str();
    }

private:
    property prop_;
    update::operation op_;
};

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<update>::value, "");
static_assert(std::is_copy_constructible<update>::value, "");
static_assert(std::is_copy_assignable<update>::value, "");
static_assert(std::is_move_constructible<update>::value, "");
static_assert(std::is_move_assignable<update>::value, "");
#endif

//! \brief Contains the methods to perform operations on an Exchange server
//!
//! The service class contains all methods that can be performed on an
//! Exchange server.
//!
//! Will get a *huge* public interface over time, e.g.,
//!
//! - create_item
//! - find_conversation
//! - find_folder
//! - find_item
//! - find_people
//! - get_contact
//! - get_task
//!
//! and so on and so on.
template <typename RequestHandler = internal::http_request>
class basic_service final
{
public:
#ifdef EWS_HAS_DEFAULT_AND_DELETE
    basic_service() = delete;
#endif

    //! \brief Constructs a new service with given credentials to a server
    //! specified by \p server_uri.
    //!
    //! This constructor will always use NTLM authentication.
    basic_service(const std::string& server_uri, const std::string& domain,
                  const std::string& username, const std::string& password)
        : request_handler_(server_uri), server_version_("Exchange2013_SP1")
    {
        request_handler_.set_method(RequestHandler::method::POST);
        request_handler_.set_content_type("text/xml; charset=utf-8");
        ntlm_credentials creds(username, password, domain);
        request_handler_.set_credentials(creds);
    }

    //! \brief Constructs a new service with given credentials to a server
    //! specified by \p server_uri
    basic_service(const std::string& server_uri, const internal::credentials& creds)
        : request_handler_(server_uri), server_version_("Exchange2013_SP1")
    {
        request_handler_.set_method(RequestHandler::method::POST);
        request_handler_.set_content_type("text/xml; charset=utf-8");
        request_handler_.set_credentials(creds);
    }

    //! \brief Sets the schema version that will be used in requests made
    //! by this service
    void set_request_server_version(server_version vers)
    {
        server_version_ = internal::enum_to_str(vers);
    }

    //! \brief Sets maximum time the request is allowed to take.
    //!
    //! This has been tested and works for short timeout values ( \c <2),
    //! longer periods seem not to work.
    //!
    //! To remove any hard limit on a network communication (the default),
    //! set the timeout to \c 0.
    void set_timeout(std::chrono::seconds d)
    {
        request_handler_.set_timeout(d);
    }

    //! \brief Returns the schema version that is used in requests by this
    //! service
    server_version get_request_server_version() const
    {
        return internal::str_to_server_version(server_version_);
    }

    //! Gets a task from the Exchange store.
    task get_task(const item_id& id)
    {
        return get_item_impl<task>(id, base_shape::all_properties);
    }

    //! \brief Gets a task from the Exchange store.
    //!
    //! The returned task includes specified additional properties.
    task get_task(const item_id& id,
                  const std::vector<property_path>& additional_properties)
    {
        return get_item_impl<task>(id, base_shape::all_properties,
                                   additional_properties);
    }

    //! Gets a contact from the Exchange store.
    contact get_contact(const item_id& id)
    {
        return get_item_impl<contact>(id, base_shape::all_properties);
    }

    //! \brief Gets a contact from the Exchange store.
    //!
    //! The returned contact includes specified additional properties.
    contact get_contact(const item_id& id,
                        const std::vector<property_path>& additional_properties)
    {
        return get_item_impl<contact>(id, base_shape::all_properties,
                                      additional_properties);
    }

    //! Gets a calendar item from the Exchange store.
    calendar_item get_calendar_item(const item_id& id)
    {
        return get_item_impl<calendar_item>(id, base_shape::all_properties);
    }

    //! \brief Gets a calendar item from the Exchange store.
    //!
    //! The returned calendar item includes specified additional
    //! properties.
    calendar_item
    get_calendar_item(const item_id& id,
                      const std::vector<property_path>& additional_properties)
    {
        return get_item_impl<calendar_item>(id, base_shape::all_properties,
                                            additional_properties);
    }

    //! Gets a bunch of calendar items from the Exchange store at once.
    std::vector<calendar_item>
    get_calendar_items(const std::vector<item_id>& ids, base_shape shape)
    {
        return get_item_impl<calendar_item>(ids, shape);
    }

    //! \brief Gets a bunch of calendar items from the Exchange store at once.
    //!
    //! Each of the returned calendar items include the specified additional
    //! properties.
    std::vector<calendar_item>
    get_calendar_items(const std::vector<item_id>& ids, base_shape shape,
                       const std::vector<property_path>& additional_properties)
    {
        return get_item_impl<calendar_item>(ids, shape, additional_properties);
    }

    //! Gets a message item from the Exchange store.
    message get_message(const item_id& id)
    {
        return get_item_impl<message>(id, base_shape::all_properties);
    }

    //! \brief Gets a message from the Exchange store.
    //!
    //! The returned message includes specified additional properties.
    message get_message(const item_id& id,
                        const std::vector<property_path>& additional_properties)
    {
        return get_item_impl<message>(id, base_shape::all_properties,
                                      additional_properties);
    }

    message get_message(const item_id& id,
                        const std::vector<extended_field_uri>& ext_field_uri)
    {
        return get_item_impl<message>(id, ext_field_uri);
    }

    //! Delete an arbitrary item from the Exchange store
    void delete_item(const item_id& id,
                     delete_type del_type = delete_type::hard_delete,
                     affected_task_occurrences affected =
                         affected_task_occurrences::all_occurrences,
                     send_meeting_cancellations cancellations =
                         send_meeting_cancellations::send_to_none)
    {
        using internal::delete_item_response_message;

        const std::string request_string =
            "<m:DeleteItem "
            "DeleteType=\"" +
            internal::enum_to_str(del_type) + "\" "
                                              "SendMeetingCancellations=\"" +
            internal::enum_to_str(cancellations) +
            "\" "
            "AffectedTaskOccurrences=\"" +
            internal::enum_to_str(affected) + "\">"
                                              "<m:ItemIds>" +
            id.to_xml() + "</m:ItemIds>"
                          "</m:DeleteItem>";

        auto response = request(request_string);
        const auto response_message =
            delete_item_response_message::parse(std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
    }

    //! Delete a task item from the Exchange store
    void delete_task(task&& the_task,
                     delete_type del_type = delete_type::hard_delete,
                     affected_task_occurrences affected =
                         affected_task_occurrences::all_occurrences)
    {
        delete_item(the_task.get_item_id(), del_type, affected);
        the_task = ews::task();
    }

    //! Delete a contact from the Exchange store
    void delete_contact(contact&& the_contact)
    {
        delete_item(the_contact.get_item_id());
        the_contact = ews::contact();
    }

    //! Delete a calendar item from the Exchange store
    void delete_calendar_item(calendar_item&& the_calendar_item,
                              delete_type del_type = delete_type::hard_delete,
                              send_meeting_cancellations cancellations =
                                  send_meeting_cancellations::send_to_none)
    {
        delete_item(the_calendar_item.get_item_id(), del_type,
                    affected_task_occurrences::all_occurrences, cancellations);
        the_calendar_item = ews::calendar_item();
    }

    //! Delete a message item from the Exchange store
    void delete_message(message&& the_message)
    {
        delete_item(the_message.get_item_id());
        the_message = ews::message();
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

    //! \brief Creates a new task item from the given object in the
    //! Exchange store.
    //!
    //! Returns it's item_id if successful.
    item_id create_item(const task& the_task)
    {
        return create_item_impl(the_task);
    }

    //! \brief Creates a new contact item from the given object in the
    //! Exchange store
    //!
    //! Returns it's item_id if successful.
    item_id create_item(const contact& the_contact)
    {
        return create_item_impl(the_contact);
    }

    //! \brief Creates a new calendar item from the given object in the
    //! Exchange store
    //!
    //! Returns it's item_id if successful.
    item_id create_item(const calendar_item& the_calendar_item,
                        send_meeting_invitations invitations =
                            send_meeting_invitations::send_to_none)
    {
        using internal::create_item_response_message;

        auto response =
            request(the_calendar_item.create_item_request_string(invitations));

        const auto response_message =
            create_item_response_message::parse(std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        EWS_ASSERT(!response_message.items().empty() &&
                   "Expected a message item");
        return response_message.items().front();
    }

    //! \brief Creates a new message in the Exchange store.
    //!
    //! Creates a new message and, depending of the chosen message
    //! disposition, sends it to the recipients.
    //!
    //! Note that if you pass message_disposition::send_only or
    //! message_disposition::send_and_save_copy this function always
    //! returns an invalid item id because Exchange does not include
    //! the item identifier in the response. A common workaround for this
    //! would be to create the item with message_disposition::save_only, get
    //! the item identifier, and then use the send_item to send the message.
    //!
    //! \return The item id of the saved message when
    //! message_disposition::save_only was given; otherwise an invalid item
    //! id.
    item_id create_item(const message& the_message,
                        ews::message_disposition disposition)
    {
        using internal::create_item_response_message;

        auto response =
            request(the_message.create_item_request_string(disposition));

        const auto response_message =
            create_item_response_message::parse(std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }

        if (disposition == message_disposition::save_only)
        {
            EWS_ASSERT(!response_message.items().empty() &&
                       "Expected a message item");
            return response_message.items().front();
        }

        return item_id();
    }

    //! Sends a message that is already in the Exchange store.
    //!
    //! \param id The item id of the message you want to send
    void send_item(const item_id& id) { send_item(id, folder_id()); }

    //! \brief Sends a message that is already in the Exchange store.
    //!
    //! \param id The item id of the message you want to send
    //! \param folder The folder in the mailbox in which the send message is
    //! saved. If you pass an invalid id here, the message won't be saved.
    void send_item(const item_id& id, const folder_id& folder)
    {
        std::stringstream sstr;
        sstr << "<m:SendItem SaveItemToFolder=\"" << std::boolalpha
             << folder.valid() << "\">"
             << "<m:ItemIds>" << id.to_xml() << "</m:ItemIds>";
        if (folder.valid())
        {
            sstr << "<m:SavedItemFolderId>" << folder.to_xml()
                 << "</m:SavedItemFolderId>";
        }
        sstr << "</m:SendItem>";

        auto response = request(sstr.str());

        const auto response_message =
            internal::send_item_response_message::parse(std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
    }

    std::vector<item_id> find_item(const folder_id& parent_folder_id)
    {
        const std::string request_string = "<m:FindItem Traversal=\"Shallow\">"
                                           "<m:ItemShape>"
                                           "<t:BaseShape>IdOnly</t:BaseShape>"
                                           "</m:ItemShape>"
                                           "<m:ParentFolderIds>" +
                                           parent_folder_id.to_xml() +
                                           "</m:ParentFolderIds>"
                                           "</m:FindItem>";

        auto response = request(request_string);
        const auto response_message =
            internal::find_item_response_message::parse(std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        return response_message.items();
    }

    //! \brief Returns all calendar items in given calendar view.
    //!
    //! Sends a \<FindItem/> operation to the server containing a
    //! \<CalendarView/> element. It returns single calendar items and all
    //! occurrences of recurring meetings.
    std::vector<calendar_item> find_item(const calendar_view& view,
                                         const folder_id& parent_folder_id,
                                         base_shape shape = base_shape::id_only)
    {
        const std::string request_string =
            "<m:FindItem Traversal=\"Shallow\">"
            "<m:ItemShape>"
            "<t:BaseShape>" +
            internal::enum_to_str(shape) + "</t:BaseShape>"
                                           "</m:ItemShape>" +
            view.to_xml() + "<m:ParentFolderIds>" + parent_folder_id.to_xml() +
            "</m:ParentFolderIds>"
            "</m:FindItem>";

        auto response = request(request_string);
        const auto response_message =
            internal::find_calendar_item_response_message::parse(response);
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        return response_message.items();
    }

    //! \brief Sends a \<FindItem/> operation to the server
    //!
    //! Allows you to search for items that are located in a user's
    //! mailbox.
    //!
    //! \param parent_folder_id The folder in the mailbox that is searched
    //! \param restriction A search expression that restricts the elements
    //! returned by this operation
    //!
    //! Returns a list of items (item_ids) that match given folder and
    //! restrictions
    std::vector<item_id> find_item(const folder_id& parent_folder_id,
                                   search_expression restriction)
    {
        const std::string request_string =
            "<m:FindItem Traversal=\"Shallow\">"
            "<m:ItemShape>"
            "<t:BaseShape>IdOnly</t:BaseShape>"
            "</m:ItemShape>"
            "<m:Restriction>" +
            restriction.to_xml() + "</m:Restriction>"
                                   "<m:ParentFolderIds>" +
            parent_folder_id.to_xml() + "</m:ParentFolderIds>"
                                        "</m:FindItem>";

        auto response = request(request_string);
        const auto response_message =
            internal::find_item_response_message::parse(std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        return response_message.items();
    }

    item_id
    update_item(item_id id, update change,
                conflict_resolution res = conflict_resolution::auto_resolve,
                send_meeting_cancellations cancellations =
                    send_meeting_cancellations::send_to_none)
    {
        const std::string request_string =
            "<m:UpdateItem "
            "MessageDisposition=\"SaveOnly\" "
            "ConflictResolution=\"" +
            internal::enum_to_str(res) +
            "\" "
            "SendMeetingInvitationsOrCancellations=\"" +
            internal::enum_to_str(cancellations) + "\">"
                                                   "<m:ItemChanges>"
                                                   "<t:ItemChange>" +
            id.to_xml() + "<t:Updates>" + change.to_xml() + "</t:Updates>"
                                                            "</t:ItemChange>"
                                                            "</m:ItemChanges>"
                                                            "</m:UpdateItem>";

        auto response = request(request_string);
        const auto response_message =
            internal::update_item_response_message::parse(std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        EWS_ASSERT(!response_message.items().empty() &&
                   "Expected at least one item");
        return response_message.items().front();
    }

    item_id
    update_item(item_id id, const std::vector<update>& changes,
                conflict_resolution res = conflict_resolution::auto_resolve,
                send_meeting_cancellations cancellations =
                    send_meeting_cancellations::send_to_none)
    {
        std::string request_string =
            "<m:UpdateItem "
            "MessageDisposition=\"SaveOnly\" "
            "ConflictResolution=\"" +
            internal::enum_to_str(res) +
            "\" "
            "SendMeetingInvitationsOrCancellations=\"" +
            internal::enum_to_str(cancellations) + "\">"
                                                   "<m:ItemChanges>"
                                                   "<t:ItemChange>" +
            id.to_xml() + "<t:Updates>";

        for (const auto& change : changes)
        {
            request_string += change.to_xml();
        }

        request_string += "</t:Updates>"
                          "</t:ItemChange>"
                          "</m:ItemChanges>"
                          "</m:UpdateItem>";

        auto response = request(request_string);
        const auto response_message =
            internal::update_item_response_message::parse(std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        EWS_ASSERT(!response_message.items().empty() &&
                   "Expected at least one item");
        return response_message.items().front();
    }

    //! \brief Lets you attach a file (or another item) to an existing item.
    //!
    //! \param parent_item An existing item in the Exchange store
    //! \param a The <tt>\<FileAttachment></tt> or
    //!        <tt>\<ItemAttachment></tt> you want to attach to
    //!        \p parent_item
    attachment_id create_attachment(const item_id& parent_item,
                                    const attachment& a)
    {
        auto response = request("<m:CreateAttachment>"
                                "<m:ParentItemId Id=\"" +
                                parent_item.id() + "\" ChangeKey=\"" +
                                parent_item.change_key() + "\"/>"
                                                           "<m:Attachments>" +
                                a.to_xml() + "</m:Attachments>"
                                             "</m:CreateAttachment>");

        const auto response_message =
            internal::create_attachment_response_message::parse(
                std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        EWS_ASSERT(!response_message.attachment_ids().empty() &&
                   "Expected at least one attachment");
        return response_message.attachment_ids().front();
    }

#if 0
        // TODO

        //! \brief Attach one or more files (or items) to an existing item.
        //!
        //! Use this member-function when you want to create multiple
        //! attachments to a single item. If you want to create attachments for
        //! multiple items, you will need to call create_attachment several
        //! times--one for each item that you want to add attachments to.
        std::vector<attachment_id>
        create_attachments(const item& parent_item,
                           const std::vector<attachment>& attachments)
        {
            (void)parent_item;
            (void)attachments;
            EWS_ASSERT(false);
            return std::vector<attachment_id>();
        }
#endif // 0

    //! Retrieves an attachment from the Exchange store
    attachment get_attachment(const attachment_id& id)
    {
        auto response = request("<m:GetAttachment>"
                                "<m:AttachmentShape>"
                                "<m:IncludeMimeContent/>"
                                "<m:BodyType/>"
                                "<m:FilterHtmlContent/>"
                                "<m:AdditionalProperties/>"
                                "</m:AttachmentShape>"
                                "<m:AttachmentIds>" +
                                id.to_xml() + "</m:AttachmentIds>"
                                              "</m:GetAttachment>");

        const auto response_message =
            internal::get_attachment_response_message::parse(
                std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        EWS_ASSERT(!response_message.attachments().empty() &&
                   "Expected at least one attachment to be returned");
        return response_message.attachments().front();
    }

    //! \brief Deletes given attachment from the Exchange store
    //!
    //! Returns the item_id of the parent item from which the attachment
    //! was removed (also known as <em>root</em> item). This item_id
    //! contains the updated change key of the parent item.
    item_id delete_attachment(const attachment_id& id)
    {
        auto response = request("<m:DeleteAttachment>"
                                "<m:AttachmentIds>" +
                                id.to_xml() + "</m:AttachmentIds>"
                                              "</m:DeleteAttachment>");
        const auto response_message =
            internal::delete_attachment_response_message::parse(
                std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        return response_message.get_root_item_id();
    }

private:
    RequestHandler request_handler_;
    std::string server_version_;

    // Helper for doing requests.  Adds the right headers, credentials, and
    // checks the response for faults.
    internal::http_response request(const std::string& request_string)
    {
        using rapidxml::internal::compare;
        using internal::get_element_by_qname;

        auto soap_headers = std::vector<std::string>();
        soap_headers.emplace_back("<t:RequestServerVersion Version=\"" +
                                  server_version_ + "\"/>");

        auto response = internal::make_raw_soap_request(
            request_handler_, request_string, soap_headers);

        if (response.ok())
        {
            return response;
        }
        else if (response.is_soap_fault())
        {
            std::unique_ptr<rapidxml::xml_document<char>> doc;

            try
            {
                doc = parse_response(std::move(response));
            }
            catch (xml_parse_error&)
            {
                throw soap_fault("The request failed for unknown reason "
                                 "(could not parse response)");
            }

            auto elem = get_element_by_qname(
                *doc, "ResponseCode", internal::uri<>::microsoft::errors());
            if (!elem)
            {
                throw soap_fault("The request failed for unknown reason "
                                 "(unexpected XML in response)");
            }

            if (compare(elem->value(), elem->value_size(),
                        "ErrorSchemaValidation", 21))
            {
                // Get some more helpful details
                elem = get_element_by_qname(
                    *doc, "LineNumber", internal::uri<>::microsoft::types());
                EWS_ASSERT(elem && "Expected <LineNumber> element in response");
                const auto line_number =
                    std::stoul(std::string(elem->value(), elem->value_size()));

                elem = get_element_by_qname(
                    *doc, "LinePosition", internal::uri<>::microsoft::types());
                EWS_ASSERT(elem &&
                           "Expected <LinePosition> element in response");
                const auto line_position =
                    std::stoul(std::string(elem->value(), elem->value_size()));

                elem = get_element_by_qname(
                    *doc, "Violation", internal::uri<>::microsoft::types());
                EWS_ASSERT(elem && "Expected <Violation> element in response");
                throw schema_validation_error(
                    line_number, line_position,
                    std::string(elem->value(), elem->value_size()));
            }
            else
            {
                elem = get_element_by_qname(*doc, "faultstring", "");
                EWS_ASSERT(elem &&
                           "Expected <faultstring> element in response");
                throw soap_fault(elem->value());
            }
        }
        else
        {
            throw http_error(response.code());
        }
    }

    // Gets an item from the server
    template <typename ItemType>
    ItemType get_item_impl(const item_id& id, base_shape shape)
    {
        const std::string request_string = "<m:GetItem>"
                                           "<m:ItemShape>"
                                           "<t:BaseShape>" +
                                           internal::enum_to_str(shape) +
                                           "</t:BaseShape>"
                                           "</m:ItemShape>"
                                           "<m:ItemIds>" +
                                           id.to_xml() + "</m:ItemIds>"
                                                         "</m:GetItem>";

        auto response = request(request_string);
        const auto response_message =
            internal::get_item_response_message<ItemType>::parse(
                std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        EWS_ASSERT(!response_message.items().empty() &&
                   "Expected at least one item");
        return response_message.items().front();
    }

    // Gets an item from the server with additional properties
    template <typename ItemType>
    ItemType
    get_item_impl(const item_id& id, base_shape shape,
                  const std::vector<property_path>& additional_properties)
    {
        EWS_ASSERT(!additional_properties.empty());

        std::stringstream sstr;
        sstr << "<m:GetItem>"
                "<m:ItemShape>"
                "<t:BaseShape>"
             << internal::enum_to_str(shape) << "</t:BaseShape>"
                                                "<t:AdditionalProperties>";
        for (const auto& prop : additional_properties)
        {
            sstr << prop.to_xml();
        }
        sstr << "</t:AdditionalProperties>"
                "</m:ItemShape>"
                "<m:ItemIds>"
             << id.to_xml() << "</m:ItemIds>"
                               "</m:GetItem>";

        auto response = request(sstr.str());
        const auto response_message =
            internal::get_item_response_message<ItemType>::parse(
                std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        EWS_ASSERT(!response_message.items().empty() &&
                   "Expected at least one item");
        return response_message.items().front();
    }

    // Gets a bunch of items from the server all at once
    template <typename ItemType>
    std::vector<ItemType> get_item_impl(const std::vector<item_id>& ids,
                                        base_shape shape)
    {
        EWS_ASSERT(!ids.empty());

        std::stringstream sstr;
        sstr << "<m:GetItem>"
                "<m:ItemShape>"
                "<t:BaseShape>"
             << internal::enum_to_str(shape) << "</t:BaseShape>"
                                                "</m:ItemShape>"
                                                "<m:ItemIds>";
        for (const auto& id : ids)
        {
            sstr << id.to_xml();
        }
        sstr << "</m:ItemIds>"
                "</m:GetItem>";

        auto response = request(sstr.str());
        const auto response_messages =
            internal::get_item_response_messages<ItemType>::parse(
                std::move(response));
        if (!response_messages.success())
        {
            throw exchange_error(response_messages.first_error_or_warning());
        }
        return response_messages.items();
    }

    // Gets a bunch of items from the server all at once including given
    // additional properties
    template <typename ItemType>
    std::vector<ItemType>
    get_item_impl(const std::vector<item_id>& ids, base_shape shape,
                  const std::vector<property_path>& additional_properties)
    {
        EWS_ASSERT(!ids.empty());
        EWS_ASSERT(!additional_properties.empty());

        std::stringstream sstr;
        sstr << "<m:GetItem>"
                "<m:ItemShape>"
                "<t:BaseShape>"
             << internal::enum_to_str(shape) << "</t:BaseShape>"
                                                "<t:AdditionalProperties>";
        for (const auto& prop : additional_properties)
        {
            sstr << prop.to_xml();
        }
        sstr << "</t:AdditionalProperties>"
                "</m:ItemShape>"
                "<m:ItemIds>";
        for (const auto& id : ids)
        {
            sstr << id.to_xml();
        }
        sstr << "</m:ItemIds>"
                "</m:GetItem>";

        auto response = request(sstr.str());
        const auto response_messages =
            internal::get_item_response_messages<ItemType>::parse(
                std::move(response));
        if (!response_messages.success())
        {
            throw exchange_error(response_messages.first_error_or_warning());
        }
        return response_messages.items();
    }

    template <typename ItemType>
    ItemType
    get_item_impl(const item_id& id,
                  const std::vector<extended_field_uri>& ext_field_uris)
    {
        EWS_ASSERT(!ext_field_uris.empty());

        std::stringstream sstr;
        sstr << "<m:GetItem>"
                "<m:ItemShape>"
                "<t:BaseShape>AllProperties</t:BaseShape>"
                "<t:AdditionalProperties>";
        for (auto item : ext_field_uris)
        {
            sstr << item.to_xml();
        }
        sstr << "</t:AdditionalProperties>"
                "</m:ItemShape>"
                "<m:ItemIds>"
             << id.to_xml() << "</m:ItemIds>"
                               "</m:GetItem>";

        auto response = request(sstr.str());
        const auto response_message =
            internal::get_item_response_message<ItemType>::parse(
                std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        EWS_ASSERT(!response_message.items().empty() &&
                   "Expected at least one item");
        return response_message.items().front();
    }

    // Creates an item on the server and returns it's item_id.
    template <typename ItemType>
    item_id create_item_impl(const ItemType& the_item)
    {
        using internal::create_item_response_message;

        auto response = request(the_item.create_item_request_string());
        const auto response_message =
            create_item_response_message::parse(std::move(response));
        if (!response_message.success())
        {
            throw exchange_error(response_message.get_response_code());
        }
        EWS_ASSERT(!response_message.items().empty() &&
                   "Expected at least one item");
        return response_message.items().front();
    }
};

typedef basic_service<> service;

#ifdef EWS_HAS_NON_BUGGY_TYPE_TRAITS
static_assert(!std::is_default_constructible<service>::value, "");
static_assert(!std::is_copy_constructible<service>::value, "");
static_assert(!std::is_copy_assignable<service>::value, "");
static_assert(std::is_move_constructible<service>::value, "");
static_assert(std::is_move_assignable<service>::value, "");
#endif

// Implementations

inline void basic_credentials::certify(internal::http_request* request) const
{
    EWS_ASSERT(request != nullptr);

    std::string login = username_ + ":" + password_;
    request->set_option(CURLOPT_USERPWD, login.c_str());
    request->set_option(CURLOPT_HTTPAUTH, CURLAUTH_BASIC);
}

inline void ntlm_credentials::certify(internal::http_request* request) const
{
    EWS_ASSERT(request != nullptr);

    // CURLOPT_USERPWD: domain\username:password
    std::string login = domain_ + "\\" + username_ + ":" + password_;
    request->set_option(CURLOPT_USERPWD, login.c_str());
    request->set_option(CURLOPT_HTTPAUTH, CURLAUTH_NTLM);
}

namespace internal
{
    // FIXME: a CreateItemResponse can contain multiple ResponseMessages
    inline create_item_response_message
    create_item_response_message::parse(http_response&& response)
    {
        const auto doc = parse_response(std::move(response));
        auto elem = get_element_by_qname(*doc, "CreateItemResponseMessage",
                                         uri<>::microsoft::messages());

        EWS_ASSERT(elem && "Expected <CreateItemResponseMessage>, got nullptr");
        const auto result = parse_response_class_and_code(*elem);
        auto item_ids = std::vector<item_id>();
        auto items_elem =
            elem->first_node_ns(uri<>::microsoft::messages(), "Items");
        EWS_ASSERT(items_elem && "Expected <Items> element");

        for_each_child_node(
            *items_elem, [&item_ids](const rapidxml::xml_node<>& item_elem) {
                auto item_id_elem = item_elem.first_node();
                EWS_ASSERT(item_id_elem && "Expected <ItemId> element");
                item_ids.emplace_back(item_id::from_xml_element(*item_id_elem));
            });
        return create_item_response_message(result.first, result.second,
                                            std::move(item_ids));
    }

    inline find_item_response_message
    find_item_response_message::parse(http_response&& response)
    {
        const auto doc = parse_response(std::move(response));
        auto elem = get_element_by_qname(*doc, "FindItemResponseMessage",
                                         uri<>::microsoft::messages());

        EWS_ASSERT(elem && "Expected <FindItemResponseMessage>, got nullptr");
        const auto result = parse_response_class_and_code(*elem);

        auto root_folder =
            elem->first_node_ns(uri<>::microsoft::messages(), "RootFolder");

        auto items_elem =
            root_folder->first_node_ns(uri<>::microsoft::types(), "Items");
        EWS_ASSERT(items_elem && "Expected <t:Items> element");

        auto items = std::vector<item_id>();
        for (auto item_elem = items_elem->first_node(); item_elem;
             item_elem = item_elem->next_sibling())
        {
            EWS_ASSERT(item_elem && "Expected an element, got nullptr");
            auto item_id_elem = item_elem->first_node();
            EWS_ASSERT(item_id_elem && "Expected <ItemId> element");
            items.emplace_back(item_id::from_xml_element(*item_id_elem));
        }
        return find_item_response_message(result.first, result.second,
                                          std::move(items));
    }

    inline find_calendar_item_response_message
    find_calendar_item_response_message::parse(http_response& response)
    {
        const auto doc = parse_response(std::move(response));
        auto elem = get_element_by_qname(*doc, "FindItemResponseMessage",
                                         uri<>::microsoft::messages());

        EWS_ASSERT(elem && "Expected <FindItemResponseMessage>, got nullptr");
        const auto result = parse_response_class_and_code(*elem);

        auto root_folder =
            elem->first_node_ns(uri<>::microsoft::messages(), "RootFolder");

        auto items_elem =
            root_folder->first_node_ns(uri<>::microsoft::types(), "Items");
        EWS_ASSERT(items_elem && "Expected <t:Items> element");

        auto items = std::vector<calendar_item>();
        for_each_child_node(
            *items_elem, [&items](const rapidxml::xml_node<>& item_elem) {
                items.emplace_back(calendar_item::from_xml_element(item_elem));
            });
        return find_calendar_item_response_message(result.first, result.second,
                                                   std::move(items));
    }

    inline update_item_response_message
    update_item_response_message::parse(http_response&& response)
    {
        const auto doc = parse_response(std::move(response));
        auto elem = get_element_by_qname(*doc, "UpdateItemResponseMessage",
                                         uri<>::microsoft::messages());

        EWS_ASSERT(elem && "Expected <UpdateItemResponseMessage>, got nullptr");
        const auto result = parse_response_class_and_code(*elem);

        auto items_elem =
            elem->first_node_ns(uri<>::microsoft::messages(), "Items");
        EWS_ASSERT(items_elem && "Expected <m:Items> element");

        auto items = std::vector<item_id>();
        for (auto item_elem = items_elem->first_node(); item_elem;
             item_elem = item_elem->next_sibling())
        {
            EWS_ASSERT(item_elem && "Expected an element, got nullptr");
            auto item_id_elem = item_elem->first_node();
            EWS_ASSERT(item_id_elem && "Expected <ItemId> element");
            items.emplace_back(item_id::from_xml_element(*item_id_elem));
        }
        return update_item_response_message(result.first, result.second,
                                            std::move(items));
    }

    template <typename ItemType>
    inline get_item_response_message<ItemType>
    get_item_response_message<ItemType>::parse(http_response&& response)
    {
        const auto doc = parse_response(std::move(response));
        auto elem = get_element_by_qname(*doc, "GetItemResponseMessage",
                                         uri<>::microsoft::messages());
        EWS_ASSERT(elem && "Expected <GetItemResponseMessage>, got nullptr");
        const auto result = parse_response_class_and_code(*elem);
        auto items_elem =
            elem->first_node_ns(uri<>::microsoft::messages(), "Items");
        EWS_ASSERT(items_elem && "Expected <Items> element");
        auto items = std::vector<ItemType>();
        for_each_child_node(
            *items_elem, [&items](const rapidxml::xml_node<>& item_elem) {
                items.emplace_back(ItemType::from_xml_element(item_elem));
            });
        return get_item_response_message(result.first, result.second,
                                         std::move(items));
    }

    template <typename ItemType>
    inline get_item_response_messages<ItemType>
    get_item_response_messages<ItemType>::parse(http_response&& response)
    {
        const auto doc = parse_response(std::move(response));

        auto response_messages = get_element_by_qname(
            *doc, "ResponseMessages", uri<>::microsoft::messages());
        EWS_ASSERT(response_messages &&
                   "Expected <ResponseMessages> node, got nullptr");

        std::vector<get_item_response_messages::response_message> messages;
        for_each_child_node(*response_messages, [&](const rapidxml::xml_node<>&
                                                        response_message) {
            auto result = parse_response_class_and_code(response_message);

            auto items_elem = response_message.first_node_ns(
                uri<>::microsoft::messages(), "Items");
            EWS_ASSERT(items_elem && "Expected <Items> element");

            auto items = std::vector<ItemType>();
            for_each_child_node(
                *items_elem, [&items](const rapidxml::xml_node<>& item_elem) {
                    items.emplace_back(ItemType::from_xml_element(item_elem));
                });

            messages.emplace_back(
                std::make_tuple(result.first, result.second, std::move(items)));
        });

        return get_item_response_messages(std::move(messages));
    }
}

//! \brief Creates a new <tt>\<ItemAttachment></tt> from a given item.
//!
//! It is not necessary for the item to already exist in the Exchange
//! store. If it doesn't, it will be automatically created.
inline attachment attachment::from_item(const item& the_item,
                                        const std::string& name)
{
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
// There is a shortcut for Calendar, E-mail message items, and Posting
// notes: use <BaseShape>IdOnly</BaseShape> and <AdditionalProperties>
// with item::MimeContent in <GetItem> call, remove <ItemId> from the
// response and pass that to <CreateAttachment>.

#if 0
        const auto item_class = the_item.get_item_class();
        if (       item_class == ""          // Calendar items
                || item_class == "IPM.Note"  // E-mail messages
                || item_class == "IPM.Post"  // Posting notes in a folder
            )
        {
            auto mime_content_node = props.get_node("MimeContent");
        }
#endif

    auto& props = the_item.xml();

// Filter out read-only property paths
#ifdef EWS_HAS_INITIALIZER_LISTS
    for (const auto& property_name :
         // item
         {"ItemId", "ParentFolderId", "DateTimeReceived", "Size", "IsSubmitted",
          "IsDraft", "IsFromMe", "IsResend", "IsUnmodified", "DateTimeSent",
          "DateTimeCreated", "ResponseObjects", "DisplayCc", "DisplayTo",
          "HasAttachments", "EffectiveRights", "LastModifiedName",
          "LastModifiedTime", "IsAssociated", "WebClientReadFormQueryString",
          "WebClientEditFormQueryString", "ConversationId", "InstanceKey",

          // message
          "ConversationIndex", "ConversationTopic"})
#else
    std::vector<std::string> properties;

    // item
    properties.push_back("ItemId");
    properties.push_back("ParentFolderId");
    properties.push_back("DateTimeReceived");
    properties.push_back("Size");
    properties.push_back("IsSubmitted");
    properties.push_back("IsDraft");
    properties.push_back("IsFromMe");
    properties.push_back("IsResend");
    properties.push_back("IsUnmodified");
    properties.push_back("DateTimeSent");
    properties.push_back("DateTimeCreated");
    properties.push_back("ResponseObjects");
    properties.push_back("DisplayCc");
    properties.push_back("DisplayTo");
    properties.push_back("HasAttachments");
    properties.push_back("EffectiveRights");
    properties.push_back("LastModifiedName");
    properties.push_back("LastModifiedTime");
    properties.push_back("IsAssociated");
    properties.push_back("WebClientReadFormQueryString");
    properties.push_back("WebClientEditFormQueryString");
    properties.push_back("ConversationId");
    properties.push_back("InstanceKey");

    // message
    properties.push_back("ConversationIndex");
    properties.push_back("ConversationTopic");

    for (const auto& property_name : properties)
#endif
    {
        auto node = props.get_node(property_name);
        if (node)
        {
            auto parent = node->parent();
            EWS_ASSERT(parent);
            parent->remove_node(node);
        }
    }

    auto obj = attachment();
    obj.type_ = type::item;

    using internal::create_node;
    auto& attachment_node =
        create_node(*obj.xml_.document(), "t:ItemAttachment");
    create_node(attachment_node, "t:Name", name);
    props.append_to(attachment_node);

    return obj;
}

inline std::unique_ptr<recurrence_pattern>
recurrence_pattern::from_xml_element(const rapidxml::xml_node<>& elem)
{
    using rapidxml::internal::compare;
    using namespace internal;

    EWS_ASSERT(std::string(elem.local_name()) == "Recurrence" &&
               "Expected a <Recurrence> element");

    auto node = elem.first_node_ns(uri<>::microsoft::types(),
                                   "AbsoluteYearlyRecurrence");
    if (node)
    {
        auto mon = ews::month::jan;
        std::uint32_t day_of_month = 0U;

        for (auto child = node->first_node(); child;
             child = child->next_sibling())
        {
            if (compare(child->local_name(), child->local_name_size(), "Month",
                        std::strlen("Month")))
            {
                mon = str_to_month(
                    std::string(child->value(), child->value_size()));
            }
            else if (compare(child->local_name(), child->local_name_size(),
                             "DayOfMonth", std::strlen("DayOfMonth")))
            {
                day_of_month = std::stoul(
                    std::string(child->value(), child->value_size()));
            }
        }

#ifdef EWS_HAS_MAKE_UNIQUE
        return std::make_unique<absolute_yearly_recurrence>(
            std::move(day_of_month), std::move(mon));
#else
        return std::unique_ptr<absolute_yearly_recurrence>(
            new absolute_yearly_recurrence(std::move(day_of_month),
                                           std::move(mon)));
#endif
    }

    node = elem.first_node_ns(uri<>::microsoft::types(),
                              "RelativeYearlyRecurrence");
    if (node)
    {
        auto mon = ews::month::jan;
        auto index = day_of_week_index::first;
        auto days_of_week = day_of_week::sun;

        for (auto child = node->first_node(); child;
             child = child->next_sibling())
        {
            if (compare(child->local_name(), child->local_name_size(), "Month",
                        std::strlen("Month")))
            {
                mon = str_to_month(
                    std::string(child->value(), child->value_size()));
            }
            else if (compare(child->local_name(), child->local_name_size(),
                             "DayOfWeekIndex", std::strlen("DayOfWeekIndex")))
            {
                index = str_to_day_of_week_index(
                    std::string(child->value(), child->value_size()));
            }
            else if (compare(child->local_name(), child->local_name_size(),
                             "DaysOfWeek", std::strlen("DaysOfWeek")))
            {
                days_of_week = str_to_day_of_week(
                    std::string(child->value(), child->value_size()));
            }
        }

#ifdef EWS_HAS_MAKE_UNIQUE
        return std::make_unique<relative_yearly_recurrence>(
            std::move(days_of_week), std::move(index), std::move(mon));
#else
        return std::unique_ptr<relative_yearly_recurrence>(
            new relative_yearly_recurrence(std::move(days_of_week),
                                           std::move(index), std::move(mon)));
#endif
    }

    node = elem.first_node_ns(uri<>::microsoft::types(),
                              "AbsoluteMonthlyRecurrence");
    if (node)
    {
        std::uint32_t interval = 0U;
        std::uint32_t day_of_month = 0U;

        for (auto child = node->first_node(); child;
             child = child->next_sibling())
        {
            if (compare(child->local_name(), child->local_name_size(),
                        "Interval", std::strlen("Interval")))
            {
                interval = std::stoul(
                    std::string(child->value(), child->value_size()));
            }
            else if (compare(child->local_name(), child->local_name_size(),
                             "DayOfMonth", std::strlen("DayOfMonth")))
            {
                day_of_month = std::stoul(
                    std::string(child->value(), child->value_size()));
            }
        }

#ifdef EWS_HAS_MAKE_UNIQUE
        return std::make_unique<absolute_monthly_recurrence>(
            std::move(interval), std::move(day_of_month));
#else
        return std::unique_ptr<absolute_monthly_recurrence>(
            new absolute_monthly_recurrence(std::move(interval),
                                            std::move(day_of_month)));
#endif
    }

    node = elem.first_node_ns(uri<>::microsoft::types(),
                              "RelativeMonthlyRecurrence");
    if (node)
    {
        std::uint32_t interval = 0U;
        auto days_of_week = day_of_week::sun;
        auto index = day_of_week_index::first;

        for (auto child = node->first_node(); child;
             child = child->next_sibling())
        {
            if (compare(child->local_name(), child->local_name_size(),
                        "Interval", std::strlen("Interval")))
            {
                interval = std::stoul(
                    std::string(child->value(), child->value_size()));
            }
            else if (compare(child->local_name(), child->local_name_size(),
                             "DaysOfWeek", std::strlen("DaysOfWeek")))
            {
                days_of_week = str_to_day_of_week(
                    std::string(child->value(), child->value_size()));
            }
            else if (compare(child->local_name(), child->local_name_size(),
                             "DayOfWeekIndex", std::strlen("DayOfWeekIndex")))
            {
                index = str_to_day_of_week_index(
                    std::string(child->value(), child->value_size()));
            }
        }

#ifdef EWS_HAS_MAKE_UNIQUE
        return std::make_unique<relative_monthly_recurrence>(
            std::move(interval), std::move(days_of_week), std::move(index));
#else
        return std::unique_ptr<relative_monthly_recurrence>(
            new relative_monthly_recurrence(std::move(interval),
                                            std::move(days_of_week),
                                            std::move(index)));
#endif
    }

    node = elem.first_node_ns(uri<>::microsoft::types(), "WeeklyRecurrence");
    if (node)
    {
        std::uint32_t interval = 0U;
        auto days_of_week = std::vector<day_of_week>();
        auto first_day_of_week = day_of_week::mon;

        for (auto child = node->first_node(); child;
             child = child->next_sibling())
        {
            if (compare(child->local_name(), child->local_name_size(),
                        "Interval", std::strlen("Interval")))
            {
                interval = std::stoul(
                    std::string(child->value(), child->value_size()));
            }
            else if (compare(child->local_name(), child->local_name_size(),
                             "DaysOfWeek", std::strlen("DaysOfWeek")))
            {
                const auto list =
                    std::string(child->value(), child->value_size());
                std::stringstream sstr(list);
                std::string temp;
                while (std::getline(sstr, temp, ' '))
                {
                    days_of_week.emplace_back(str_to_day_of_week(temp));
                }
            }
            else if (compare(child->local_name(), child->local_name_size(),
                             "FirstDayOfWeek", std::strlen("FirstDayOfWeek")))
            {
                first_day_of_week = str_to_day_of_week(
                    std::string(child->value(), child->value_size()));
            }
        }

#ifdef EWS_HAS_MAKE_UNIQUE
        return std::make_unique<weekly_recurrence>(
            std::move(interval), std::move(days_of_week),
            std::move(first_day_of_week));
#else
        return std::unique_ptr<weekly_recurrence>(
            new weekly_recurrence(std::move(interval), std::move(days_of_week),
                                  std::move(first_day_of_week)));
#endif
    }

    node = elem.first_node_ns(uri<>::microsoft::types(), "DailyRecurrence");
    if (node)
    {
        std::uint32_t interval = 0U;

        for (auto child = node->first_node(); child;
             child = child->next_sibling())
        {
            if (compare(child->local_name(), child->local_name_size(),
                        "Interval", std::strlen("Interval")))
            {
                interval = std::stoul(
                    std::string(child->value(), child->value_size()));
            }
        }

#ifdef EWS_HAS_MAKE_UNIQUE
        return std::make_unique<daily_recurrence>(std::move(interval));
#else
        return std::unique_ptr<daily_recurrence>(
            new daily_recurrence(std::move(interval)));
#endif
    }

    EWS_ASSERT(false &&
               "Expected one of "
               "<AbsoluteYearlyRecurrence>, <RelativeYearlyRecurrence>, "
               "<AbsoluteMonthlyRecurrence>, <RelativeMonthlyRecurrence>, "
               "<WeeklyRecurrence>, <DailyRecurrence>");
    return std::unique_ptr<recurrence_pattern>();
}

inline std::unique_ptr<recurrence_range>
recurrence_range::from_xml_element(const rapidxml::xml_node<>& elem)
{
    using rapidxml::internal::compare;

    EWS_ASSERT(std::string(elem.local_name()) == "Recurrence" &&
               "Expected a <Recurrence> element");

    auto node = elem.first_node_ns(internal::uri<>::microsoft::types(),
                                   "NoEndRecurrence");
    if (node)
    {
        date_time start_date;

        for (auto child = node->first_node(); child;
             child = child->next_sibling())
        {
            if (compare(child->local_name(), child->local_name_size(),
                        "StartDate", std::strlen("StartDate")))
            {
                start_date =
                    date_time(std::string(child->value(), child->value_size()));
            }
        }

#ifdef EWS_HAS_MAKE_UNIQUE
        return std::make_unique<no_end_recurrence_range>(std::move(start_date));
#else
        return std::unique_ptr<no_end_recurrence_range>(
            new no_end_recurrence_range(std::move(start_date)));
#endif
    }

    node = elem.first_node_ns(internal::uri<>::microsoft::types(),
                              "EndDateRecurrence");
    if (node)
    {
        date_time start_date;
        date_time end_date;

        for (auto child = node->first_node(); child;
             child = child->next_sibling())
        {
            if (compare(child->local_name(), child->local_name_size(),
                        "StartDate", std::strlen("StartDate")))
            {
                start_date =
                    date_time(std::string(child->value(), child->value_size()));
            }
            else if (compare(child->local_name(), child->local_name_size(),
                             "EndDate", std::strlen("EndDate")))
            {
                end_date =
                    date_time(std::string(child->value(), child->value_size()));
            }
        }

#ifdef EWS_HAS_MAKE_UNIQUE
        return std::make_unique<end_date_recurrence_range>(
            std::move(start_date), std::move(end_date));
#else
        return std::unique_ptr<end_date_recurrence_range>(
            new end_date_recurrence_range(std::move(start_date),
                                          std::move(end_date)));
#endif
    }

    node = elem.first_node_ns(internal::uri<>::microsoft::types(),
                              "NumberedRecurrence");
    if (node)
    {
        date_time start_date;
        std::uint32_t no_of_occurrences = 0U;

        for (auto child = node->first_node(); child;
             child = child->next_sibling())
        {
            if (compare(child->local_name(), child->local_name_size(),
                        "StartDate", std::strlen("StartDate")))
            {
                start_date =
                    date_time(std::string(child->value(), child->value_size()));
            }
            else if (compare(child->local_name(), child->local_name_size(),
                             "NumberOfOccurrences",
                             std::strlen("NumberOfOccurrences")))
            {
                no_of_occurrences = std::stoul(
                    std::string(child->value(), child->value_size()));
            }
        }

#ifdef EWS_HAS_MAKE_UNIQUE
        return std::make_unique<numbered_recurrence_range>(
            std::move(start_date), std::move(no_of_occurrences));
#else
        return std::unique_ptr<numbered_recurrence_range>(
            new numbered_recurrence_range(std::move(start_date),
                                          std::move(no_of_occurrences)));
#endif
    }

    EWS_ASSERT(false &&
               "Expected one of "
               "<NoEndRecurrence>, <EndDateRecurrence>, <NumberedRecurrence>");
    return std::unique_ptr<recurrence_range>();
}
}

// vim:et ts=4 sw=4
