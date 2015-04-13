#pragma once

#include <exception>
#include <stdexcept>
#include <string>
#include <vector>
#include <unordered_map>
#include <sstream>
#include <algorithm>
#include <functional>
#include <utility>
#include <memory>
#ifndef _MSC_VER
#include <type_traits>
#endif
#include <cstddef>
#include <cstring>

#include <curl/curl.h>

#include "rapidxml/rapidxml.hpp"

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

// Visual Studio 12 does not support noexcept specifier.
#ifndef _MSC_VER
# define EWS_NOEXCEPT noexcept
#else
# define EWS_NOEXCEPT 
#endif

namespace ews
{
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
        // smtp address. You must supply one of these identifiers and of course,
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
    inline response_code
    string_to_response_code_enum(const char* str)
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
            throw exception("Unrecognized response code");
        }
        return it->second;
    }

    enum class base_shape { id_only, default_shape, all_properties };

    // TODO: move to internal namespace
    inline std::string base_shape_str(base_shape shape)
    {
        switch (shape)
        {
            case base_shape::id_only: return "IdOnly";
            case base_shape::default_shape: return "Default";
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

    // Exception thrown when a request to an Exchange server was not successful
    class exchange_error final : public exception
    {
    public:
        explicit exchange_error(response_code code)
            : exception("Request failed"),
              code_(code)
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
        inline curl_error make_curl_error(const std::string& msg, CURLcode rescode)
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

#ifndef _MSC_VER
        // <type_traits> is broken in Visual Studio 12, disregard
        static_assert(!std::is_copy_constructible<on_scope_exit>::value, "");
        static_assert(!std::is_copy_assignable<on_scope_exit>::value, "");
        static_assert(!std::is_default_constructible<on_scope_exit>::value, "");
#endif

        // Raised when a response from a server could not be parsed.
        class parse_error final : public exception
        {
        public:
            explicit parse_error(const std::string& what)
                : exception(what)
            {
            }
            explicit parse_error(const char* what) : exception(what) {}
        };

        // Forward declarations
        class http_request;

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
                : data_(std::move(data)), doc_(), code_(code), parsed_(false)
            {
                EWS_ASSERT(!data_.empty());
            }

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
                if (!parsed_)
                {
                    on_scope_exit set_parsed([&]()
                    {
                        parsed_ = true;
                    });
                    parse();
                }
                return doc_;
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

            // Returns whether the HTTP respone code is 200 (OK).
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
                    doc_.parse<flags>(&data_[0]);
                }
                catch (rapidxml::parse_error& e)
                {
                    // Swallow and erase type
                    throw parse_error(e.what());
                }
            }

            std::vector<char> data_;
            rapidxml::xml_document<char> doc_;
            long code_;
            bool parsed_;
        };

        // Traverse elements, depth first, beginning with given node.
        //
        // Applies given function to every element during traversal, stopping as
        // soon as that function returns true.
        template <typename Function>
        inline void traverse_elements(const rapidxml::xml_node<>& node,
                                      Function func)
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
                             const char* namespace_uri)
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
                    "The request failed for unknown reason (no XML in response)");
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

#ifndef NDEBUG
# ifdef EWS_ENABLE_VERBOSE
                // Print HTTP headers to stderr
                set_option(CURLOPT_VERBOSE, 1L);
# endif
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
        inline http_response make_raw_soap_request(
            const std::string& url, const std::string& username,
            const std::string& password, const std::string& domain,
            const std::string& soap_body,
            const std::vector<std::string>& soap_headers)
        {
            http_request request{url};
            request.set_method(http_request::method::POST);
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
                code =
                    string_to_response_code_enum(response_code_elem->value());
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

    } // internal

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

#ifndef _MSC_VER
    static_assert(std::is_default_constructible<item_id>::value, "");
    static_assert(std::is_copy_constructible<item_id>::value, "");
    static_assert(std::is_copy_assignable<item_id>::value, "");
    static_assert(std::is_move_constructible<item_id>::value, "");
    static_assert(std::is_move_assignable<item_id>::value, "");
#endif

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

    // A date/time string wrapper class for xs:dateTime formatted strings.
    //
    // See Note About Dates in EWS above.
    class date_time final
    {
    public:
        date_time(std::string str) // intentionally not explicit
            : date_time_string_(std::move(str))
        {
        }

    private:
        std::string date_time_string_;
    };

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

    // Represents the actual body content of a message.
    //
    // This can be of type Best, HTML, or plain-text. See EWS XML elements
    // documentation on MSDN.
    class body final
    {
    public:
        body(const std::string& text) // intentionally not explicit
        {
            (void)text;
        }
    };

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
        explicit item(item_id id) : item_id_(std::move(id)) {}

        const item_id& get_item_id() const EWS_NOEXCEPT { return item_id_; }

        void set_subject(const std::string& str) { subject_ = str; }
        const std::string& get_subject() const EWS_NOEXCEPT { return subject_; }

    private:
        item_id item_id_;
        std::string subject_;

        friend class service;
        // Sub-classes reimplement functions below
        std::string create_item_request_string() const
        {
            EWS_ASSERT(false && "TODO");
            return std::string();
        }
    };

#ifndef _MSC_VER
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
        explicit task(item_id id) : item(id) {}

        void set_body(const body&) {} // TODO: getter

        void set_owner(const std::string&) {}
        void set_start_date(const ews::date_time&) {}
        void set_due_date(const ews::date_time&) {}
        void set_reminder_enabled(bool) {}
        void set_reminder_due_by(const ews::date_time&) {}

        // Makes a task instance from a <Task> XML element
        static task from_xml_element(const rapidxml::xml_node<>& elem)
        {
            using uri = internal::uri<>;

            auto node = elem.first_node_ns(uri::microsoft::types(), "ItemId");
            EWS_ASSERT(node && "Expected <ItemId>");
            auto t = task(item_id::from_xml_element(*node));
            node = elem.first_node_ns(uri::microsoft::types(), "Subject");
            if (node)
            {
                t.set_subject(std::string(node->value(), node->value_size()));
            }
            return t;
        }

    private:
        friend class service;
        std::string create_item_request_string() const
        {
            return
              "<CreateItem " \
                  "xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " \
                  "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" >" \
              "<Items>" \
              "<t:Task>" \
              "<t:Subject>" + get_subject() + "</t:Subject>" \
              "</t:Task>" \
              "</Items>" \
              "</CreateItem>";
        }
    };

#ifndef _MSC_VER
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

        void set_given_name(const std::string& str) { given_name_ = str; }
        const std::string& get_given_name() const EWS_NOEXCEPT
        {
            return given_name_;
        }
        void set_surname(const std::string& str) { surname_ = str; }
        const std::string& get_surname() const EWS_NOEXCEPT { return surname_; }

        // Makes a contact instance from a <Contact> XML element
        static contact from_xml_element(const rapidxml::xml_node<>& elem)
        {
            using uri = internal::uri<>;

            auto node = elem.first_node_ns(uri::microsoft::types(), "ItemId");
            EWS_ASSERT(node && "Expected <ItemId>");
            auto c = contact(item_id::from_xml_element(*node));
            node = elem.first_node_ns(uri::microsoft::types(), "Subject");
            if (node)
            {
                c.set_subject(std::string(node->value(), node->value_size()));
            }
            return c;
        }

    private:
        std::string given_name_;
        std::string surname_;

        friend class service;
        std::string create_item_request_string() const
        {
            return
              "<CreateItem " \
                  "xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " \
                  "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" >" \
              "<Items>" \
              "<t:Contact>" \
              "<t:Subject>" + get_subject() + "</t:Subject>" \
              "</t:Contact>" \
              "</Items>" \
              "</CreateItem>";
        }
    };

#ifndef _MSC_VER
    static_assert(std::is_default_constructible<contact>::value, "");
    static_assert(std::is_copy_constructible<contact>::value, "");
    static_assert(std::is_copy_assignable<contact>::value, "");
    static_assert(std::is_move_constructible<contact>::value, "");
    static_assert(std::is_move_assignable<contact>::value, "");
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
    class service final
    {
    public:
        // FIXME: credentials are stored plain-text in memory
        //
        // That'll be bad. We wouldn't want random Joe at first-level support to
        // see plain-text passwords and user-names just because the process
        // crashed and some automatic mechanism sent a minidump over the wire.
        // What are our options? Security-by-obscurity: we could hash
        // credentials with a hash of the process-id or something.
        service(std::string server_uri, std::string domain,
                std::string username, std::string password)
            : server_uri_(std::move(server_uri)),
              domain_(std::move(domain)),
              username_(std::move(username)),
              password_(std::move(password)),
              server_version_("Exchange2013_SP1")
        {
        }

        // Gets a task from the Exchange store
        task get_task(item_id id)
        {
            return get_item<task>(id, base_shape::all_properties);
        }

        // Gets a contact from the Exchange store
        contact get_contact(item_id id)
        {
            return get_item<contact>(id, base_shape::all_properties);
        }

        // Delete a task item from the Exchange store
        void delete_task(task&& the_task,
                         delete_type del_type,
                         affected_task_occurrences affected)
        {
            using internal::delete_item_response_message;

            const std::string request_string =
              "<DeleteItem " \
                  "xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " \
                  "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" " \
                  "DeleteType=\"" + delete_type_str(del_type) + "\" " \
                  "AffectedTaskOccurrences=\"" + affected_task_occurrences_str(affected) + "\">" \
              "<ItemIds>" + the_task.get_item_id().to_xml("t") + "</ItemIds>" \
              "</DeleteItem>";
            auto response = request(request_string);
            const auto response_message =
                delete_item_response_message::parse(response);
            if (!response_message.success())
            {
                throw exchange_error(response_message.get_response_code());
            }
            the_task = task();
        }

        // Delete a contact item from the Exchange store
        void delete_contact(contact&& the_contact)
        {
            using internal::delete_item_response_message;

            const std::string request_string =
              "<DeleteItem " \
                  "xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " \
                  "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">" \
              "<ItemIds>" + the_contact.get_item_id().to_xml("t") + "</ItemIds>" \
              "</DeleteItem>";
            auto response = request(request_string);
            const auto response_message =
                delete_item_response_message::parse(response);
            if (!response_message.success())
            {
                throw exchange_error(response_message.get_response_code());
            }
            the_contact = contact();
        }

        // Only purpose of this overload-set is to prevent exploding template
        // code in errors messages in caller's code
        item_id create_item(const task& the_item)
        {
            return create_item<task>(the_item);
        }

        item_id create_item(const contact& the_item)
        {
            return create_item<contact>(the_item);
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
            // TODO: support multiple dialects depending on server version
            const auto soap_headers = std::vector<std::string> {
                "<t:RequestServerVersion Version=\"Exchange2013_SP1\"/>"
            };
            auto response = internal::make_raw_soap_request(server_uri_,
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
        ItemType get_item(item_id id, base_shape shape)
        {
            using get_item_response_message =
                internal::get_item_response_message<ItemType>;

            const std::string request_string =
              "<GetItem " \
                  "xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\" " \
                  "xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" >" \
              "<ItemShape>" \
              "<t:BaseShape>" + base_shape_str(shape) + "</t:BaseShape>" \
              "</ItemShape>" \
              "<ItemIds>" + id.to_xml("t") + "</ItemIds>" \
              "</GetItem>";
            auto response = request(request_string);
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
        item_id create_item(const ItemType& the_item)
        {
            using internal::create_item_response_message;

            auto response = request(the_item.create_item_request_string());
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
    };

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
