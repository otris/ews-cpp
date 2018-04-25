## Changelog
You'll find a complete list of changes at the project site on [GitHub](https://github.com/otris/ews-cpp).

### 0.8 (YYYY-MM-DD)

### 0.7 (2018-04-25)
New features:
- Support for the `<ResolveNames>` operation ([#73](https://github.com/otris/ews-cpp/issues/73)) has been added.
- Support for `<CreateFolder>`, `<GetFolder>`, `<DeleteFolder>`, `<FindFolder>` operations ([#96](https://github.com/otris/ews-cpp/pull/96)) has been added.
- Support for EWS notification operations (`<Subscribe>`, `<GetEvents>`, `<Unsubscribe>`) has been added ([#85](https://github.com/otris/ews-cpp/pull/85)).
- There is a new `date_time::to_epoch` and a `date_time::from_epoch` function to convert a xs:dateTime to the epoch back and forth.
- Support for the `<SyncFolderItems>` operation has been added ([#101](https://github.com/otris/ews-cpp/pull/101)). See the [example](https://github.com/otris/ews-cpp/blob/master/examples/sync_folder_items.cpp).
- Support for the `<GetRooms>` operation has been added ([#108](https://github.com/otris/ews-cpp/pull/108)).

And lots of fixes.

### 0.6 (2011-11-27)
New features:
- Support for EWS Delegate Access ([#2](https://github.com/otris/ews-cpp/issues/2)) and Impersonation ([#62](https://github.com/otris/ews-cpp/issues/61)) has been added.
- Allow retrieval of multiple items at once with `service::get_tasks()`, `get_messages()`, `get_contacts()`, and `get_calendar_items()` member-functions ([#55](https://github.com/otris/ews-cpp/issues/55) and [#69](https://github.com/otris/ews-cpp/issues/69)).
- The `service::create_item()` overload set has been extended to allow creating multiple items at once ([#64](https://github.com/otris/ews-cpp/issues/64)).
- `message` class gained some important properties.

Other:
- Doxygen generated API docs are automatically deployed to GitHub pages. This way, up-to-date API docs are always available at https://otris.github.io/ews-cpp/ ([#67](https://github.com/otris/ews-cpp/issues/67)).
- We now have a Changelog file ([#62](https://github.com/otris/ews-cpp/issues/62)).
