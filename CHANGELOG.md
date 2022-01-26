## Changelog
You'll find a complete list of changes at the project site on [GitHub](https://github.com/otris/ews-cpp).

### 0.10 (2022-01-26)
New features:
- Add class name functions to recurrence_pattern ([#185](https://github.com/otris/ews-cpp/pull/185), [#186](https://github.com/otris/ews-cpp/pull/186))
- Credentials clone ([#192](https://github.com/otris/ews-cpp/pull/192)
- Allow to set the Expect header ([#194](https://github.com/otris/ews-cpp/pull/194))
- Methods for setting server URI and credentials ([#195](https://github.com/otris/ews-cpp/pull/195))
- Set Expect header for BASIC auth ([#196](https://github.com/otris/ews-cpp/pull/196))
- Constructor with `caPath` arg ([#197](https://github.com/otris/ews-cpp/pull/197))

Several fixes.

### 0.9 (2021-03-15)
New features:
- OAuth2 support ([#148](https://github.com/otris/ews-cpp/issues/148)
- Extended settings for connecting to EWS via a proxy ([#172](https://github.com/otris/ews-cpp/issues/172))
- Debug callback for EWS requests ([#175](https://github.com/otris/ews-cpp/issues/175))
- Support Contact information in Resolve Names resolutions ([#165](https://github.com/otris/ews-cpp/issues/165))

Several fixes.

### 0.8 (2020-01-02)
New features:
- Support updating extended properties ([#149](https://github.com/otris/ews-cpp/issues/142))
- Updating physical address of contact  ([#141](https://github.com/otris/ews-cpp/issues/141))
- Inline attachments with content ID  ([#144](https://github.com/otris/ews-cpp/issues/144))
- Implementation of UpdateFolder operation ([#128](https://github.com/otris/ews-cpp/issues/128))

Several fixes.

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
