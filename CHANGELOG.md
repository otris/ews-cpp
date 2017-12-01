## Changelog

You'll find a complete list of changes at the project site on [GitHub](https://github.com/otris/ews-cpp).

### 0.7 (YYYY-MM-DD)

New features:
- Support for the <ResolveNames> operation ([#73](https://github.com/otris/ews-cpp/issues/73)) has been added.

### 0.6 (2011-11-27)

New features:
- Support for EWS Delegate Access ([#2](https://github.com/otris/ews-cpp/issues/2)) and Impersonation ([#62](https://github.com/otris/ews-cpp/issues/61)) has been added.
- Allow retrieval of multiple items at once with `service::get_tasks()`, `get_messages()`, `get_contacts()`, and `get_calendar_items()` member-functions ([#55](https://github.com/otris/ews-cpp/issues/55) and [#69](https://github.com/otris/ews-cpp/issues/69)).
- The `service::create_item()` overload set has been extended to allow creating multiple items at once ([#64](https://github.com/otris/ews-cpp/issues/64)).
- `message` class gained some important properties.

Other:
- Doxygen generated API docs are automatically deployed to GitHub pages. This way, up-to-date API docs are always available at https://otris.github.io/ews-cpp/ ([#67](https://github.com/otris/ews-cpp/issues/67)).
- We now have a Changelog file ([#62](https://github.com/otris/ews-cpp/issues/62)).
