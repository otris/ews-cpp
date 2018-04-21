[![Build Status](https://travis-ci.org/otris/ews-cpp.svg?branch=master)](https://travis-ci.org/otris/ews-cpp)
[![Build Status](https://ci.appveyor.com/api/projects/status/github/otris/ews-cpp?svg=true)](https://ci.appveyor.com/project/BenjaminKircher/ews-cpp)

# ews-cpp
A C++11 header-only library for Microsoft Exchange Web Services.

<!-- TOC -->

- [Example](#example)
- [Documentation](#documentation)
- [Overview](#overview)
    - [Supported Operations and Elements](#supported-operations-and-elements)
    - [Supported Authentication Schemes](#supported-authentication-schemes)
    - [Supported Compilers](#supported-compilers)
    - [Supported Operating Systems](#supported-operating-systems)
    - [Supported Microsoft Exchange Server Versions](#supported-microsoft-exchange-server-versions)
    - [Run-time Dependencies](#run-time-dependencies)
    - [Dev Dependencies](#dev-dependencies)
    - [Note Windows Users](#note-windows-users)
    - [Source Code](#source-code)
    - [Building](#building)
        - [Linux](#linux)
        - [Windows](#windows)
        - [API Docs](#api-docs)
        - [Test Suite](#test-suite)
    - [Design Notes](#design-notes)
    - [API](#api)
- [More EWS Resources](#more-ews-resources)
- [Legal Notice](#legal-notice)

<!-- /TOC -->

# Example
You can find a couple of small executable examples in the
[examples/](https://github.com/otris/ews-cpp/tree/master/examples) folder.

As an appetizer, this is how it looks when you create and send an email message
with ews-cpp:

```cpp
#include <ews/ews.hpp>

#include <exception>
#include <iostream>
#include <ostream>
#include <string>
#include <vector>

int main() {
    ews::set_up();

    try {
        auto service = ews::service("https://example.com/ews/Exchange.asmx",
                                    "ACME",
                                    "myuser",
                                    "mysecret");
        auto message = ews::message();
        message.set_subject("Test mail from outer space");
        std::vector<ews::mailbox> recipients{ ews::mailbox("president@example.com") };
        message.set_to_recipients(recipients);
        auto text = ews::body("ようこそ (Welcome!)\n\nThis is a test.\n");
        message.set_body(text);
        service.create_item(message, ews::message_disposition::send_and_save_copy);
    } catch (std::exception& exc) {
        std::cout << exc.what() << std::endl;
    }

    ews::tear_down();
    return 0;
}
```

# Documentation
We host automatically generated API documentation here:
[otris.github.io/ews-cpp](https://otris.github.io/ews-cpp/).

# Overview
EWS is an API that third-party programmers can use to communicate with
Microsoft Exchange Server. The API exists since Exchange Server 2007 and is
continuously up-dated by Microsoft and present in the latest iteration of the
product, Exchange Server 2016.

This library provides a native and platform-independent way to use EWS in your
C++ application.

<img src="ews-overview.png" width=80%>

## Supported Operations and Elements

* Most items are supported: `<CalendarItem>`, `<Message>`, `<Contact>`,
  `<Task>`. Note that we still don't support all properties of all of these
  items. But we're working on it.
* We support `<CreateItem>`, `<SendItem>`, `<FindItem>`, `<GetItem>`,
  `<UpdateItem>`, and `<DeleteItem>` operations.
* `<CreateFolder>`, `<GetFolder>`, `<DeleteFolder>`, `<FindFolder>` operations
  on folders.
* Basic support for attachments: `<CreateAttachment>`, `<GetAttachment>`,
  `<DeleteAttachment>` for file attachments. Note: Item attachments are not
  supported yet.
* Allow someone else to manage your mail and calendar:
  - Impersonation. When you have a service process that needs to do regular
    tasks for a group of mailboxes or every mailbox in a database.
  - Delegate access. For granting access to individual  users. You can add
    permissions individually to each mailbox and directly modify folder
    permissions.
* Discovering Exchange servers in your network with Autodiscover is supported.
* EWS notification operations (`<Subscribe>`, `<GetEvents>`, `<Unsubscribe>`)
  to inform you about specific item or folder changes are supported.

## Supported Authentication Schemes

* HTTP basic auth
* NTLM

Note: Kerberos is currently not supported but its on the TODO list.

## Supported Compilers

* Microsoft Visual Studio since VS 2012
* Clang since 3.5
    - with libc++ on Mac OS X
    - with libstdc++ on Linux (Note that libc++ on Linux is not supported)
* GCC since 4.8 with libstdc++

## Supported Operating Systems

* Microsoft Windows 10
* macOS starting with 10.12
* RHEL 7
* Ubuntu starting with 16.04 LTS

## Supported Microsoft Exchange Server Versions

* Microsoft Exchange Server 2013 SP1

However, our goal is to support all Exchange Server versions since 2007.

## Run-time Dependencies
The only thing you need for ews-cpp to run is [libcurl](https://curl.haxx.se/).

## Dev Dependencies
If you want to hack on ews-cpp itself you additionally need

* Git
* CMake
* Doxygen (optional)
* Python 2 or 3 (optional)

## Note Windows Users
You can obtain an up-to-date and easy-to-use binary distribution of libcurl
from here: [confusedbycode.com/curl](http://www.confusedbycode.com/curl/)

Additionally, you probably need to tell CMake where to find it. Just set
`CMAKE_PREFIX_PATH` to the path where you installed libcurl (e.g.
``C:\Program Files\cURL``) and re-configure.

You can also use the Windows batch script provided in
``scripts\build-curl.bat`` to download and compile libcurl for your particular
version of Visual Studio.

## Source Code
ews-cpp's source code is available as a Git repository. To obtain it, type:

```bash
git clone --recursive https://github.com/otris/ews-cpp.git
```

## Building
### Linux
The library is header-only. So there is no need to build anything. Just copy the
`include/ews/` directory wherever you may like.

To build the accompanied tests with debugging symbols and Address Sanitizer
enabled do something like this:

```bash
cmake -DCMAKE_BUILD_TYPE=Debug -DENABLE_SANITIZERS=ON /path/to/source
make
```

Type `make edit_cache` to see all configuration options. `make help` shows you
all available targets.

### Windows
To build the tests and examples on Windows you can use `cmake-gui`.  For more
see: https://cmake.org/runningcmake/

If you do not want to use any GUI to compile the examples and tests you could
do something like this with the Windows `cmd.exe` command prompt:

```bat
set PATH=%PATH%;C:\Program Files (x86)\CMake\bin
mkdir builddir
cd builddir
cmake -G "Visual Studio 14 2015 Win64" ^
    -DCURL_LIBRARY="C:\Program Files\cURL\7.49.1\win64-debug\lib\libcurl_debug.lib" ^
    -DCURL_INCLUDE_DIR="C:\Program Files\cURL\7.49.1\win64-debug\include" ^
    C:\path\to\source
cmake --build .
```

Make sure to choose the right generator for your environment.

### API Docs
Use the `doc` target to create the API documentation with Doxygen.  Type:

```bash
make doc
open html/index.html
```

### Test Suite
In order to run individual examples or the test suite export following
environment variables like this:

```bash
export EWS_TEST_DOMAIN="DUCKBURG"
export EWS_TEST_USERNAME="mickey"
export EWS_TEST_PASSWORD="pluto"
export EWS_TEST_URI"https://hire-a-detective.com/ews/Exchange.asmx"
```

Be sure to change the values to an actual account on some Exchange server that
you can use for testing. Do not run the tests on your own production account.

Once you've build the project, you can execute the tests with:

```bash
./tests --assets=/path/to/source/tests/assets
```

## Design Notes
ews-cpp is written in a "modern C++" way:

* C++ Standard Library, augmented with rapidxml for XML parsing
* Smart pointers instead of raw pointers
* Pervasive RAII idiom
* Handle errors using C++ exceptions
* Coding conventions inspired by Boost

## API
Just add:

```cpp
#include <ews/ews.hpp>
```

to your include directives and you are good to go.

Take a look at the `examples/` directory to get an idea of how the API feels.
ews-cpp is a thin wrapper around Microsoft's EWS API. You will need to refer to
the [EWS reference for
Exchange](https://msdn.microsoft.com/en-us/library/office/bb204119%28v=exchg.150%29.aspx)
for all available parameters to pass and all available attributes on items.
From 10.000ft it looks like this:

<img src="ews-objects.png" width=80%>

You have items and you have **the** service. You use the service whenever you
want to talk to the Exchange server.

Please note one important caveat though. ews-cpp's API is designed to be
"blocking". This means whenever you call one of the service's member functions
to talk to an Exchange server that call blocks until it receives a request from
the server. And that may, well, just take forever (actually until a timeout is
reached). You need to keep this in mind in order to not block your main or UI
thread.

Implications of this design choice

Pros:

* A blocking API is much easier to use and understand

Cons:

* You just might accidentally block your UI thread
* You cannot issue thousands of EWS requests asynchronously simply because you
  cannot spawn thousands of threads in your process. You may need additional
  effort here

# More EWS Resources
* [EWS Editor](http://ewseditor.codeplex.com/) is an excellent Open Source tool to test EWS Managed API and do raw SOAP POSTs. Sources are available here: [github.com/dseph/EwsEditor](https://github.com/dseph/EwsEditor).
* [This article](https://blogs.msdn.microsoft.com/webdav_101/2015/05/11/best-practices-ews-authentication-and-access-issues/) on blogs.msdn.microsoft.com describing best practices in EWS Authentication and solving access issues.

# Legal Notice
ews-cpp is developed by otris software AG and was initially released to the
public in July 2016. It is licensed under the Apache License, Version 2.0 (see
[LICENSE file](https://github.com/otris/ews-cpp/blob/master/LICENSE)).

For more information about otris software AG visit our website
[otris.de](https://www.otris.de/) or our Open Source repositories at
[github.com/otris](https://github.com/otris).


<!-- vim: et sw=4 ts=4:
-->
