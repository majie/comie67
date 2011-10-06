/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#ifndef STRICT
#define STRICT
#endif

#ifndef WINVER
#define WINVER 0x0501 // Windows XP and Windows .NET Server
#endif

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0501	// Windows XP and Windows .NET Server
#endif						

#ifndef _WIN32_WINDOWS
#define _WIN32_WINDOWS 0x0490 // Windows Me
#endif

#ifndef _WIN32_IE
#define _WIN32_IE 0x0600 // Internet Explorer 6.0
#endif

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

// turns off ATL's hiding of some common and often safely ignored warning messages
#define _ATL_ALL_WARNINGS

//#define _ATL_DEBUG_INTERFACES

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <windows.h>

using namespace ATL;
