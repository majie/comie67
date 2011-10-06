/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include <tchar.h>

typedef enum {
	LOG_FEW = 0,
	LOG_FUNC,
	LOG_MOREFUNC,
	LOG_ERROR,
	LOG_DUMP, // Dump data structures using macros in dump.h
	LOG_ALL
} LogLevel;

#ifdef _DEBUG

#define InitLog __InitLog
#define UninitLog __UninitLog
#define Log __Log

BOOL __InitLog(const TCHAR* filename);
VOID __UninitLog(void);
VOID __Log(LogLevel level, const TCHAR* format, ...);

#else //NDEBUG

#define InitLog
#define UninitLog void
#define Log

#endif //_DEBUG
