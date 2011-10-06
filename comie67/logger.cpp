/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

/* This file provides a simple printf-style log function to print
 * strings to a log file.
 * This can be implemented using ILog and IFileBasedLogInit if newer versions
 * of Visual Studio is in use.
 */

#include "stdafx.h"

#include <stdio.h>
#include <stdarg.h>
#include <tchar.h>

#include "logger.h"

#ifdef _DEBUG

// =====================================
static LogLevel gCurrentLevel = LOG_ALL;
// =====================================

static FILE* glogger;
static CRITICAL_SECTION gCrtSec;

BOOL __InitLog(const TCHAR* filename)
{
	glogger = _tfopen(filename, _T("w"));
	if (glogger == NULL) {
		return FALSE;
	}

	InitializeCriticalSection(&gCrtSec);

	return TRUE;
}

VOID __UninitLog()
{
	if (glogger != NULL) fclose(glogger);
	DeleteCriticalSection(&gCrtSec);
}

VOID __Log(LogLevel level, const TCHAR* format, ...)
{
	va_list vargs;

	if (glogger == NULL)
		return;

	if (level > gCurrentLevel)
		return;

	va_start(vargs, format);
	EnterCriticalSection(&gCrtSec);
	_vftprintf(glogger, format, vargs);
	LeaveCriticalSection(&gCrtSec);
	va_end(vargs);
}

#endif //_DEBUG
