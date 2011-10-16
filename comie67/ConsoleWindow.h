/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include "resource.h"
#include <atlhost.h>

#include "logger.h"

/*
 * CConsoleWindow is the widget that shows the results of CConsoleObject::Fire_PrintEvent().
 * Instead of being an integrate part of CExplorerBar, It's a seperate implementation
 * because I want to reuse CConsoleWindow somewhere else.
 */
class CConsoleWindow :
	public CDialogImpl<CConsoleWindow>
{
public:
	enum {IDD = IDD_CONSOLEWINDOW};

	CConsoleWindow() : mRichEditDll(NULL), mInitError(ERROR_SUCCESS) {}
	~CConsoleWindow() {}

BEGIN_MSG_MAP(CConsoleWindow)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
ALT_MSG_MAP(RICHEDIT_MSG_MAP_ID)
END_MSG_MAP()

	HWND GetRichEditWnd(void) {return mRichEditCtrl;}

private:
	CConsoleWindow(const CConsoleWindow&);
	CConsoleWindow& operator=(const CConsoleWindow&);

	enum {RICHEDIT_MSG_MAP_ID = 1};

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	HMODULE mRichEditDll;
	CContainedWindow mRichEditCtrl;

	DWORD mInitError;
};
