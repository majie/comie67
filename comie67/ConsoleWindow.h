/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include "resource.h"
#include <atlhost.h>

#include "comie67.h"
#include "logger.h"

static enum {EVENT_SOURCE_ID = 1};

/*
 * CConsoleWindow is the widget that shows the results of CConsoleObject::Fire_PrintEvent().
 * Instead of being an integrate part of CExplorerBar, It's a seperate implementation
 * because I want to reuse CConsoleWindow somewhere else.
 */
class CConsoleWindow :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConsoleWindow, &CLSID_ConsoleWindow>,
	public CDialogImpl<CConsoleWindow>,
	public IDispEventSimpleImpl<EVENT_SOURCE_ID, CConsoleWindow, &DIID__IConsoleObjectEvents>,
	public IConsoleWindow
{
public:
	enum {IDD = IDD_CONSOLEWINDOW};

	CConsoleWindow() : mRichEditDll(NULL), mHasRichEdit(TRUE) {}
	virtual ~CConsoleWindow() {}

DECLARE_REGISTRY_RESOURCEID(IDR_CONSOLEWINDOW)

DECLARE_NOT_AGGREGATABLE(CConsoleWindow)

BEGIN_COM_MAP(CConsoleWindow)
	COM_INTERFACE_ENTRY(IConsoleWindow)
END_COM_MAP()

BEGIN_MSG_MAP(CConsoleWindow)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
ALT_MSG_MAP(RICHEDIT_MSG_MAP_ID)
END_MSG_MAP()

BEGIN_SINK_MAP(CConsoleWindow)
	SINK_ENTRY_INFO(EVENT_SOURCE_ID, DIID__IConsoleObjectEvents,
		COMIE67_PRINT_EVENT_ID, OnPrintEvent, &kPrintEventInfo)
	SINK_ENTRY_INFO(EVENT_SOURCE_ID, DIID__IConsoleObjectEvents,
		COMIE67_CLEAR_EVENT_ID, OnClearEvent, &kClearEventInfo)
END_SINK_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()
	STDMETHOD(FinalConstruct)(void);
	STDMETHOD_(void, FinalRelease)(void);

	STDMETHOD(Create)(HWND parent, LPARAM initParam, HWND* newWnd);
	STDMETHOD(ShowWindow)(int cmdShow);

private:
	CConsoleWindow(const CConsoleWindow&);
	CConsoleWindow& operator=(const CConsoleWindow&);

	enum {RICHEDIT_MSG_MAP_ID = 1};

	static _ATL_FUNC_INFO kPrintEventInfo;
	static _ATL_FUNC_INFO kClearEventInfo;

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);	

	STDMETHOD_(void, OnPrintEvent)(PrintLevel level, BSTR str);
	STDMETHOD_(void, OnClearEvent)(void){mRichEditCtrl.SetWindowText(_T(""));}

	HMODULE mRichEditDll;
	CContainedWindow mRichEditCtrl;
	CComQIPtr<IConsoleObject> mConsole;

	BOOL mHasRichEdit;
};

OBJECT_ENTRY_AUTO(__uuidof(ConsoleWindow), CConsoleWindow)
