/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#include "stdafx.h"

#include "ConsoleWindow.h"

#include <comutil.h>
#include <richedit.h>

#include "dump.h"

_ATL_FUNC_INFO CConsoleWindow::kPrintEventInfo = {
	CC_STDCALL, VT_EMPTY, 2, {VT_I4, VT_BSTR}
};

_ATL_FUNC_INFO CConsoleWindow::kClearEventInfo = {
	CC_STDCALL, VT_EMPTY, 0
};

STDMETHODIMP CConsoleWindow::FinalConstruct()
{
	AddRef();

	// Get ConsoleObject
	CComPtr<IRunningObjectTable> rot;
	HRESULT hr = ::GetRunningObjectTable(0, &rot);
	if (FAILED(hr)) {
		Log(LOG_ERROR, _T("GetRunningObjectTable() failed\n"));
		return hr;
	}

	TCHAR name[64];
	_stprintf(name, _T("comie67:%u"), GetCurrentProcessId());

	CComPtr<IMoniker> moniker;
	hr = ::CreateItemMoniker(OLESTR("!"), name, &moniker);
	if (FAILED(hr)) {
		Log(LOG_ERROR, _T("CreateItemMoniker() failed\n"));
		return hr;
	}

	CComPtr<IUnknown> unknown;
	hr = rot->GetObject(moniker, &unknown);
	if (FAILED(hr)) {
		Log(LOG_ERROR, _T("rot->GetObject() failed\n"));
		return hr;
	}

	mConsole = unknown;
	if (mConsole == NULL) {
		Log(LOG_ERROR, _T("QI IConsoleObject\n"));
		return hr;
	}

	hr = DispEventAdvise(mConsole, &DIID__IConsoleObjectEvents);
	if (FAILED(hr)) {
		Log(LOG_ERROR, _T("DispEventAdvise() failed\n"));
		return hr;
	}

	mRichEditDll = ::LoadLibrary(_T("riched20.dll"));
	if (mRichEditDll == NULL) {
		Log(LOG_ERROR, _T("LoadLibrary rich edit failed\n"));
		hasRichEdit = FALSE;
	}

	return S_OK;
}

STDMETHODIMP_(void) CConsoleWindow::FinalRelease()
{
	DispEventUnadvise(mConsole, &DIID__IConsoleObjectEvents);

	::FreeLibrary(mRichEditDll);

	Release();
}

STDMETHODIMP CConsoleWindow::Create(HWND parent, LPARAM initParam, HWND* newWnd)
{
	HWND wnd = CDialogImpl<CConsoleWindow>::Create(parent, initParam);

	if (newWnd != NULL)
		*newWnd = wnd;

	return S_OK;
}

STDMETHODIMP CConsoleWindow::ShowWindow(int cmdShow)
{
	CDialogImpl<CConsoleWindow>::ShowWindow(cmdShow);

	return S_OK;
}

LRESULT CConsoleWindow::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	Log(LOG_FUNC, _T("CConsoleWindow::OnInitDialog()\n"));

	// Init GUI
	SetWindowText(_T("comie67.ConsoleWindow"));

	CWindow textArea;
	CWindow staticCtrl;
	staticCtrl.Attach(GetDlgItem(IDC_STATIC1));

	if (hasRichEdit) {
		// Rich Edit Control is available.

		RECT replace;
		staticCtrl.GetWindowRect(&replace);

		CWindow parentOfStatic;
		parentOfStatic.Attach(staticCtrl.GetParent().m_hWnd);
		parentOfStatic.ScreenToClient(&replace);
		parentOfStatic.Detach();

		staticCtrl.DestroyWindow();

		mRichEditCtrl.Create(RICHEDIT_CLASS, this, RICHEDIT_MSG_MAP_ID,
			this->m_hWnd, replace, _T("Rich Edit Control"),
			ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_NOHIDESEL |
			ES_READONLY | ES_NUMBER | WS_VSCROLL | WS_HSCROLL | WS_TABSTOP |
			WS_CHILD | WS_VISIBLE);

		HDC richEditDC = mRichEditCtrl.GetDC();
		int height = -MulDiv(9, GetDeviceCaps(richEditDC, LOGPIXELSY), 72);
		mRichEditCtrl.ReleaseDC(richEditDC);

		HFONT hFont = ::CreateFont(height, 0, 0, 0,	FW_NORMAL, FALSE, FALSE,
			FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
			ANTIALIASED_QUALITY, FIXED_PITCH, _T("Courier New"));

		Log(LOG_MOREFUNC, _T("CreateFont(): %x\n"), hFont);
		if (hFont != NULL)
			mRichEditCtrl.SetFont(hFont);
		::DeleteObject(hFont);

		mRichEditCtrl.SetWindowText(_T(""));

		textArea.Attach(mRichEditCtrl);
	} else {
		// Rich Edit Control is not available, show some error messages using the static text control.

		staticCtrl.SetWindowText(_T("riched20.dll load error"));
		textArea.Attach(staticCtrl);
		staticCtrl.Detach();
	}
	
	textArea.Detach();

	return 1;
}

LRESULT CConsoleWindow::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	Log(LOG_FUNC, _T("CConsoleWindow::OnDestroy()\n"));

	return 0;
}

LRESULT CConsoleWindow::OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled) {
	Log(LOG_FUNC, _T("OnSize: type: %d, width: %d, height: %d\n"), wParam, LOWORD(lParam), HIWORD(lParam));

	RECT rect = {0, 0, LOWORD(lParam), HIWORD(lParam)};
	if (mRichEditCtrl.IsWindow())
		mRichEditCtrl.MoveWindow(&rect, TRUE);
	return 0;
}

STDMETHODIMP_(void) CConsoleWindow::OnPrintEvent(PrintLevel level, BSTR str)
{
	Log(LOG_FUNC, _T("CConsoleWindow::OnPrintEvent\n"));

	_bstr_t bstr;

	// TODO: use small icons.
	switch(level) {
		case LEVEL_NORMAL:
			break;
		case LEVEL_DEBUG:
			break;
		case LEVEL_INFO:
			bstr = _T("[i]");
			break;
		case LEVEL_WARN:
			bstr = _T("[!]");
			break;
		case LEVEL_ERROR:
			bstr = _T("[X]");
			break;
		default:
			break;
	}

	bstr += str;
	bstr += _T("\n");

	// append text
	::SendMessage(mRichEditCtrl, EM_SETSEL, -1, -1);
	::SendMessage(mRichEditCtrl, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(static_cast<LPCTSTR>(bstr)));
}
