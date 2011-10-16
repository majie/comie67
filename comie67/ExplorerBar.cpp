/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#include "stdafx.h"

#include "ExplorerBar.h"

#include <new>

#include <oleidl.h>
#include <comutil.h>

#include "ConsoleWindow.h"
#include "logger.h"
#include "dump.h"

_ATL_FUNC_INFO CExplorerBar::kPrintEventInfo = {
	CC_STDCALL, VT_EMPTY, 2, {VT_I4, VT_BSTR}
};

_ATL_FUNC_INFO CExplorerBar::kClearEventInfo = {
	CC_STDCALL, VT_EMPTY, 0
};

STDMETHODIMP CExplorerBar::SetSite(IUnknown* site)
{
	Log(LOG_FUNC, _T("CExplorerBar::SetSite(0x%08p)\n"), site);

	HRESULT hr = S_OK;
	if (site) {
		hr = IObjectWithSiteImpl<CExplorerBar>::SetSite(site);
		if (FAILED(hr)) return hr;

		AddRef();

		CComPtr<IRunningObjectTable> rot;
		hr = ::GetRunningObjectTable(0, &rot);
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

		CComQIPtr<IConsoleObject> console = unknown;
		if (console == NULL) {
			Log(LOG_ERROR, _T("QI IConsoleObject\n"));
			return hr;
		}

		hr = DispEventAdvise(console, &DIID__IConsoleObjectEvents);
		if (FAILED(hr)) {
			Log(LOG_ERROR, _T("DispEventAdvise() failed\n"));
			return hr;
		}

		CComQIPtr<IOleWindow> ieWindow = m_spUnkSite;
		if (!ieWindow) {
			Log(LOG_ERROR, _T("QI IOleWindow failed\n"));
			return E_NOINTERFACE;
		}

		HWND barWnd;
		hr = ieWindow->GetWindow(&barWnd);
		if (FAILED(hr)) {
			Log(LOG_ERROR, _T("IOleWindow::GetWindow() failed\n"));
			return hr;
		}
		Log(LOG_MOREFUNC, _T("bar window: 0x%X\n"), barWnd);

		CWindow bar;
		bar.Attach(barWnd);
		bar.SetWindowText(_T("comie67.ExplorerBar"));
		bar.Detach();

		CConsoleWindow* window = new(std::nothrow) CConsoleWindow();
		if (window == NULL) {
			return E_OUTOFMEMORY;
		}

		m_hWndDlg = window->Create(barWnd, 0);
		if (m_hWndDlg == NULL) {
			Log(LOG_ERROR, _T("window->Create(); failed\n"));
			return HRESULT_FROM_WIN32(::GetLastError());
		}

		mRichEdit.Attach(window->GetRichEditWnd());

		window->ShowWindow(SW_SHOW);
	} else {
		CComQIPtr<IConsoleObject> console = m_spUnkSite;
		if (console) {
			DispEventUnadvise(console, &DIID_DWebBrowserEvents2);
		}

		mRichEdit.Detach();

		Release();
		IObjectWithSiteImpl<CExplorerBar>::SetSite(site);
	}

	return hr;
}

STDMETHODIMP_(void) CExplorerBar::OnPrintEvent(PrintLevel level, BSTR str)
{
	Log(LOG_FUNC, _T("CExplorerBar::OnPrintEvent\n"));

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
	::SendMessage(mRichEdit, EM_SETSEL, -1, -1);
	::SendMessage(mRichEdit, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(static_cast<LPCTSTR>(bstr)));
}
