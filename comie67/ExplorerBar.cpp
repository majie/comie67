/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#include "stdafx.h"

#include "ExplorerBar.h"

#include <oleidl.h>
#include <comutil.h>

#include "ConsoleWindow.h"
#include "logger.h"
#include "dump.h"

STDMETHODIMP CExplorerBar::SetSite(IUnknown* site)
{
	Log(LOG_FUNC, _T("CExplorerBar::SetSite(0x%08p)\n"), site);

	HRESULT hr = S_OK;
	if (site) {
		AddRef();

		hr = IObjectWithSiteImpl<CExplorerBar>::SetSite(site);
		if (FAILED(hr)) return hr;

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

		CComPtr<IConsoleWindow> consoleWindow;
		consoleWindow.CoCreateInstance(CLSID_ConsoleWindow, NULL, CLSCTX_INPROC_SERVER);

		hr = consoleWindow->Create(barWnd, 0, &m_hWndDlg);
		if (FAILED(hr)) {
			Log(LOG_ERROR, _T("window->Create(); failed\n"));
			return AtlHresultFromLastError();
		}

		consoleWindow->ShowWindow(SW_SHOW);
	} else {
		IObjectWithSiteImpl<CExplorerBar>::SetSite(site);
		Release();
	}

	return hr;
}
