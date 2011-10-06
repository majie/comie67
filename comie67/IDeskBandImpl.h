/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include <shlobj.h>
#include "logger.h"

template <typename T>
class ATL_NO_VTABLE IDeskBandImpl
	: public IDeskBand
{
public:

	// IOleWindow
	STDMETHOD(GetWindow)(HWND *phWnd);
	STDMETHOD(ContextSensitiveHelp)(BOOL) {return E_NOTIMPL;}

	// IDockingWindow
	STDMETHOD(ShowDW)(BOOL show);
	STDMETHOD(CloseDW)(DWORD reserved);
	STDMETHOD(ResizeBorderDW)(LPCRECT, IUnknown*, BOOL) {return E_NOTIMPL;}

	// IDeskBand
	STDMETHOD(GetBandInfo)(DWORD bandID, DWORD viewMode, DESKBANDINFO* info);
};

template <typename T>
STDMETHODIMP IDeskBandImpl<T>::GetWindow(HWND* wnd)
{
	Log(LOG_FUNC, _T("IDeskBandImpl<T>::GetWindow()\n"));

	if (!wnd) return E_POINTER;

	*wnd = static_cast<T*>(this)->m_hWndDlg;
	return S_OK;
}

template <typename T>
STDMETHODIMP IDeskBandImpl<T>::ShowDW(BOOL show)
{
	Log(LOG_FUNC, _T("IDeskBandImpl<T>::ShowDW(\"%s\")\n"), show?_T("true"):_T("false"));
	HWND wnd = static_cast<T*>(this)->m_hWndDlg;
	if (wnd)
		::ShowWindow(wnd, show?SW_SHOW:SW_HIDE);

	return S_OK;
}

template <typename T>
STDMETHODIMP IDeskBandImpl<T>::CloseDW(DWORD)
{
	Log(LOG_FUNC, _T("IDeskBandImpl<T>::CloseDW()\n"));
	ShowDW(FALSE);

	HWND* pWnd = &(static_cast<T*>(this)->m_hWndDlg);

	if (::IsWindow(*pWnd))
		::DestroyWindow(*pWnd);

	*pWnd = NULL;

	return S_OK;
}

template <typename T>
STDMETHODIMP IDeskBandImpl<T>::GetBandInfo(DWORD bandID, DWORD viewMode, DESKBANDINFO* info)
{
	Log(LOG_FUNC, _T("IDeskBandImpl<T>::GetBandInfo()\n"));

	if (!info)
		return E_INVALIDARG;

	if(info->dwMask & DBIM_MINSIZE)
	{
		info->ptMinSize.x = 30;
		info->ptMinSize.y = 30;
	}
	if(info->dwMask & DBIM_MAXSIZE)
	{
		info->ptMaxSize.x = -1;
		info->ptMaxSize.y = -1;
	}
	if(info->dwMask & DBIM_INTEGRAL)
	{
		info->ptIntegral.x = 1;
		info->ptIntegral.y = 1;
	}
	if(info->dwMask & DBIM_ACTUAL)
	{
		info->ptActual.x = 0;
		info->ptActual.y = 0;
	}
	if(info->dwMask & DBIM_TITLE)
	{
		lstrcpyW(info->wszTitle, L"comie67 output window");
	}
	if(info->dwMask & DBIM_MODEFLAGS)
	{
		info->dwModeFlags = DBIMF_VARIABLEHEIGHT;
	}
	if(info->dwMask & DBIM_BKCOLOR)
	{
		//Use the default background color by removing this flag.
		info->dwMask &= ~DBIM_BKCOLOR;
	}
	return S_OK;
}
