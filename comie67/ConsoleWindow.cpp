/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#include "stdafx.h"

#include "ConsoleWindow.h"

#include <richedit.h>

#include "dump.h"

LRESULT CConsoleWindow::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	Log(LOG_FUNC, _T("CConsoleWindow::OnInitDialog()\n"));

	RECT clientRect;
	GetParent().GetClientRect(&clientRect);
	MoveWindow(&clientRect);

	mRichEditDll = ::LoadLibrary(_T("riched20.dll"));
	if (mRichEditDll == NULL) {
		Log(LOG_ERROR, _T("LoadLibrary rich edit failed\n"));
		mInitError = ::GetLastError();
	}

	CWindow textArea;
	CWindow staticCtrl;
	staticCtrl.Attach(GetDlgItem(IDC_STATIC1));

	if (mInitError == ERROR_SUCCESS) {
		// Rich Edit Control is available.

		RECT replace;
		staticCtrl.GetWindowRect(&replace);

		CWindow parentOfStatic;
		parentOfStatic.Attach(staticCtrl.GetParent().m_hWnd);
		parentOfStatic.ScreenToClient(&replace);
		parentOfStatic.Detach();

		staticCtrl.ShowWindow(SW_HIDE);

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
	}

	staticCtrl.Detach();
	
	textArea.MoveWindow(&clientRect);
	textArea.Detach();

	return 1;
}

LRESULT CConsoleWindow::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	Log(LOG_FUNC, _T("CConsoleWindow::OnDestroy()\n"));

	return 0;
}
