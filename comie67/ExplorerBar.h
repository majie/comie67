/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include "resource.h"

#include <atlwin.h>

#include "comie67.h"
#include "IDeskBandImpl.h"

/*
 * CExplorerBar registers itself as a horizontal explorer bar. It provides a space
 * for a child window (e.g. CConsoleWindow) to fill up.
 *
 * Any new widgets should be added to CConsoleWindow instead of CExplorerBar.
 */
class ATL_NO_VTABLE CExplorerBar : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CExplorerBar, &CLSID_ExplorerBar>,
	public IObjectWithSiteImpl<CExplorerBar>,
	public IPersistStreamInitImpl<CExplorerBar>,
	public IDeskBandImpl<CExplorerBar>,
	public IExplorerBar
{
public:
	CExplorerBar() : m_bRequiresSave(FALSE), m_hWndDlg(NULL)
	{
	}
	virtual ~CExplorerBar() {}

	// IPersistStreamInitImpl
	BOOL m_bRequiresSave;

	// IDeskBandImpl
	HWND m_hWndDlg;

DECLARE_REGISTRY_RESOURCEID(IDR_EXPLORERBAR)

DECLARE_NOT_AGGREGATABLE(CExplorerBar)

BEGIN_COM_MAP(CExplorerBar)
	COM_INTERFACE_ENTRY(IObjectWithSite)
	COM_INTERFACE_ENTRY_IID(IID_IPersist, IPersistStreamInit)
	COM_INTERFACE_ENTRY_IID(IID_IPersistStream, IPersistStreamInit)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY(IDeskBand)
	COM_INTERFACE_ENTRY(IExplorerBar)
END_COM_MAP()

BEGIN_PROP_MAP(CExplorerBar)
END_PROP_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()
	HRESULT FinalConstruct(){return S_OK;}
	void FinalRelease(){}

	// IObjectWithSite
	STDMETHOD(SetSite)(IUnknown* site);	

private:
	CExplorerBar(const CExplorerBar&);
	CExplorerBar& operator=(const CExplorerBar&);
};

OBJECT_ENTRY_AUTO(__uuidof(ExplorerBar), CExplorerBar)
