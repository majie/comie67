/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include "resource.h"

#include <exdisp.h>
#include <exdispid.h>
#include <atlcoll.h>
#include <atlstr.h>

#include "comie67.h"

#include "IDispatchExImpl.h"
#include "_IConsoleObjectEvents_CP.H"

static enum {EVENT_SOURCE_ID = 1};

/*
 * CConsoleObject is the very COM object to handle these kinds of calls
 * in Javascript: console.log('hello world');
 *
 * The Browser Helper Object part of CConsoleObject is in charge to automatically load
 * the ActiveX part to the Javascript context before a new web page is loaded.
 * While the ActiveX part interacts with Javascript scripts
 * and handles Javascript function calls.
 */
class ATL_NO_VTABLE CConsoleObject :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConsoleObject, &CLSID_ConsoleObject>,
	// ActiveX inheritances
	public IConnectionPointContainerImpl<CConsoleObject>,
	public CProxy_IConsoleObjectEvents<CConsoleObject>, 
	public IDispatchExImpl<IConsoleObject>,
	// BHO inheritances
	public IObjectWithSiteImpl<CConsoleObject>,
	public IDispEventSimpleImpl<EVENT_SOURCE_ID, CConsoleObject, &DIID_DWebBrowserEvents2>
{
public:
	CConsoleObject() {}
	~CConsoleObject() {}

DECLARE_REGISTRY_RESOURCEID(IDR_CONSOLEOBJECT)

DECLARE_NOT_AGGREGATABLE(CConsoleObject)

BEGIN_COM_MAP(CConsoleObject)
	COM_INTERFACE_ENTRY(IConsoleObject)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IDispatchEx)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
	COM_INTERFACE_ENTRY(IObjectWithSite)
END_COM_MAP()

BEGIN_CONNECTION_POINT_MAP(CConsoleObject)
	CONNECTION_POINT_ENTRY(__uuidof(_IConsoleObjectEvents))
END_CONNECTION_POINT_MAP()

BEGIN_SINK_MAP(CConsoleObject)
	SINK_ENTRY_INFO(EVENT_SOURCE_ID, DIID_DWebBrowserEvents2, DISPID_NAVIGATECOMPLETE2,
		OnNavigateComplete2, &kNavigateComplete2Info)
END_SINK_MAP()

	// CComObjectRootEx
	DECLARE_PROTECT_FINAL_CONSTRUCT()
	STDMETHOD(FinalConstruct)(void);
	STDMETHOD_(void, FinalRelease)(void);

	// IObjectWithSite
	STDMETHOD(SetSite)(IUnknown* site);

	// IConsoleObject
	STDMETHOD(log)(SAFEARRAY* varg);
	STDMETHOD(debug)(SAFEARRAY* varg);
	STDMETHOD(info)(SAFEARRAY* varg);
	STDMETHOD(warn)(SAFEARRAY* varg);
	STDMETHOD(error)(SAFEARRAY* varg);
#pragma push_macro("assert")
#undef assert
	STDMETHOD(assert)(VARIANT_BOOL expr, SAFEARRAY* varg);
#pragma pop_macro("assert")
	STDMETHOD(clear)(void);
	STDMETHOD(trace)(void);
	STDMETHOD(time)(BSTR name);
	STDMETHOD(timeEnd)(BSTR name);
	STDMETHOD(profile)(BSTR title);
	STDMETHOD(profileEnd)(void);
	STDMETHOD(count)(BSTR title);

private:
	CConsoleObject(const CConsoleObject&);
	CConsoleObject& operator=(const CConsoleObject&);

	static _ATL_FUNC_INFO kNavigateComplete2Info;

	// WebBrowser2 event
	STDMETHOD_(void, OnNavigateComplete2)(IDispatch* dispatch, VARIANT* url);

	// simulate sprintf(char* buf, const char* format, ...)
	STDMETHOD(BstrPrintf)(/*out*/BSTR* result, /*in*/const SAFEARRAY& varg) const;

	DWORD mRotCookie;

	unsigned long	mRunningTimer;
	UINT			mMinResolution;

	CAtlMap<CAtlString, DWORD>			mTimerMap;
	CAtlMap<CAtlString, unsigned long>	mCountMap;
};

OBJECT_ENTRY_AUTO(__uuidof(ConsoleObject), CConsoleObject)
