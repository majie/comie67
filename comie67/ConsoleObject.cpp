/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#include "stdafx.h"

#include "ConsoleObject.h"

#include <stdarg.h>

#include <atlsafe.h>
#include <atlstr.h>
#include <atlcomcli.h>

#include <mshtml.h>
#include <exdisp.h>
#include <exdispid.h>

#include "CComPtrIDispatchEx.h"
#include "logger.h"
#include "dump.h"


_ATL_FUNC_INFO CConsoleObject::kNavigateComplete2Info = {
	CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH, VT_VARIANT | VT_BYREF}
};

STDMETHODIMP CConsoleObject::FinalConstruct()
{
	HRESULT hr = InitTypeInfo();
	if (FAILED(hr))
		return hr;

	TIMECAPS timeCaps;
	MMRESULT mm = ::timeGetDevCaps(&timeCaps, sizeof(TIMECAPS));
	if (mm == TIMERR_STRUCT)
		return E_UNEXPECTED;
	mMinResolution = timeCaps.wPeriodMin;

	return S_OK;
}

STDMETHODIMP_(void) CConsoleObject::FinalRelease()
{
}

STDMETHODIMP CConsoleObject::SetSite(IUnknown* site)
{
	Log(LOG_FUNC, _T("CConsoleObject::SetSite(0x%08p)\n"), site);

	HRESULT hr = S_OK;
	if (site) {
		AddRef();

		hr = IObjectWithSiteImpl<CConsoleObject>::SetSite(site);
		if (FAILED(hr)) return hr;

		CComQIPtr<IWebBrowser2> webBrowser2 = m_spUnkSite;
		if (!webBrowser2) {
			Log(LOG_ERROR, _T("QI IWebBrowser2 failed\n"));
			return E_NOINTERFACE;
		}

		hr = DispEventAdvise(webBrowser2, &DIID_DWebBrowserEvents2);
		if (FAILED(hr)) return hr;

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

		CComQIPtr<IConsoleObject> obj = this;
		hr = rot->Register(ROTFLAGS_REGISTRATIONKEEPSALIVE, obj, moniker, &mRotCookie);
		if (FAILED(hr)) {
			Log(LOG_ERROR, _T("rot->Register() failed: 0x%x\n"), hr);
			return hr;
		}
	} else {
		CComQIPtr<IWebBrowser2> webBrowser2 = m_spUnkSite;
		if (webBrowser2) {
			DispEventUnadvise(webBrowser2, &DIID_DWebBrowserEvents2);
		}

		CComPtr<IRunningObjectTable> rot;
		hr = GetRunningObjectTable(0, &rot);
		if (FAILED(hr)) return hr;

		hr = rot->Revoke(mRotCookie);
		if (FAILED(hr)) return hr;

		IObjectWithSiteImpl<CConsoleObject>::SetSite(site);
		Release();
	}

	return hr;
}

STDMETHODIMP_(void) CConsoleObject::OnNavigateComplete2(IDispatch* dispatch, VARIANT* url)
{
	Log(LOG_FUNC, _T("CConsoleObject::NavigateComplete2(url=%s)\n"), url->bstrVal);

	HRESULT hr;

	/*
	 * Get the IDispatchEx of the global script object
	 */

	CComQIPtr<IWebBrowser2> webBrowser2 = dispatch;
	if (!webBrowser2) return;

	CComPtr<IDispatch> documentDispatch;
	hr = webBrowser2->get_Document(&documentDispatch);
	if (FAILED(hr)) return;

	CComQIPtr<IHTMLDocument> document = documentDispatch;
	if (!document) {
		Log(LOG_ERROR, _T("QI IHTMLDocument failed\n"));
		return;
	}

	CComPtr<IDispatch> scriptDispatch;
	hr = document->get_Script(&scriptDispatch);
	if (FAILED(hr)) {
		Log(LOG_ERROR, _T("get_Script() failed\n"));
		return;
	}

	CComQIPtr<IDispatchEx> scriptDispEx = scriptDispatch;
	if (!scriptDispEx) {
		Log(LOG_ERROR, _T("QI IDispatchEx failed\n"));
		return;
	}

	/*
	 * Determine IE version.
	 * Use console67 to avoid name collision with the Developer Tools of IE8.
	 */

	CRegKey regKey;
	long versionNumber = 0xffffffff; // fail safe
	LONG regError;

	regError = regKey.Open(HKEY_LOCAL_MACHINE, _T("SOFTWARE\\Microsoft\\Internet Explorer"), KEY_READ);
	if (regError == ERROR_SUCCESS) {
		TCHAR versionString[128];
		ULONG size = sizeof(versionString);

		regError = regKey.QueryStringValue(_T("Version"), versionString, &size);
		if (regError == ERROR_SUCCESS) {
			versionString[size - 1] = _T('\0');

			TCHAR* dot = NULL;
			versionNumber = _tcstol(versionString, &dot, 10);
		}
	}

	_bstr_t console67Name(_T("console"));
	if (regError != ERROR_SUCCESS || versionNumber >= 8) {
		console67Name += _T("67");
	}

	_variant_t me(this, true);
	hr = scriptDispEx.PutPropertyAlways(console67Name, &me);
	if (FAILED(hr)) {
		Log(LOG_ERROR, _T("Put(\"console67\") failed: 0x%x\n"), hr);
		return;
	}

	/*CComQIPtr<IHTMLDocument2> document2 = documentDispatch;
	if (!document2) return;

	CComPtr<IHTMLWindow2> htmlWindow2;
	hr = document2->get_parentWindow(&htmlWindow2);
	if (FAILED(hr)) return;
	
	_bstr_t script =
		_T("this.Function;\n")
		_T("var aaa = new Function();\n");
	_variant_t dummy;
	hr = htmlWindow2->execScript(script, _bstr_t(_T("JavaScript")), &dummy);
	if (FAILED(hr)) {
		Log(LOG_ERROR, _T("NavigateComplete2(): execScript() failed: 0x%x\n"), hr);
		return;
	}*/

	DUMP_IDISPATCH(this);

	Log(LOG_FUNC, _T("CConsoleObject::NavigateComplete2() end\n"));
}

STDMETHODIMP CConsoleObject::log(SAFEARRAY* varg)
{
	Log(LOG_FUNC, _T("CConsoleObject::log(0x%p)\n"), varg);

	HRESULT hr;

	_bstr_t bstr;
	hr = BstrPrintf(bstr.GetAddress(), *varg);
	if (FAILED(hr))
		return hr;

	Fire_PrintEvent(LEVEL_NORMAL, bstr);
	Log(LOG_MOREFUNC, _T("CConsoleObject::log() \"%s\"\n"), static_cast<TCHAR*>(bstr));

	return S_OK;
}

STDMETHODIMP CConsoleObject::debug(SAFEARRAY* varg)
{
	Log(LOG_FUNC, _T("CConsoleObject::debug()\n"));
		
	_bstr_t str(L"");

	Fire_PrintEvent(LEVEL_DEBUG, str);
	return S_OK;
}

STDMETHODIMP CConsoleObject::info(SAFEARRAY* varg)
{
	Log(LOG_FUNC, _T("CConsoleObject::info()\n"));
	
	HRESULT hr;
	_bstr_t bstr;

	hr = BstrPrintf(bstr.GetAddress(), *varg);
	if (FAILED(hr))
		return hr;

	Fire_PrintEvent(LEVEL_INFO, bstr);
	return S_OK;
}

STDMETHODIMP CConsoleObject::warn(SAFEARRAY* varg)
{
	Log(LOG_FUNC, _T("CConsoleObject::warn()\n"));
	
	HRESULT hr;
	_bstr_t bstr;

	hr = BstrPrintf(bstr.GetAddress(), *varg);
	if (FAILED(hr))
		return hr;

	Fire_PrintEvent(LEVEL_WARN, bstr);
	return S_OK;
}

STDMETHODIMP CConsoleObject::error(SAFEARRAY* varg)
{
	Log(LOG_FUNC, _T("CConsoleObject::error()\n"));
	
	HRESULT hr;
	_bstr_t bstr;

	hr = BstrPrintf(bstr.GetAddress(), *varg);
	if (FAILED(hr))
		return hr;

	Fire_PrintEvent(LEVEL_ERROR, bstr);
	return S_OK;
}

#pragma push_macro("assert")
#undef assert
STDMETHODIMP CConsoleObject::assert(VARIANT_BOOL expr, SAFEARRAY* varg)
#pragma pop_macro("assert")
{
	Log(LOG_FUNC, _T("CConsoleObject::assert(expr=%s)\n"), expr?_T("true"):_T("false"));
	
	if (expr == VARIANT_FALSE) {
		_bstr_t str(L"");
		Fire_PrintEvent(LEVEL_ERROR, str);
	}

	return S_OK;
}

STDMETHODIMP CConsoleObject::clear(void)
{
	Log(LOG_FUNC, _T("CConsoleObject::clear()\n"));
	Fire_ClearEvent();
	return S_OK;
}

STDMETHODIMP CConsoleObject::trace(void)
{
	Log(LOG_FUNC, _T("CConsoleObject::trace()\n"));
	return S_OK;
}

STDMETHODIMP CConsoleObject::time(BSTR name)
{
	Log(LOG_FUNC, _T("CConsoleObject::time(\"%s\")\n"), name);

	POSITION pos;

	{
		_variant_t variant(name);
		CAtlString atlString(variant);
		// don't ::VariantClear(&variant); It's caller's duty to free parameter BSTR name;
		variant.Detach();

		// two stages SetAt() and SetValueAt(),
		// try to reduce string hashing overhead impact to the timer.
		pos = mTimerMap.SetAt(atlString, 0);
	}

	if (mRunningTimer == 0) {
		::timeBeginPeriod(mMinResolution);
	}
	mRunningTimer++;
	
	mTimerMap.SetValueAt(pos, ::timeGetTime());

	return S_OK;
}

STDMETHODIMP CConsoleObject::timeEnd(BSTR name)
{
	Log(LOG_FUNC, _T("CConsoleObject::timeEnd(\"%s\")\n"), name);

	DWORD end = ::timeGetTime();

	mRunningTimer--;
	if (mRunningTimer == 0) {
		::timeEndPeriod(mMinResolution);
	}

	_variant_t variant(name);
	CAtlString atlName(variant);
	variant.Detach();

	DWORD start;
	mTimerMap.Lookup(atlName, start);

	DWORD diff = end - start;
	// fire print event about this diff

	Log(LOG_FUNC, _T("CConsoleObject::timeEnd() diff: %lums\n"), diff);

	CAtlString toBePrint(_T(""));
	toBePrint.Format(_T("%s: %lums"), static_cast<CAtlString::PCXSTR>(atlName), diff);

	_bstr_t bstrTemp;
	bstrTemp.Attach(toBePrint.AllocSysString());
	Fire_PrintEvent(LEVEL_INFO, bstrTemp);
	// bstrTemp automatically releases the buffer allocated by toBePrint.AllocSysString()

	return S_OK;
}

STDMETHODIMP CConsoleObject::profile(BSTR title)
{
	Log(LOG_FUNC, _T("CConsoleObject::profile(\"%s\")\n"), title);
	return S_OK;
}

STDMETHODIMP CConsoleObject::profileEnd(void)
{
	Log(LOG_FUNC, _T("CConsoleObject::profileEnd()\n"));
	return S_OK;
}

STDMETHODIMP CConsoleObject::count(BSTR name)
{
	Log(LOG_FUNC, _T("CConsoleObject::count(\"%s\")\n"), name);

	// TODO: This function is currently not working.
	// Need to know when the script execution stops, then print the result.
	_variant_t variant(name);
	CAtlString atlString(variant);
	variant.Detach();

	CAtlMap<CAtlString, unsigned long>::CPair* pair = mCountMap.Lookup(atlString);
	if (pair != NULL) {
		pair->m_value++;
	} else {
		mCountMap[atlString] = 0;
	}		

	return S_OK;
}

#define GET_NEXT_VARIANT(variant, type)	\
do {									\
	i++;								\
	if (i >= upper)						\
		return DISP_E_BADPARAMCOUNT;	\
										\
	(variant) = safearray[i];			\
										\
	if ((variant).vt != (type))			\
		return DISP_E_BADVARTYPE;		\
} while (false)

STDMETHODIMP CConsoleObject::BstrPrintf(BSTR* ret, const SAFEARRAY& varg) const
{
	//DUMP_VARIANTSAFEARRAY(&varg);

	if (ret == NULL)
		return E_POINTER;

	CComSafeArray<VARIANT> safearray;

	HRESULT hr = safearray.Attach(&varg);
	if (FAILED(hr)) return hr;

	if (safearray.GetDimensions() != 1)
		return E_INVALIDARG;

	CAtlString result(_T(""));

	_variant_t current;
	int upper = safearray.GetUpperBound() + 1;
	for (int i = safearray.GetLowerBound(); i < upper; i++) {
		current.Attach(safearray[i]);

		if (current.vt == VT_I4) {
			result.AppendFormat(_T("%ld"), current.lVal);
		} else if (current.vt == VT_BSTR) {
			_bstr_t str(current.bstrVal);
			int specCount = 0;
			TCHAR* lastFound = str;
			TCHAR* found;
			TCHAR next;
			int specLength = 0;

			found = _tcschr(str, _T('%'));

			if (found == NULL) {
				result += static_cast<TCHAR*>(str);
			} else {
				while (found) {
					next = *(found + 1);

					switch (next) {
					case _T('%') : {
						specLength = 2;
						result.CopyChars(result.GetBuffer(), lastFound,
							static_cast<int>(found + 1 - lastFound));
								   }
					case _T('s') : {
						specLength = 2;

						_variant_t varstr;
						GET_NEXT_VARIANT(varstr, VT_BSTR);

						result.Append(lastFound, static_cast<int>(found - lastFound));
						result += varstr.bstrVal;

						break;
								   }
					case _T('d') : {
						specLength = 2;

						_variant_t varint;
						GET_NEXT_VARIANT(varint, VT_I4);

						result.Append(lastFound, static_cast<int>(found - lastFound));
						result.AppendFormat(_T("%ld"), varint.lVal);

						break;
								   }
					case _T('f') : {
						specLength = 2;

						//GET_NEXT_VARIANT(VT_R8);
						break;
								   }
					case _T('x') :
						break;
					default:
						break;
					}

					found += specLength;
					lastFound = found;
					found = _tcschr(found, _T('%'));
				} // while loop
			}
		} else {
			result += _T("[unsupported argument type]");
		}

		result += _T(' ');
		current.Detach();
	} // for loop

	safearray.Detach();

	result.SetAt(result.GetLength() - 1, _T('\0'));
	BSTR tmp = result.AllocSysString();
	if (tmp == NULL) return E_OUTOFMEMORY;

	*ret = tmp;
	return S_OK;
}

#undef GET_NEXT_VARIANT
