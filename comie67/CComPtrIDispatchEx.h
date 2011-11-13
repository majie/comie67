/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include <atlcomcli.h>
#include <dispex.h>

/*
 * CComPtr<IDispatchEx> mimics the implementation of CComPtr<IDispatch>.
 * See atlcomcli.h for the implementation.
 *
 * Currently only case-sensitive names look up is supported.
 *
 * CComPtr<IDispatchEx> inherits CComPtr<IDispatch>, so all CComPtr<IDispatch>
 * functions, like GetPropertyByName() and Invoke0() etc., can be called from
 * CComPtr<IDispatehEx> instances.
 */
template <>
class CComPtr<IDispatchEx> : public CComPtr<IDispatch>
{
public:
	CComPtr() {}
	virtual ~CComPtr() {}

	_NoAddRefReleaseOnCComPtr<IDispatchEx>* operator->() const
	{
		ATLASSERT(p!=NULL);
		return (_NoAddRefReleaseOnCComPtr<IDispatchEx>*)p;
	}

	/*
	 * IDispatchEx specific helper functions
	 *
	 * For GetProperty and Invoke operations,
	 * call corresponding CComPtr<IDispatch> functions.
	 */
	STDMETHOD(PutPropertyAlways)(DISPID dispId, VARIANT* variant);
	STDMETHOD(PutPropertyAlways)(LPCOLESTR name, VARIANT* variant);

	STDMETHOD(HasProperty)(DISPID dispId, BOOL* has);
	STDMETHOD(HasProperty)(LPCOLESTR name, BOOL* has);

	STDMETHOD(DeleteProperty)(DISPID dispId)
	{
		return static_cast<IDispatchEx*>(p)->DeleteMemberByDispID(dispId);
	}

	STDMETHOD(DeleteProperty)(LPCOLESTR name)
	{
		_bstr_t bstr_name(name);
		return static_cast<IDispatchEx*>(p)->DeleteMemberByName(bstr_name,
			fdexNameCaseSensitive);
	}

private:
	CComPtr(const CComPtr&);
	CComPtr& operator=(const CComPtr&);
};

STDMETHODIMP CComPtr<IDispatchEx>::PutPropertyAlways(DISPID dispId, VARIANT* variant)
{
	DISPID dispIdPut = DISPID_PROPERTYPUT;

	DISPPARAMS params;
	params.cArgs = 1;
	params.rgvarg = variant;
	params.cNamedArgs = 1;
	params.rgdispidNamedArgs = &dispIdPut;

	EXCEPINFO ei;

	HRESULT hr = static_cast<IDispatchEx*>(p)->InvokeEx(dispId,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, NULL, &ei, NULL);
	if (FAILED(hr))	return hr;

	return S_OK;
}

STDMETHODIMP CComPtr<IDispatchEx>::PutPropertyAlways(LPCOLESTR name, VARIANT* variant)
{
	if (variant == NULL)
		return E_POINTER;

	_bstr_t bstr_name(name);
	DISPID dispId;
	HRESULT hr = static_cast<IDispatchEx*>(p)->GetDispID(bstr_name,
		fdexNameEnsure | fdexNameCaseSensitive, &dispId);
	if (FAILED(hr)) return hr;

	return PutPropertyAlways(dispId, variant);
}

STDMETHODIMP CComPtr<IDispatchEx>::HasProperty(DISPID dispId, BOOL* has)
{
	if (has == NULL) {
		return E_POINTER;
	}

	_bstr_t dummy;
	HRESULT hr = static_cast<IDispatchEx*>(p)->GetMemberName(dispId,
		dummy.GetAddress());
	switch (hr) {
		case S_OK:
			*has = TRUE;
			break;
		case DISP_E_UNKNOWNNAME:
			*has = FALSE;
			break;
		default:
			return hr;
	}

	return S_OK;
}

STDMETHODIMP CComPtr<IDispatchEx>::HasProperty(LPCOLESTR name, BOOL* has)
{
	if (has == NULL) {
		return E_POINTER;
	}

	_bstr_t bstr_name(name);
	DISPID dispId;
	HRESULT hr = static_cast<IDispatchEx*>(p)->GetDispID(bstr_name,
		fdexNameCaseSensitive, &dispId);
	switch (hr) {
		case S_OK:
			*has = TRUE;
			break;
		case DISP_E_UNKNOWNNAME:
			*has = FALSE;
			break;
		default:
			return hr;
	}

	return S_OK;
}
