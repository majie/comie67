/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include <stdarg.h>
#include <atlcomcli.h>
#include <dispex.h>

/*
 * CComPtr<IDispatchEx> mimics the implementation of CComPtr<IDispatch>.
 * See atlcomcli.h for the implementation.
 *
 */
template <>
class CComPtr<IDispatchEx> : public CComPtrBase<IDispatchEx>
{
public:
	CComPtr() :
		mCaseSens(TRUE)
	{
	}
	CComPtr(IDispatchEx* disp) :
		CComPtrBase<IDispatchEx>(disp),
		mCaseSens(TRUE)
	{
	}
	CComPtr(const CComPtr<IDispatchEx>& disp) :
		CComPtrBase<IDispatchEx>(disp.p),
		mCaseSens(TRUE)
	{
	}
	IDispatchEx* operator=(IDispatchEx* disp)
	{
		return static_cast<IDispatchEx*>(
			AtlComPtrAssign(reinterpret_cast<IUnknown**>(&p), disp));
	}
	IDispatchEx* operator=(const CComPtr<IDispatchEx>& disp)
	{
		return static_cast<IDispatchEx*>(
			AtlComPtrAssign(reinterpret_cast<IUnknown**>(&p), disp.p));
	}

	//IDispatchEx specific helper functions
	STDMETHOD(PutPropertyAlways)(DISPID dispId, VARIANT* variant)
	{
		DISPID dispIdPut = DISPID_PROPERTYPUT;
		DISPPARAMS params = {variant, &dispIdPut, 1, 1};

		return p->InvokeEx(dispId, LOCALE_USER_DEFAULT,
			DISPATCH_PROPERTYPUT, &params, NULL, NULL, NULL);
	}
	STDMETHOD(PutPropertyAlways)(BSTR name, VARIANT* variant);

	STDMETHOD(GetProperty)(DISPID dispId, VARIANT* variant)
	{
		DISPPARAMS params = {NULL, NULL, 0, 0};
		return p->InvokeEx(dispId, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
			&params, variant, NULL, NULL);
	}
	STDMETHOD(GetProperty)(BSTR name, VARIANT* variant);

	STDMETHOD(Invoke)(DISPID dispId, VARIANT* varRet, ULONG n, ...)
	{
		DISPPARAMS params = {NULL, NULL, 0, 0};
		if (n > 0) {
			va_list vargs;

			va_start(vargs, n);
			params.rgvarg = va_arg(vargs, VARIANT*);
			va_end(vargs);

			params.cArgs = n;
		}
		return p->InvokeEx(dispId, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
			&params, varRet, NULL, NULL);
	}
	STDMETHOD(Invoke)(BSTR name, VARIANT* varRet, ULONG n, ...);

	STDMETHOD(HasProperty)(DISPID dispId, BOOL* has);
	STDMETHOD(HasProperty)(BSTR name, BOOL* has);

	STDMETHOD(DeleteProperty)(DISPID dispId)
	{
		return p->DeleteMemberByDispID(dispId);
	}

	STDMETHOD(DeleteProperty)(BSTR name)
	{
		return p->DeleteMemberByName(name,
			mCaseSens?fdexNameCaseSensitive:fdexNameCaseInsensitive);
	}

	BOOL mCaseSens;
};

STDMETHODIMP CComPtr<IDispatchEx>::PutPropertyAlways(BSTR name, VARIANT* variant)
{
	DISPID dispId;
	HRESULT hr = p->GetDispID(name,	fdexNameEnsure |
		(mCaseSens?fdexNameCaseSensitive:fdexNameCaseInsensitive), &dispId);
	if (FAILED(hr)) return hr;

	return PutPropertyAlways(dispId, variant);
}

STDMETHODIMP CComPtr<IDispatchEx>::GetProperty(BSTR name, VARIANT* variant)
{
	DISPID dispId;
	HRESULT hr = p->GetDispID(name,
		(mCaseSens?fdexNameCaseSensitive:fdexNameCaseInsensitive), &dispId);
	if (FAILED(hr)) return hr;

	return GetProperty(dispId, variant);
}

STDMETHODIMP CComPtr<IDispatchEx>::Invoke(BSTR name, VARIANT* varRet, ULONG n, ...)
{
	DISPID dispId;
	HRESULT hr = p->GetDispID(name,
		(mCaseSens?fdexNameCaseSensitive:fdexNameCaseInsensitive), &dispId);
	if (FAILED(hr)) return hr;

	va_list vargs;
	va_start(vargs, n);
	hr = Invoke(dispId, varRet, n, vargs);
	va_end(vargs);

	return hr;
}

STDMETHODIMP CComPtr<IDispatchEx>::HasProperty(DISPID dispId, BOOL* has)
{
	if (has == NULL) {
		return E_POINTER;
	}

	_bstr_t dummy;
	HRESULT hr = p->GetMemberName(dispId, dummy.GetAddress());
	switch (hr) {
		case S_OK:
			*has = TRUE;
			break;
		case DISP_E_UNKNOWNNAME:
			*has = FALSE;
			hr = S_OK;
			break;
		default:
			break;
	}

	return hr;
}

STDMETHODIMP CComPtr<IDispatchEx>::HasProperty(BSTR name, BOOL* has)
{
	if (has == NULL) {
		return E_POINTER;
	}

	DISPID dispId;
	HRESULT hr = p->GetDispID(name,
		mCaseSens?fdexNameCaseSensitive:fdexNameCaseInsensitive, &dispId);
	switch (hr) {
		case S_OK:
			*has = TRUE;
			break;
		case DISP_E_UNKNOWNNAME:
			*has = FALSE;
			hr = S_OK;
			break;
		default:
			break;
	}

	return hr;
}
