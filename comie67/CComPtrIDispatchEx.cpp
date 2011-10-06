#include "stdafx.h"

#include <dispex.h>

#include "CIDispatchExHelper.h"

HRESULT CIDispatchExHelper::Get(BSTR name, IDispatch** disp) const
{
	if (disp == NULL) return E_POINTER;

	VARIANT variant;
	HRESULT hr = Get(name, &variant);
	if (FAILED(hr)) return hr;

	if (variant.vt == VT_DISPATCH) {
		*disp = variant.pdispVal;
		(*disp)->AddRef();
	} else {
		return DISP_E_TYPEMISMATCH;
	}

	return S_OK;
}

HRESULT CIDispatchExHelper::Get(BSTR name, IDispatchEx** dispex) const
{
	if (dispex == NULL) return E_POINTER;

	VARIANT variant;
	HRESULT hr = Get(name, &variant);
	if (FAILED(hr)) return hr;

	if (variant.vt == VT_DISPATCH) {
		IDispatchEx* ex;
		hr = variant.pdispVal->QueryInterface(IID_IDispatchEx, reinterpret_cast<void**>(&ex));
		if (FAILED(hr)) return hr;

		*dispex = ex;
		ex->AddRef();
	} else {
		return DISP_E_TYPEMISMATCH;
	}

	return S_OK;
}

HRESULT CIDispatchExHelper::Get(BSTR name, VARIANT* variant) const
{
	if (variant == NULL) return E_POINTER;

	DISPID dispId;
	HRESULT hr = mDisp->GetDispID(name, fdexNameCaseSensitive, &dispId);
	if (FAILED(hr)) return hr;

	VARIANTARG varg;
	DISPID id;
	DISPPARAMS dispparams = {&varg, &id, 0, 0};
	VARIANT varResult; 

	hr = mDisp->InvokeEx(dispId, LOCALE_USER_DEFAULT,
		DISPATCH_PROPERTYGET, &dispparams, &varResult, NULL, NULL);
	if (FAILED(hr)) return hr;

	*variant = varResult;

	return S_OK;
}

HRESULT CIDispatchExHelper::Put(BSTR name, IDispatch* disp) const
{
	if (disp == NULL)
		return E_POINTER;

	VARIANTARG args;
	VariantInit(&args);
	args.vt = VT_DISPATCH;
	args.pdispVal = disp;

	disp->AddRef();
	HRESULT hr = Put(name, &args);

	VariantClear(&args);

	if (FAILED(hr))	{
		disp->Release();
		return hr;
	}

	return S_OK;
}

HRESULT CIDispatchExHelper::Put(BSTR name, VARIANT* variant) const
{
	if (variant == NULL)
		return E_POINTER;

	DISPID dispId;
	HRESULT hr = mDisp->GetDispID(name, fdexNameEnsure | fdexNameCaseSensitive,	&dispId);
	if (FAILED(hr)) return hr;

	/*VARIANTARG args;
	VariantInit(&args);
	VariantCopy(args, variant);*/

	DISPID dispIdPut = DISPID_PROPERTYPUT;

	DISPPARAMS params;
	params.cArgs = 1;
	//params.rgvarg = &args;
	params.rgvarg = variant;
	params.cNamedArgs = 1;
	params.rgdispidNamedArgs = &dispIdPut;

	EXCEPINFO ei;

	hr = mDisp->InvokeEx(dispId, LOCALE_USER_DEFAULT,
		DISPATCH_PROPERTYPUT, &params, NULL, &ei, NULL);
	
	//VariantClear(&args);

	if (FAILED(hr))	return hr;

	return S_OK;
}

HRESULT CIDispatchExHelper::Call(BSTR name, long num, ...) const
{
	return S_OK;
}

bool CIDispatchExHelper::Has(BSTR name) const
{
	DISPID dispId;
	HRESULT hr = mDisp->GetDispID(name, fdexNameCaseSensitive,	&dispId);
	switch (hr) {
		case S_OK:
			return true;
		case DISP_E_UNKNOWNNAME:
			return false;
		default:
			//impossible
			return false;
	}
}

HRESULT CIDispatchExHelper::Delete(BSTR name) const
{
	return mDisp->DeleteMemberByName(name, fdexNameCaseSensitive);
}
