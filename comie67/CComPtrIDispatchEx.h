/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include <atlcomcli.h>
#include <dispex.h>

// TODO: mimic CComPtr<IDispatch> implementation
template <>
class CComPtr<IDispatchEx> : public CComPtrBase<IDispatchEx>
{
public:
	CComPtr() throw()
	{
	}
	CComPtr(IDispatchEx* lp) throw() :
	CComPtrBase<IDispatchEx>(lp)
	{
	}
	CComPtr(const CComPtr<IDispatchEx>& lp) throw() :
	CComPtrBase<IDispatchEx>(lp.p)
	{
	}
	IDispatchEx* operator=(IDispatchEx* lp) throw()
	{
		return static_cast<IDispatchEx*>(AtlComPtrAssign((IUnknown**)&p, lp));
	}
	IDispatchEx* operator=(const CComPtr<IDispatchEx>& lp) throw()
	{
		return static_cast<IDispatchEx*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
	}
};

/*class CIDispatchExHelper {
public:
	explicit CIDispatchExHelper(IDispatchEx* disp) : mDisp(disp) {}
	~CIDispatchExHelper() {}

	HRESULT Get(BSTR name, long* num) const;
	HRESULT Get(BSTR name, BSTR* str) const;
	HRESULT Get(BSTR name, IDispatch** disp) const;
	HRESULT Get(BSTR name, IDispatchEx** dispex) const;
	HRESULT Get(BSTR name, VARIANT* variant) const;

	HRESULT Put(BSTR name, long num) const;
	HRESULT Put(BSTR name, BSTR str) const;
	HRESULT Put(BSTR name, IDispatch* disp) const;
	HRESULT Put(BSTR name, VARIANT* variant) const;

	HRESULT Call(BSTR name, long num, ...) const;

	bool Has(BSTR name) const;
	HRESULT Delete(BSTR name) const;

private:
	CIDispatchExHelper();
	CIDispatchExHelper(const CIDispatchExHelper&);
	CIDispatchExHelper& operator=(const CIDispatchExHelper&);

	CComPtr<IDispatchEx> mDisp;
};*/
