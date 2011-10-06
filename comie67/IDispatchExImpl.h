/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include <stdlib.h>

#include <dispex.h>

#include <comutil.h>
#include <comdef.h>

#include <atlcoll.h>
#include <atlconv.h>
#include <atlstr.h>

#include "dump.h"
#include "logger.h"

template <typename T>
class ATL_NO_VTABLE IDispatchExImpl : public T
{
public:
	IDispatchExImpl() : mNextId(DISPID_VALUE + 1), mTypeInfo(NULL), mIndexInMap(0){}
	virtual ~IDispatchExImpl();

	// IDispatch methods
	STDMETHOD(GetTypeInfoCount)(UINT* pctinfo);
	STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
	STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames,	LCID lcid, DISPID* rgdispid);
	STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams,
		VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

	// IDispatchEx methods
	STDMETHOD(GetDispID)(BSTR name, DWORD flags, DISPID* pid);
	STDMETHOD(InvokeEx)(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS* pdp,
		VARIANT* pvarRes, EXCEPINFO* pei, IServiceProvider* pspCaller);
	STDMETHOD(DeleteMemberByName)(BSTR bstrName, DWORD flags);
	STDMETHOD(DeleteMemberByDispID)(DISPID id);
	STDMETHOD(GetMemberProperties)(DISPID id, DWORD grfdexFetch, DWORD* pgrfdex);
	STDMETHOD(GetMemberName)(DISPID id, BSTR* pbstrName);
	STDMETHOD(GetNextDispID)(DWORD grfdex, DISPID id, DISPID* pid);
	STDMETHOD(GetNameSpaceParent)(IUnknown** ppunk) {return E_NOTIMPL;}

protected:
	/*
	 * Call InitTypeInfo() before calling other methods of IDispatchExImpl, otherwise methods
	 * defined in typelib won't be add to the entry maps and can't be called by the scripts.
	 *
	 * BOOL all = TRUE if you want all methods defined in the typelib to be added to the maps.
	 * if all = FALSE, only methods listed in kInternalMethods will be added to the maps.
	 */
	STDMETHOD(InitTypeInfo)(BOOL all = TRUE);
	STDMETHOD(AddEntryToMaps)(DISPID id, BSTR name, bool internal, DWORD attr);

private:
	IDispatchExImpl(const IDispatchExImpl&);
	IDispatchExImpl& operator=(const IDispatchExImpl&);

	STDMETHOD(InvokeImpl)(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS* pdp,
		VARIANT* pvarRes, EXCEPINFO* pei, IServiceProvider* pspCaller, UINT* puArgErr);
	STDMETHOD(GetIDsOfNamesImpl)(LPOLESTR* names, UINT countNames, bool caseSens, DISPID* dispids);

	static int BsearchCompare(const void* v1, const void* v2) {
		return _tcscmp(*static_cast<const LPCOLESTR*>(v1), *static_cast<const LPCOLESTR*>(v2));
	}

	class CDispatchExEntry;
	CAtlMap<DISPID, CDispatchExEntry*>											mIdToEntryMap;
	CAtlMap<CAtlString, CDispatchExEntry*, CStringElementTraits<CAtlString> >	mNameToEntryMap;
	CAtlMap<CAtlString, CDispatchExEntry*, CStringElementTraitsI<CAtlString> >	mNameToEntryIMap;

	DISPID		mNextId;
	ITypeInfo*	mTypeInfo;
	POSITION	mIndexInMap;

	static const LPCOLESTR kInternalMethods[15];

	class CDispatchExEntry {
	public:
		CDispatchExEntry(DISPID id, BSTR name, bool internal, DWORD attr) :
		  mDispId(id), mDeleted(false), mInternal(internal), mAttributes(attr)
		  {
			  mName = _bstr_t(name, true);
			  mVariant = _variant_t();
		  }

		DISPID		GetDispID(void)		const {return mDispId;}
		BSTR		GetName(void)		const {return mName.copy(true);}
		BSTR		GetNameRef(void)	const {return mName.copy(false);}
		VARIANT		GetVariant(void)	const {return mVariant;}
		DWORD		GetAttributes(void)	const {return mAttributes;}

		bool IsDeleted(void)	const {return mDeleted;}
		bool IsInternal(void)	const {return mInternal;}

		bool MarkAsDeleted(void)
		{
			if (IsInternal())
				return false;
			mDeleted = true;
			mVariant.Clear();
			return true;
		}

		void SetVariant(const VARIANT& v)
		{
			mDeleted = false;
			mVariant = v;
		}

		void AttachVariant(VARIANT& v)
		{
			mDeleted = false;
			mVariant.Clear();
			mVariant.Attach(v);
		}

		void SetAttributes(DWORD attr) {mAttributes = attr;}

	private:
		CDispatchExEntry();
		CDispatchExEntry(const CDispatchExEntry&);
		CDispatchExEntry& operator=(const CDispatchExEntry&);

		DISPID		mDispId;
		_bstr_t		mName; // _bstr_t and _variant_t are more lightweight than CComBSTR and CComVariant
		_variant_t	mVariant;
		bool		mDeleted;
		bool		mInternal; // Internal COM methods. e.g. QueryInterface() etc.
		DWORD		mAttributes; //fdexPropCanGet etc.
	};
};



template <typename T>
IDispatchExImpl<T>::~IDispatchExImpl()
{
	CDispatchExEntry* entry = NULL;
	POSITION pos = mIdToEntryMap.GetStartPosition();
	entry = mIdToEntryMap.GetValueAt(pos);
	while (pos != NULL) {
		delete entry;
		mIdToEntryMap.SetValueAt(pos, NULL);
		entry = mIdToEntryMap.GetNextValue(pos);
	}

	mIdToEntryMap.RemoveAll();
	mNameToEntryMap.RemoveAll();
	mNameToEntryIMap.RemoveAll();

	mNextId = 0;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::GetTypeInfoCount(UINT* pctinfo)
{
	if (pctinfo == NULL)
		return E_POINTER;
	*pctinfo = 1;
	return S_OK;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	if (ppTInfo == NULL)
		return E_POINTER;
	*ppTInfo = NULL;

	if(iTInfo != 0)
		return DISP_E_BADINDEX;

	mTypeInfo->AddRef();
	*ppTInfo = mTypeInfo;

	return S_OK;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames,
	UINT cNames, LCID lcid, DISPID* rgdispid)
{
	return GetIDsOfNamesImpl(rgszNames, cNames, false, rgdispid);
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
	DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	return InvokeImpl(dispIdMember, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, NULL, puArgErr);
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::GetDispID(BSTR name, DWORD flags, DISPID* pid)
{
	Log(LOG_FUNC, _T("IDispatchExImpl::GetDispID(\"%s\", 0x%x) called\n"), name, flags);

	if (pid == NULL)
		return E_POINTER;

	bool caseSens;
	if (flags & fdexNameCaseSensitive) {
		caseSens = true;
	} else if (flags & fdexNameCaseInsensitive) {
		caseSens = false;
	}

	HRESULT hr = GetIDsOfNamesImpl(&name, 1, caseSens, pid);
	if (SUCCEEDED(hr)) {
		return S_OK;
	} else if (hr == DISP_E_UNKNOWNNAME) {
		if (flags & fdexNameEnsure) {
			DISPID id = mNextId;
			hr = AddEntryToMaps(id, name, false,
				fdexPropDynamicType | fdexPropCanGet | fdexPropCanPut | fdexPropCanCall);
			if (FAILED(hr)) return hr;
			*pid = id;
			return S_OK;
		} else {
			return DISP_E_UNKNOWNNAME;
		}
	} else {
		return hr;
	}
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS* pdp,
	VARIANT* pvarRes, EXCEPINFO* pei, IServiceProvider* pspCaller)
{
	return InvokeImpl(id, lcid, wFlags, pdp, pvarRes, pei, pspCaller, NULL);
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::DeleteMemberByName(BSTR name, DWORD flags)
{
	Log(LOG_FUNC, _T("IDispatchExImpl::DeleteMemberByName(%s) called\n"), name);

	CDispatchExEntry* entry = NULL;

	if (flags & fdexNameCaseSensitive) {
		if (!mNameToEntryMap.Lookup(name, entry))
			return DISP_E_MEMBERNOTFOUND;
	} else if (flags & fdexNameCaseInsensitive) {
		if (!mNameToEntryIMap.Lookup(name, entry))
			return DISP_E_MEMBERNOTFOUND;
	}

	if (entry && !entry->MarkAsDeleted())
		return S_FALSE;

	return S_OK;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::DeleteMemberByDispID(DISPID id)
{
	Log(LOG_FUNC, _T("IDispatchExImpl::DeleteMemberByDispID(0x%x) called\n"), id);

	CDispatchExEntry* entry = NULL;
	if (!mIdToEntryMap.Lookup(id, entry)) {
		return DISP_E_MEMBERNOTFOUND;
	}
	if (entry && !entry->MarkAsDeleted())
		return S_FALSE;
	return S_OK;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD* pgrfdex)
{
	Log(LOG_FUNC, _T("IDispatchExImpl::GetMemberProperties(0x%x, 0x%x) called\n"), id, grfdexFetch);
	if (pgrfdex == NULL)
		return  E_POINTER;

	CDispatchExEntry* entry = NULL;
	if (!mIdToEntryMap.Lookup(id, entry)) {
		return DISP_E_MEMBERNOTFOUND;
	}

	*pgrfdex = entry->GetAttributes();

	return S_OK;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::GetMemberName(DISPID id, BSTR* pbstrName)
{
	Log(LOG_MOREFUNC, _T("IDispatchExImpl::GetMemberName(0x%x) called\n"), id);

	CDispatchExEntry* entry = NULL;

	bool found = mIdToEntryMap.Lookup(id, entry);
	if (!found || entry == NULL)
		return DISP_E_MEMBERNOTFOUND;
	if (entry->IsDeleted())
		return DISP_E_MEMBERNOTFOUND;

	*pbstrName = entry->GetName();

	return S_OK;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::GetNextDispID(DWORD grfdex, DISPID id, DISPID* pid)
{
	Log(LOG_MOREFUNC, _T("IDispatchExImpl::GetNextDispID(0x%x) called\n"), id);

	// grfdex is not used
	// I don't know what is the meaning of fdexEnumDefault and fdexEnumAll

	if (pid == NULL)
		return E_POINTER;

	CDispatchExEntry* entry = NULL;
	if (id == DISPID_STARTENUM) {
		mIndexInMap = mIdToEntryMap.GetStartPosition();
		entry = mIdToEntryMap.GetValueAt(mIndexInMap);
	} else {
		entry = mIdToEntryMap.GetValueAt(mIndexInMap);
	}

	while (mIndexInMap != NULL) {
		if (entry && !entry->IsDeleted() && !entry->IsInternal()) {
			*pid = entry->GetDispID();
			mIdToEntryMap.GetNextValue(mIndexInMap);
			return S_OK;
		}
		entry = mIdToEntryMap.GetNextValue(mIndexInMap);
	}

	return S_FALSE;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::InitTypeInfo(BOOL all)
{
	mIdToEntryMap.InitHashTable(75);
	mNameToEntryMap.InitHashTable(75);
	mNameToEntryIMap.InitHashTable(75);

	TCHAR filename[MAX_PATH + 1];
	DWORD length;

	length = ::GetModuleFileName(_AtlBaseModule.GetModuleInstance(), filename, MAX_PATH);
	if (length == 0)
		return HRESULT_FROM_WIN32(::GetLastError());

	ITypeLib* typeLib;

	HRESULT hr;
	hr = LoadTypeLib(filename, &typeLib);
	if (FAILED(hr)) return hr;

	ITypeInfo* typeInfo;
	hr = typeLib->GetTypeInfoOfGuid(__uuidof(T), &typeInfo);
	if (FAILED(hr)) return hr;

	mTypeInfo = typeInfo;

	TYPEATTR* pta;
	hr = mTypeInfo->GetTypeAttr(&pta);
	if (FAILED(hr)) return hr;

	FUNCDESC* pfd;
	BSTR name;
	LPOLESTR olestr;

	// use _bstr_t to safely convert BSTR to LPOLESTR, which can be LPWSTR or LPSTR.
	// _bstr_t doesn't provide operator&(), so use _bstr_t::Attach(name);
	_bstr_t name2;
	bool isInternal;
	bool bsearchResult;
	DWORD type;
	int count = pta->cFuncs;

	mIdToEntryMap.DisableAutoRehash();
	mNameToEntryMap.DisableAutoRehash();
	mNameToEntryIMap.DisableAutoRehash();

	for (int i = 0; i < count; i++) {
		hr = mTypeInfo->GetFuncDesc(i, &pfd);
		if (FAILED(hr)) {
			mTypeInfo->ReleaseTypeAttr(pta);
			return hr;
		}

		hr = mTypeInfo->GetDocumentation(pfd->memid, &name, NULL, NULL, NULL);
		if (FAILED(hr)) {
			mTypeInfo->ReleaseFuncDesc(pfd);
			mTypeInfo->ReleaseTypeAttr(pta);
			return hr;
		}

		//Log(_T("0x%x, %s\n"), pfd->memid, name);

		/* Sample output of the line above:
		*
		* 0x60000000, QueryInterface
		* 0x60000001, AddRef
		* 0x60000002, Release
		* 0x60010000, GetTypeInfoCount
		* 0x60010001, GetTypeInfo
		* 0x60010002, GetIDsOfNames
		* 0x60010003, Invoke
		* 0x60020000, GetDispID
		* 0x60020001, RemoteInvokeEx
		* 0x60020002, DeleteMemberByName
		* 0x60020003, DeleteMemberByDispID
		* 0x60020004, GetMemberProperties
		* 0x60020005, GetMemberName
		* 0x60020006, GetNextDispID
		* 0x60020007, GetNameSpaceParent
		* 0x1, log
		* 0x2, debug
		* 0x3, info
		* 0x4, warn
		* 0x5, error
		* 0x6, assert
		* 0x7, clear
		* 0x8, trace
		* 
		* Notice that there is a big empty hole between trace(0x8) and QueryInterface(0x60000000).
		* Although it seems to be safe to allocate new DISPIDs in this hole,
		* we'll need some protection for the internal methods like QueryInterface etc.
		*/

		name2.Attach(name);
		olestr = static_cast<LPOLESTR>(name2);

		bsearchResult = bsearch(&olestr, kInternalMethods, sizeof(kInternalMethods) / sizeof(LPCOLESTR),
			sizeof(LPCOLESTR), BsearchCompare) != NULL;

		isInternal = all ? true : bsearchResult;

		name2.Detach();

		if (pfd->invkind & INVOKE_FUNC) {
			type = fdexPropCanCall;
		}
		else {
			type = 0;
			if (pfd->invkind & INVOKE_PROPERTYGET)
				type |= INVOKE_PROPERTYGET;
			if (pfd->invkind & INVOKE_PROPERTYPUT)
				type |= fdexPropCanPut;
			else if (pfd->invkind & INVOKE_PROPERTYPUTREF)
				type |= fdexPropCanPutRef;
		}

		if (type == 0)
			type = grfdexPropCannotAll;

		hr = AddEntryToMaps(pfd->memid, name, isInternal, type);
		if (FAILED(hr)) {
			mTypeInfo->ReleaseFuncDesc(pfd);
			mTypeInfo->ReleaseTypeAttr(pta);
			::SysFreeString(name);
			return hr;
		}

		::SysFreeString(name);
		mTypeInfo->ReleaseFuncDesc(pfd);
	}

	mIdToEntryMap.EnableAutoRehash();
	mNameToEntryMap.EnableAutoRehash();
	mNameToEntryIMap.EnableAutoRehash();

	mTypeInfo->ReleaseTypeAttr(pta);

#ifdef _DEBUG
	mIdToEntryMap.AssertValid();
	mNameToEntryMap.AssertValid();
	mNameToEntryIMap.AssertValid();
#endif

	return S_OK;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::AddEntryToMaps(DISPID id, BSTR name, bool internal, DWORD attr)
{
	Log(LOG_MOREFUNC, _T("IDispatchExImpl::AddEntryToMaps(0x%x, \"%s\", %d, 0x%x) called\n"),
		id, name, internal, attr);

	CDispatchExEntry* entry = new CDispatchExEntry(id, name, internal, attr);
	if (entry == NULL)
		return E_OUTOFMEMORY;

	// take IDispatchEx internal methods into account. if mNextId already exists, jump to the next one
	while (mIdToEntryMap.Lookup(id) != NULL)
		id++;

	mIdToEntryMap.SetAt(id, entry);
	
	_variant_t var(name);
	CAtlString atlStr(var);

	// let the caller free BSTR name.
	var.Detach();

	mNameToEntryMap.SetAt(atlStr, entry);
	mNameToEntryIMap.SetAt(atlStr, entry);

	if (!internal)
		mNextId = id + 1;
	return S_OK;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::InvokeImpl(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS* pdp,
	VARIANT* pvarRes, EXCEPINFO* pei, IServiceProvider* pspCaller, UINT* puArgErr)
{
	Log(LOG_FUNC, _T("IDispatchExImpl::InvokeExImpl(0x%x, 0x%x) called\n"), id, wFlags);

	HRESULT hr = S_OK;

	CDispatchExEntry* entry = NULL;
	if (!mIdToEntryMap.Lookup(id, entry)) {
		return DISP_E_MEMBERNOTFOUND;
	}
	if (entry && !entry->IsInternal()) {
		if (wFlags & DISPATCH_CONSTRUCT) {
			// DISPATCH_CONSTRUCT is set when a JScript function is called as a constructor

			// see Re: DISPID_CONSTRUCTOR/DISPID_DESTRUCTOR
			// see http://www.tech-archive.net/Archive/Development/microsoft.public.win32.programmer.ole/2007-12/msg00001.html

			// function f() {}
			// f();     // DISPATCH_METHOD is set and DISPID is DISPID_VALUE
			// new f(); // DISPATCH_CONSTRUCT is set and DISPID is DISPID_VALUE
		}
		if (wFlags & DISPATCH_PROPERTYGET) {
			if (pvarRes == NULL)
				return E_POINTER;

			if (entry->IsDeleted())
				return DISP_E_MEMBERNOTFOUND;

			hr = ::VariantClear(pvarRes);
			if (FAILED(hr))
				return hr;

			*pvarRes = entry->GetVariant();
		}
		if (wFlags & DISPATCH_PROPERTYPUT || wFlags & DISPATCH_PROPERTYPUTREF) {
			if (pdp != NULL && pdp->cArgs > 0 && pdp->cNamedArgs > 0) {
				if (pdp->rgdispidNamedArgs[0] != DISPID_PROPERTYPUT)
					return E_INVALIDARG;
				//DUMP_DISPPARAMS(pdp);
				
				DWORD attr = entry->GetAttributes();
				if (attr & fdexPropCanPut)
					entry->SetVariant(pdp->rgvarg[0]);
				else if (attr & fdexPropCanPutRef)
					entry->AttachVariant(pdp->rgvarg[0]);

				if (entry->GetVariant().vt == VT_DISPATCH) {
					IDispatch* pdisp = entry->GetVariant().pdispVal;
					CComPtr<IDispatch> disp;
					_variant_t var;
					DISPID mydispid;
					disp.Attach(pdisp);

					Log(LOG_MOREFUNC, _T("idispatch: 0x%p\n"), pdisp);

					HRESULT hr = disp.GetProperty(0, &var);
					if (FAILED(hr) || var.vt != VT_BSTR)
						Log(LOG_MOREFUNC, _T("error value: 0x%x\n"), hr);

					Log(LOG_MOREFUNC, _T("value: %ld, %s\n"), 0, var.bstrVal);

					hr = disp.GetPropertyByName(OLESTR("constructor"), &var);
					if (FAILED(hr) || var.vt != VT_DISPATCH)
						Log(LOG_MOREFUNC, _T("error constructor: 0x%x\n"), hr);

					hr = disp.GetIDOfName(OLESTR("constructor"), &mydispid);
					if (FAILED(hr) || var.vt != VT_DISPATCH)
						Log(LOG_MOREFUNC, _T("error constructor: 0x%x\n"), hr);

					Log(LOG_MOREFUNC, _T("constructor: %ld, 0x%p\n"), mydispid, var.pdispVal);
					
					hr = disp.GetPropertyByName(OLESTR("apply"), &var);
					if (FAILED(hr) || var.vt != VT_DISPATCH)
						Log(LOG_MOREFUNC, _T("error apply: 0x%x\n"), hr);

					hr = disp.GetIDOfName(OLESTR("apply"), &mydispid);
					if (FAILED(hr) || var.vt != VT_DISPATCH)
						Log(LOG_MOREFUNC, _T("error apply: 0x%x\n"), hr);

					Log(LOG_MOREFUNC, _T("apply: %ld, 0x%p\n"), mydispid, var.pdispVal);

					hr = disp.GetPropertyByName(OLESTR("call"), &var);
					if (FAILED(hr) || var.vt != VT_DISPATCH)
						Log(LOG_MOREFUNC, _T("error call: 0x%x\n"), hr);

					hr = disp.GetIDOfName(OLESTR("call"), &mydispid);
					if (FAILED(hr) || var.vt != VT_DISPATCH)
						Log(LOG_MOREFUNC, _T("error call: 0x%x\n"), hr);

					Log(LOG_MOREFUNC, _T("call: %ld, 0x%p\n"), mydispid, var.pdispVal);

					hr = disp.GetPropertyByName(OLESTR("length"), &var);
					if (FAILED(hr) || var.vt != VT_I4)
						Log(LOG_MOREFUNC, _T("error length: 0x%x\n"), hr);

					hr = disp.GetIDOfName(OLESTR("length"), &mydispid);
					if (FAILED(hr) || var.vt != VT_I4)
						Log(LOG_MOREFUNC, _T("error length: 0x%x\n"), hr);

					Log(LOG_MOREFUNC, _T("length: %ld, %ld\n"), mydispid, var.lVal);

					disp.Detach();
				}
			}
			else
				return DISP_E_PARAMNOTOPTIONAL;
		}
		if (wFlags & DISPATCH_METHOD) {
			if (entry->IsDeleted())
				return DISP_E_MEMBERNOTFOUND;

			CComVariant variant = entry->GetVariant();
			if (V_VT(&variant) != VT_DISPATCH)
				return DISP_E_MEMBERNOTFOUND;
			CComPtr<IDispatch> dispatch = V_DISPATCH(&variant);
			hr = dispatch->Invoke(DISPID_VALUE, IID_NULL, lcid, wFlags, pdp, pvarRes, pei, puArgErr);
			if (FAILED(hr)) return hr;
		}
	} else {
		hr = ::DispInvoke(this, mTypeInfo, id, wFlags, pdp, pvarRes, pei, puArgErr);
	}

	return hr;
}

template <typename T>
STDMETHODIMP IDispatchExImpl<T>::GetIDsOfNamesImpl(LPOLESTR* names, UINT countNames, bool caseSens, DISPID* dispids)
{
	Log(LOG_FUNC, _T("IDispatchExImpl::GetIDsOfNamesImpl(\"%s\", %u) called\n"), *names, countNames);

	HRESULT hr = S_OK;
	int count = countNames;
	CDispatchExEntry* entry = NULL;
	bool found;
	for (int i = 0; i < count; i++) {
		if (caseSens)
			found = mNameToEntryMap.Lookup(names[i], entry);
		else
			found = mNameToEntryIMap.Lookup(names[i], entry);
		if (!found || entry == NULL) {
			hr = DISP_E_UNKNOWNNAME;
			dispids[i] = DISPID_UNKNOWN;
		} else {
			dispids[i] = entry->GetDispID();
		}
	}

	return hr;
}

/*
 * This is a sorted array.
 */
template <typename T>
const LPCOLESTR IDispatchExImpl<T>::kInternalMethods[] =
{
	OLESTR("AddRef"),
	OLESTR("DeleteMemberByDispID"),
	OLESTR("DeleteMemberByName"),
	OLESTR("GetDispID"),
	OLESTR("GetIDsOfNames"),
	OLESTR("GetMemberName"),
	OLESTR("GetMemberProperties"),
	OLESTR("GetNameSpaceParent"),
	OLESTR("GetNextDispID"),
	OLESTR("GetTypeInfo"),
	OLESTR("GetTypeInfoCount"),
	OLESTR("Invoke"),
	OLESTR("QueryInterface"),
	OLESTR("Release"),
	OLESTR("RemoteInvokeEx")
};
