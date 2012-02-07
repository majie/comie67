/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#pragma once

#include <new>

#include <atlsafe.h>

#include "logger.h"

#ifdef _DEBUG

#define DUMP_IDISPATCH(pdispex)															\
do {																					\
	BSTR bstrName;																		\
	int i = 0;																			\
																						\
	Log(LOG_DUMP, _T("\n===Start dumping an IDispatchEx at 0x%p...\n"), (pdispex));	\
																						\
	DISPID dispid = DISPID_STARTENUM;													\
	while ((pdispex)->GetNextDispID(fdexEnumAll, dispid, &dispid) != S_FALSE) {			\
		if(SUCCEEDED((pdispex)->GetMemberName(dispid, &bstrName))) {					\
		Log(LOG_DUMP, _T("Name[0x%x]: %s\n"), dispid, bstrName);						\
			::SysFreeString(bstrName);													\
		}																				\
		i++;																			\
	}																					\
																						\
	Log(LOG_DUMP, _T("===IDispatchEx dumping ended. No. of DispId: %d\n\n"), i);		\
} while (false)

#define DUMP_VARIANT(pv)																\
do {																					\
	_variant_t variant;																	\
	variant.Attach(*(pv));																\
																						\
	Log(LOG_DUMP, _T("\n===Start dumping an variant with type: %d\n"), variant.vt);	\
	switch (variant.vt) {																\
	case VT_BSTR:																		\
		Log(LOG_DUMP, _T("%s\n"), variant.bstrVal);									\
		break;																			\
	case VT_DISPATCH:																	\
		break;																			\
	default:																			\
		break;																			\
	}																					\
																						\
	variant.Detach();																	\
} while (false)

#define DUMP_VARIANTSAFEARRAY(psa)																			\
do {																										\
	CComSafeArray<VARIANT> sfa;																				\
	sfa.Attach((psa));																						\
																											\
	INT dim = sfa.GetDimensions();																			\
																											\
	Log(LOG_DUMP, _T("===Start dumping a SAFEARRAY: dims: %u, features: 0x%x\n, sizeof Element: %u\n"),	\
		dim, (psa)->fFeatures, (psa)->cbElements);															\
																											\
	LONG* index = new(std::nothrow) LONG[dim];																\
	if (index == NULL) {																					\
		sfa.Detach();																						\
		break;																								\
	}																										\
																											\
	int count = 1;																							\
	for (INT i = 0; i < dim; i++) {																			\
		LONG lower = sfa.GetLowerBound(i), upper = sfa.GetUpperBound(i);									\
		Log(LOG_DUMP, _T("Dimension[%d]: Lower: %ld, Upper: %ld\n"), i, lower, upper);					\
		count *= sfa.GetCount(i);																			\
	}																										\
																											\
	VARIANT* variantArray = static_cast<VARIANT*>(sfa.m_psa->pvData);										\
	_variant_t variant;																						\
	for (int i = 0; i < count; i++) {																		\
		variant.Attach(variantArray[i]);																	\
		/*DUMP_VARIANT(&variant);*/																				\
		variant.Detach();																					\
	}																										\
																											\
	Log(LOG_DUMP, _T("===SAFEARRAY dumping ended.\n\n"));												\
	delete[] index;																							\
	sfa.Detach();																							\
} while (false)

#define DUMP_DISPPARAMS(pdp)																	\
do {																							\
	Log(LOG_DUMP, _T("\n===Start dumping a DISPPARAMS at 0x%p...\n"), (pdp));				\
	int count = (pdp)->cArgs;																	\
	for (int i = 0; i < count; i++) {															\
		DUMP_VARIANT((pdp)->rgvarg[i]);															\
	}																							\
	count = (pdp)->cNamedArgs;																	\
	for (int i = 0; i < count; i++) {															\
		Log(LOG_DUMP, _T("rgdispidNamedArgs[0x%x]: 0x%x "), i, (pdp)->rgdispidNamedArgs[i]);	\
	}																							\
	Log(LOG_DUMP, _T("\n"));																	\
	Log(LOG_DUMP, _T("=== DISPPARAMS dumping ended.\n\n"));									\
} while (false)

#define DUMP_RECT(prect)															\
do {																				\
	Log(LOG_DUMP, _T("RECT: left: %ld, top: %ld, right: %ld, bottom: %ld\n"),	\
		(prect)->left, (prect)->top, (prect)->right, (prect)->bottom);				\
} while (false)

#else // _DEBUG
#define DUMP_IDISPATCH(pdispex) (void)(pdispex)
#define DUMP_VARIANT(pv) (void)(pv)
#define DUMP_VARIANTSAFEARRAY(psa) (void)(psa)
#define DUMP_DISPPARAMS(pdp) (void)(pdp)

#define DUMP_RECT(prect) (void)(prect)
#endif // _DEBUG
