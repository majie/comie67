/* Copyright (c) 2011 The comie67 Authors. All rights reserved.
 * Use of this source code is governed by a BSD-style license
 * that can be found in the LICENSE.txt file.
 */

#include "stdafx.h"
#include "resource.h"
#include "comie67.h"
#include "dlldatax.h"

#include "logger.h"

#ifdef _DEBUG
//#include "Stackwalker.h"
#endif

#if !defined(UNICODE) && !defined(_UNICODE)
#error Only supports Unicode build
#endif

class Ccomie67Module : public CAtlDllModuleT< Ccomie67Module >
{
public :
	DECLARE_LIBID(LIBID_comie67Lib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_COMIE67, "{6E9BF798-AB52-40E3-8791-E011AF5082AC}")
};

#ifdef _DEBUG
/*static struct MemoryLeakCheck {
	MemoryLeakCheck() {InitAllocCheck();}
	~MemoryLeakCheck() {DeInitAllocCheck();}
} gCheck;*/
#endif

Ccomie67Module _AtlModule;


// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
#ifdef _DEBUG
	if (dwReason == DLL_PROCESS_ATTACH) {
		TCHAR logFile[26];
		DWORD pid;

		pid = GetCurrentProcessId();
		_stprintf(logFile, _T("B:\\comie67-%u.log"), pid);

		InitLog(logFile);

		TCHAR moduleFilename[MAX_PATH];
		GetModuleFileName(NULL, moduleFilename, MAX_PATH);
		Log(LOG_FEW, _T("%s\n"), moduleFilename);
    }
    else if (dwReason == DLL_PROCESS_DETACH) {
		UninitLog();
	}
#endif
#ifdef _MERGE_PROXYSTUB
    if (!PrxDllMain(hInstance, dwReason, lpReserved))
        return FALSE;
#endif
	hInstance;
    return _AtlModule.DllMain(dwReason, lpReserved); 
}


// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
#ifdef _MERGE_PROXYSTUB
    HRESULT hr = PrxDllCanUnloadNow();
    if (FAILED(hr))
        return hr;
#endif
    return _AtlModule.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
#ifdef _MERGE_PROXYSTUB
    if (PrxDllGetClassObject(rclsid, riid, ppv) == S_OK)
        return S_OK;
#endif
    return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    HRESULT hr = _AtlModule.DllRegisterServer();
#ifdef _MERGE_PROXYSTUB
    if (FAILED(hr))
        return hr;
    hr = PrxDllRegisterServer();
#endif
	return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _AtlModule.DllUnregisterServer();
#ifdef _MERGE_PROXYSTUB
    if (FAILED(hr))
        return hr;
    hr = PrxDllRegisterServer();
    if (FAILED(hr))
        return hr;
    hr = PrxDllUnregisterServer();
#endif
	return hr;
}
