#include "stdafx.h"
#include <PIFShellExtensions.h>

LONG g_cRef;
HINSTANCE g_hinst;

STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID FAR* ppv)
{
    HRESULT hr;
    *ppv = nullptr;

    if (rclsid == CLSID_PIFShellExtensions)
    {
        CPIFFactory* pfact = new CPIFFactory();
        if (pfact)
        {
            hr = pfact->QueryInterface(riid, ppv);
            pfact->Release();
        }
        else
        {
            hr = E_OUTOFMEMORY;
        }
    }
    else
    {
        hr = CLASS_E_CLASSNOTAVAILABLE;
    }
    return hr;
}

STDAPI DllCanUnloadNow(void)
{
    return g_cRef ? S_FALSE : S_OK;
}

STDAPI DllUnregisterServer()
{
    return S_OK;
}

STDAPI DllRegisterServer()
{
    return S_OK;
}

STDAPI_(BOOL) DllMain(HINSTANCE hinst, DWORD dwReason, LPVOID)
{
    switch (dwReason)
    {
    case DLL_PROCESS_ATTACH:
        g_hinst = hinst;
        DisableThreadLibraryCalls(g_hinst);
        break;
    }
    return 1;
}