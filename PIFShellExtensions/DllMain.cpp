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

//
// Registry keys are removed here.
//
STDAPI DllUnregisterServer()
{
    // TODO: Do this better
    // Delete the keys we previously created.
    PCWSTR rgRegPaths[] = 
    {
        L".pbm\\shellex\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}",
        L".pgm\\shellex\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}",
        L".pnm\\shellex\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}",
        L".ppm\\shellex\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}",
        L".pbm\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}",
        L".pgm\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}",
        L".pnm\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}",
        L".ppm\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}",
        L".pbm\\shellex\\PropertyHandler",
        L".pgm\\shellex\\PropertyHandler",
        L".pnm\\shellex\\PropertyHandler",
        L".ppm\\shellex\\PropertyHandler",
        L"CLSID\\{70DCDE53-4190-4B4C-8438-3D858EDD45B6}",
        L"PBMFile.1"
        L"PGMFile.1"
        L"PPMFile.1"
        L"PNMFile.1"
    };

    HRESULT hr = S_OK;
    for (UINT u = 0; u < ARRAYSIZE(rgRegPaths); u++)
    {
        if (FAILED(HRESULT_FROM_WIN32(SHDeleteKey(HKEY_CLASSES_ROOT, rgRegPaths[u]))))
        {
            hr = SELFREG_E_CLASS;
        }
    }

    return hr;
}

// Structure to hold data for individual keys to register.
typedef struct
{
    HKEY  hRootKey;
    PCWSTR pszSubKey;
    PCWSTR pszValueName;
    PCWSTR pszData;
    DWORD dwType;
} REGSTRUCT;

STDAPI DllRegisterServer()
{
    // TODO: update to use szClassID or a #define
    //WCHAR szClassID[64] = {};
    WCHAR szData[MAX_PATH] = {};
    //StringFromGUID2(CLSID_PPMImageExtractor, szClassID, ARRAYSIZE(szClassID));      // Convert the IID to a string.

    // TODO: add entries under 
    // TODO: HKLM\SystemFileAssociations\.ext
    // TODO: HKLK\Software\Microsoft|windows\CurrentVersion\PropertySystem\PropertyHandlers\.ext (Default) <CLSID>

    REGSTRUCT rgRegEntries[] =
    {
        HKEY_CLASSES_ROOT, L".pbm",                                                          nullptr,           L"PBMFile.1",                              REG_SZ,
        HKEY_CLASSES_ROOT, L".pgm",                                                          nullptr,           L"PGMFile.1",                              REG_SZ,
        HKEY_CLASSES_ROOT, L".ppm",                                                          nullptr,           L"PPMFile.1",                              REG_SZ,
        HKEY_CLASSES_ROOT, L".pnm",                                                          nullptr,           L"PNMFile.1",                              REG_SZ,
        HKEY_CLASSES_ROOT, L".pbm",                                                          L"ContentType",    L"Image/bmp"   ,                           REG_SZ,
        HKEY_CLASSES_ROOT, L".pgm",                                                          L"ContentType",    L"Image/bmp",                              REG_SZ,
        HKEY_CLASSES_ROOT, L".ppm",                                                          L"ContentType",    L"Image/bmp",                              REG_SZ,
        HKEY_CLASSES_ROOT, L".pnm",                                                          L"ContentType",    L"Image/bmp",                              REG_SZ,
        HKEY_CLASSES_ROOT, L".pbm",                                                          L"PerceivedType",  L"Image",                                  REG_SZ,
        HKEY_CLASSES_ROOT, L".pgm",                                                          L"PerceivedType",  L"Image",                                  REG_SZ,
        HKEY_CLASSES_ROOT, L".ppm",                                                          L"PerceivedType",  L"Image",                                  REG_SZ,
        HKEY_CLASSES_ROOT, L".pnm",                                                          L"PerceivedType",  L"Image",                                  REG_SZ,
        HKEY_CLASSES_ROOT, L".pbm\\shellex\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}",         nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".pgm\\shellex\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}",         nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".pnm\\shellex\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}",         nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".ppm\\shellex\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}",         nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".pbm\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}",         nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".pgm\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}",         nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".pnm\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}",         nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".ppm\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}",         nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".pbm\\shellex\\PropertyHandler",                                nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".pgm\\shellex\\PropertyHandler",                                nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".pnm\\shellex\\PropertyHandler",                                nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L".ppm\\shellex\\PropertyHandler",                                nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L"CLSID\\{70DCDE53-4190-4B4C-8438-3D858EDD45B6}",                 nullptr,           nullptr,                                   REG_SZ,
        HKEY_CLASSES_ROOT, L"CLSID\\{70DCDE53-4190-4B4C-8438-3D858EDD45B6}\\InprocServer32", nullptr,           L"PIFShellExtensions.dll",                 REG_SZ,
        HKEY_CLASSES_ROOT, L"CLSID\\{70DCDE53-4190-4B4C-8438-3D858EDD45B6}\\InprocServer32", L"ThreadingModel", L"Both",                                   REG_SZ,
        HKEY_CLASSES_ROOT, L"PBMFile.1",                                                     nullptr,           L"PBM Image",                              REG_SZ,
        HKEY_CLASSES_ROOT, L"PBMFile.1\\shellex\\DataHandler",                               nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L"PGMFile.1",                                                     nullptr,           L"PGM Image",                              REG_SZ,
        HKEY_CLASSES_ROOT, L"PGMFile.1\\shellex\\DataHandler",                               nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L"PPMFile.1",                                                     nullptr,           L"PPM Image",                              REG_SZ,
        HKEY_CLASSES_ROOT, L"PPMFile.1\\shellex\\DataHandler",                               nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
        HKEY_CLASSES_ROOT, L"PNMFile.1",                                                     nullptr,           L"PNM Image",                              REG_SZ,
        HKEY_CLASSES_ROOT, L"PNMFile.1\\shellex\\DataHandler",                               nullptr,           L"{70DCDE53-4190-4B4C-8438-3D858EDD45B6}", REG_SZ,
    };

    HRESULT hr = S_OK;

    for (int i = 0; SUCCEEDED(hr) && (i < ARRAYSIZE(rgRegEntries)); i++)
    {
        HKEY hKey = nullptr;
        LONG lResult = RegCreateKeyEx(
            rgRegEntries[i].hRootKey,
            rgRegEntries[i].pszSubKey,
            0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr);
        if (S_OK == lResult)
        {
            if (rgRegEntries[i].pszData)
            {
                hr = StringCchCopy(szData, ARRAYSIZE(szData), rgRegEntries[i].pszData);
                if (SUCCEEDED(hr))
                {
                    RegSetValueEx(
                        hKey,
                        rgRegEntries[i].pszValueName,
                        0,
                        rgRegEntries[i].dwType,
                        (LPBYTE)szData,
                        (lstrlen(szData) + 1) * sizeof(WCHAR));
                }
            }

            RegCloseKey(hKey);
        }
        else
        {
            hr = SELFREG_E_CLASS;
        }
    }


    if (FAILED(hr))
    {
        // Remove the stuff we added.
        DllUnregisterServer();
    }

    return hr;
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