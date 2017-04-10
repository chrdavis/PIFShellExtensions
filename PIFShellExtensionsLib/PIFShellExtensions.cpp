#include "stdafx.h"
#include <propkey.h>
#include "PIFShellExtensions.h"

static HINSTANCE g_hinst = 0;
static LONG g_cRef = 0;

#ifndef BOOLIFY
#define BOOLIFY(expr) (!!(expr))
#endif

void DllAddRef(void)
{
    InterlockedIncrement(&g_cRef);
}

void DllRelease(void)
{
    InterlockedDecrement(&g_cRef);
}

CPIFShellExt::CPIFShellExt() :
    _cRef(1),
    _dwRecClrDepth(0),
    _dwFlags(0),
    _grfMode(0),
    _fPropertiesInitialized(false)
{
    _szTarget[0] = L'\0';
    _sizeDesired.cx = 0;
    _sizeDesired.cy = 0;
    DllAddRef();
}

CPIFShellExt::~CPIFShellExt()
{
    DllRelease();
}

IFACEMETHODIMP CPIFShellExt::QueryInterface(_In_ REFIID riid, _COM_Outptr_ LPVOID *ppvOut)
{
    static const QITAB qit[] = {
        QITABENT(CPIFShellExt, IPersist),
        QITABENT(CPIFShellExt, IPersistFile),
        QITABENT(CPIFShellExt, IDataObject),
        QITABENT(CPIFShellExt, IPropertyStoreFactory),
        QITABENT(CPIFShellExt, IPropertyStore),
        QITABENT(CPIFShellExt, IInitializeWithStream),
        QITABENT(CPIFShellExt, IExtractImage),
        QITABENT(CPIFShellExt, IThumbnailProvider),
        { 0 },
    };
    return QISearch(this, qit, riid, ppvOut);
}

IFACEMETHODIMP_(ULONG) CPIFShellExt::AddRef()
{
    return InterlockedIncrement(&_cRef);
}

IFACEMETHODIMP_(ULONG) CPIFShellExt::Release()
{
    ULONG cRef = InterlockedDecrement(&_cRef);
    if (cRef == 0)
    {
        delete this;
    }
    return cRef;
}

// IPersist
IFACEMETHODIMP CPIFShellExt::GetClassID(_Out_ CLSID* pclsid)
{
    *pclsid = CLSID_PIFShellExtensions;
    return S_OK;
}

IFACEMETHODIMP CPIFShellExt::IsDirty()
{
    // never dirty since we never modify the file
    return S_FALSE;
}

IFACEMETHODIMP CPIFShellExt::Load(LPCOLESTR pszFileName, DWORD)
{
    // Remember the filename so we can render from it later
    HRESULT hr = StringCchCopy(_szTarget, ARRAYSIZE(_szTarget), pszFileName);
    if (SUCCEEDED(hr))
    {
        _spStream = nullptr;
        hr = SHCreateStreamOnFileEx(pszFileName, STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &_spStream);
    }
    return hr;
}

IFACEMETHODIMP CPIFShellExt::Save(LPCOLESTR, BOOL)
{
    // We are just an extractor - we can't save anything
    return E_NOTIMPL;
}

IFACEMETHODIMP CPIFShellExt::SaveCompleted(LPCOLESTR)
{
    // We are just an extractor - we can't save anything
    return E_NOTIMPL;
}

IFACEMETHODIMP CPIFShellExt::GetCurFile(_Outptr_ LPOLESTR* ppszFileName)
{
    *ppszFileName = nullptr;
    HRESULT hr = E_FAIL;

    if (_szTarget[0])
    {
        hr = SHStrDup(_szTarget, ppszFileName);
    }
    return hr;
}

// IDataObject
IFACEMETHODIMP CPIFShellExt::EnumFormatEtc(_In_ DWORD, _COM_Outptr_ IEnumFORMATETC** ppiefe)
{
    *ppiefe = nullptr;
    FORMATETC rgfmtetc[] =
    {
        { CF_BITMAP, nullptr, DVASPECT_CONTENT, -1, TYMED_GDI },
    };
    return SHCreateStdEnumFmtEtc(ARRAYSIZE(rgfmtetc), rgfmtetc, ppiefe);
}

IFACEMETHODIMP CPIFShellExt::QueryGetData(_In_ FORMATETC* pfe)
{
    HRESULT hr = pfe ? S_OK : E_POINTER;
    if (SUCCEEDED(hr))
    {
        if (!(DVASPECT_CONTENT & pfe->dwAspect))
        {
            hr = DATA_E_FORMATETC;
        }
        else
        {
            if ((pfe->cfFormat != CF_BITMAP) &&
                !(pfe->tymed & TYMED_GDI))
            {
                hr = DATA_E_FORMATETC;
            }
        }
    }

    return hr;
}

IFACEMETHODIMP CPIFShellExt::GetData(_In_ FORMATETC* pformatetcIn, _Out_ STGMEDIUM* pmedium)
{
    HRESULT hr = DV_E_FORMATETC;
    if ((pformatetcIn->tymed & TYMED_GDI) && (pformatetcIn->cfFormat & CF_BITMAP))
    {
        hr = _spStream ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            CBaseImageParser* pImageParser = nullptr;
            hr = CBaseImageParser::CreateParserFromStream(_spStream, &pImageParser);
            if (SUCCEEDED(hr))
            {
                HBITMAP hbmp;
                hr = pImageParser->GetBitmap(&hbmp);
                if (SUCCEEDED(hr))
                {
                    pmedium->tymed = TYMED_GDI;
                    pmedium->hBitmap = (HBITMAP)OleDuplicateData(hbmp, CF_BITMAP, 0);
                    pmedium->pUnkForRelease = nullptr;
                    hr = pmedium->hBitmap != nullptr ? S_OK : STG_E_MEDIUMFULL;
                    DeleteObject(hbmp);
                }

                delete pImageParser;
            }
        }
    }

    return hr;
}

// IPropertyStore
IFACEMETHODIMP CPIFShellExt::GetCount(_Out_ DWORD* pcProps)
{
    HRESULT hr = _EnsurePropertyStore();
    if (SUCCEEDED(hr))
    {
        hr = _spPropertyStore->GetCount(pcProps);
    }
    return hr;
}

IFACEMETHODIMP CPIFShellExt::GetAt(_In_ DWORD iProp, _Outptr_ PROPERTYKEY* pkey)
{
    HRESULT hr = _EnsurePropertyStore();
    if (SUCCEEDED(hr))
    {
        hr = _spPropertyStore->GetAt(iProp, pkey);
    }
    return hr;
}

IFACEMETHODIMP CPIFShellExt::GetValue(_In_ REFPROPERTYKEY key, _Outptr_ PROPVARIANT* pPropVar)
{
    HRESULT hr = _EnsurePropertyStore();
    if (SUCCEEDED(hr))
    {
        hr = _spPropertyStore->GetValue(key, pPropVar);
    }
    return hr;
}

IFACEMETHODIMP CPIFShellExt::SetValue(_In_ REFPROPERTYKEY key, _In_ REFPROPVARIANT propVar)
{
    HRESULT hr = _EnsurePropertyStore();
    if (SUCCEEDED(hr))
    {
        hr = _spPropertyStore->SetValue(key, propVar);
    }
    return hr;
}

IFACEMETHODIMP CPIFShellExt::Commit()
{
    HRESULT hr = _EnsurePropertyStore();
    if (SUCCEEDED(hr))
    {
        hr = _spPropertyStore->Commit();
    }
    return hr;
}

// IPropertyStoreFactory
IFACEMETHODIMP CPIFShellExt::GetPropertyStore(
    _In_ GETPROPERTYSTOREFLAGS flags,
    _In_opt_ IUnknown* punk,
    _In_ REFIID riid,
    _COM_Outptr_ void **ppv)
{
    *ppv = nullptr;
    HRESULT hr = _EnsurePropertyStore();
    if (SUCCEEDED(hr))
    {
        CComPtr<IPropertyStoreFactory> spPropStoreFactory;
        hr = _spPropertyStore->QueryInterface(IID_PPV_ARGS(&spPropStoreFactory));
        if (SUCCEEDED(hr))
        {
            hr = spPropStoreFactory->GetPropertyStore(flags, punk, riid, ppv);
        }
    }
    return hr;
}

IFACEMETHODIMP CPIFShellExt::GetPropertyStoreForKeys(
    _In_reads_opt_(cKeys) const PROPERTYKEY* rgKeys,
    _In_ UINT cKeys,
    _In_ GETPROPERTYSTOREFLAGS flags,
    _In_ REFIID riid,
    _COM_Outptr_ void** ppv)
{
    *ppv = nullptr;
    HRESULT hr = _EnsurePropertyStore();
    if (SUCCEEDED(hr))
    {
        CComPtr<IPropertyStoreFactory> spPropStoreFactory;
        hr = _spPropertyStore->QueryInterface(IID_PPV_ARGS(&spPropStoreFactory));
        if (SUCCEEDED(hr))
        {
            hr = spPropStoreFactory->GetPropertyStoreForKeys(rgKeys, cKeys, flags, riid, ppv);
        }
    }
    return hr;
}


// IInitializeWithStream
IFACEMETHODIMP CPIFShellExt::Initialize(_In_ IStream* pStream, _In_ DWORD grfMode)
{
    _spStream = pStream;
    _grfMode = grfMode;
    return S_OK;
}

// IExtractImage
IFACEMETHODIMP CPIFShellExt::GetLocation(
    _Out_writes_(cchMax) PWSTR pszFileName,
    DWORD cchMax,
    _Out_ DWORD* pdwPriority,
    _In_ SIZE const* prgSize,
    DWORD dwRecClrDepth,
    _Inout_ DWORD* pdwFlags)
{
    HRESULT hr = E_FAIL;
    if (_szTarget[0])
    {
        *pdwPriority = IEIT_PRIORITY_NORMAL;

        /* Remember extraction parameters */
        _sizeDesired = *prgSize;
        _dwRecClrDepth = dwRecClrDepth;
        _dwFlags = *pdwFlags;

        /* Return results to caller */

        /* Tell caller which file we are about to extract */
        hr = StringCchCopy(pszFileName, cchMax, _szTarget);
        if (SUCCEEDED(hr))
        {
            /*
            * Tell caller that we aren't going to cache the result;
            * the caller will have to cache the result for us.
            */
            *pdwFlags |= IEIFLAG_CACHE;

            if (_dwFlags & IEIFLAG_ASYNC)
            {
                /* Tell caller to continue on a background thread */
                hr = E_PENDING;
            }
            else
            {
                /* Tell caller everything is safe to extract now */
                hr = S_OK;
            }
        }
    }

    return hr;
}

IFACEMETHODIMP CPIFShellExt::Extract(_Out_ HBITMAP* phBmpImage)
{
    *phBmpImage = nullptr;
    HRESULT hr = S_OK;

    // Make sure the file exists
    DWORD dwAttributes = GetFileAttributes(_szTarget);
    if (dwAttributes != 0xFFFFFFFF)
    {
        // If the file has been archived to tape, don't render it
        // unless high quality is explicitly requested.
        if ((_dwFlags & IEIFLAG_QUALITY) || !(dwAttributes & FILE_ATTRIBUTE_OFFLINE))
        {
            hr = _spStream ? S_OK : E_FAIL;
            if (SUCCEEDED(hr))
            {
                bool fMaintainAspectRatio = false;
                SIZE sizeBitmap = { 0 };
                if (_dwFlags & IEIFLAG_ASPECT)
                {
                    fMaintainAspectRatio = true;
                    sizeBitmap.cx = _sizeDesired.cx;
                    sizeBitmap.cy = _sizeDesired.cy;
                }
                else
                {
                    // Can use our preferred aspect ratio, as long as it is no larger
                    // than the requested size.  We use the largest square that fits.
                    sizeBitmap.cx = sizeBitmap.cy = min(_sizeDesired.cx, _sizeDesired.cy);
                }

                // Get the bitmap and resize it appropriately
                CBaseImageParser* pImageParser = nullptr;
                hr = CBaseImageParser::CreateParserFromStream(_spStream, &pImageParser);
                if (SUCCEEDED(hr))
                {
                    HBITMAP hbmpOriginal;
                    hr = pImageParser->GetBitmap(&hbmpOriginal);
                    if (SUCCEEDED(hr))
                    {
                        hr = pImageParser->ResizeBitmap(hbmpOriginal, fMaintainAspectRatio, &_sizeDesired, phBmpImage);
                        DeleteObject(hbmpOriginal);
                    }

                    delete pImageParser;
                    pImageParser = nullptr;
                }
            }
        }
        else
        {
            // Don't recall a file from tape just to render its thumbnail
            hr = E_FAIL;
        }
    }
    else
    {
        // Couldn't get file attributes
        hr = HRESULT_FROM_WIN32(GetLastError());
    }

    return hr;
}

// IThumbnailProvider
IFACEMETHODIMP CPIFShellExt::GetThumbnail(
    _In_ UINT cx,
    _Out_ HBITMAP* phbmp,
    _Out_ WTS_ALPHATYPE* pdwAlpha)
{
    *phbmp = nullptr;
    *pdwAlpha = WTSAT_RGB;

    HRESULT hr = _spStream ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        // Get the bitmap and resize it appropriately
        CBaseImageParser* pImageParser = nullptr;
        hr = CBaseImageParser::CreateParserFromStream(_spStream, &pImageParser);
        if (SUCCEEDED(hr))
        {
            HBITMAP hbmpOriginalSize;
            hr = pImageParser->GetBitmap(&hbmpOriginalSize);
            if (SUCCEEDED(hr))
            {
                _sizeDesired.cx = cx;
                _sizeDesired.cy = cx;
                hr = pImageParser->ResizeBitmap(hbmpOriginalSize, true, &_sizeDesired, phbmp);
                DeleteObject(hbmpOriginalSize);
            }

            delete pImageParser;
            pImageParser = nullptr;
        }
    }
    return hr;
}

HRESULT CPIFShellExt::_EnsurePropertyStore()
{
    HRESULT hr = S_OK;
    if (_spPropertyStore == nullptr)
    {
        hr = _spStream ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            hr = CoCreateInstance(CLSID_InMemoryPropertyStore, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_spPropertyStore));
        }
    }

    if (SUCCEEDED(hr) && _spStream && !_fPropertiesInitialized)
    {
        hr = _InitializeProperties();
    }

    return hr;
}

HRESULT CPIFShellExt::_InitializeProperties()
{
    HRESULT hr = _spPropertyStore ? S_OK : E_FAIL;
    if (SUCCEEDED(hr) && !_fPropertiesInitialized)
    {
        PROPVARIANT pv = {};
        CBaseImageParser* pImageParser = nullptr;
        hr = CBaseImageParser::CreateParserFromStream(_spStream, &pImageParser);

        if (SUCCEEDED(hr))
        {
            pv.vt = VT_LPWSTR;
            pv.pwszVal = L"image/bmp";
            hr = _spPropertyStore->SetValue(PKEY_MIMEType, pv);
        }

        if (SUCCEEDED(hr))
        {
            pv.vt = VT_UI4;
            pv.ulVal = pImageParser->GetWidth();
            hr = _spPropertyStore->SetValue(PKEY_Image_HorizontalSize, pv);
        }

        if (SUCCEEDED(hr))
        {
            pv.vt = VT_UI4;
            pv.ulVal = pImageParser->GetHeight();
            hr = _spPropertyStore->SetValue(PKEY_Image_VerticalSize, pv);
        }

        if (SUCCEEDED(hr))
        {
            pv.vt = VT_UI4;
            pv.ulVal = pImageParser->GetBitDepth();
            hr = _spPropertyStore->SetValue(PKEY_Image_BitDepth, pv);
        }

        if (SUCCEEDED(hr))
        {
            wchar_t szResolution[30];
            hr = StringCchPrintf(szResolution, ARRAYSIZE(szResolution), L"%i x %i", pImageParser->GetWidth(), pImageParser->GetHeight());
            pv.vt = VT_LPWSTR;
            pv.pwszVal = szResolution;
            hr = _spPropertyStore->SetValue(PKEY_Image_Dimensions, pv);
        }

        if (SUCCEEDED(hr))
        {
            _fPropertiesInitialized = true;
        }

        delete pImageParser;
        pImageParser = nullptr;

        IStream_Reset(_spStream);
    }
    return hr;
}

// Class factories do not count towards the DLL reference count, so we
// do not call DllAddRef/DllRelease.

CPIFFactory::CPIFFactory() : _cRef(1)
{
}

CPIFFactory::~CPIFFactory()
{
}

IFACEMETHODIMP CPIFFactory::QueryInterface(_In_ REFIID riid, _COM_Outptr_ LPVOID* ppvOut)
{
    static const QITAB qit[] = {
        QITABENT(CPIFFactory, IClassFactory),
        { 0 },
    };
    return QISearch(this, qit, riid, ppvOut);
}

IFACEMETHODIMP_(ULONG) CPIFFactory::AddRef()
{
    return InterlockedIncrement(&_cRef);
}

IFACEMETHODIMP_(ULONG) CPIFFactory::Release()
{
    ULONG cRef = InterlockedDecrement(&_cRef);
    if (cRef == 0)
    {
        delete this;
    }
    return cRef;
}

IFACEMETHODIMP CPIFFactory::LockServer(_In_ BOOL fLock)
{
    if (fLock)
    {
        DllAddRef();
    }
    else
    {
        DllRelease();
    }
    return S_OK;
}

HRESULT CPIFFactory::CreateInstance(_In_ IUnknown* punkOuter, _In_ REFIID riid, _COM_Outptr_ LPVOID* ppvOut)
{
    HRESULT hr;
    *ppvOut = nullptr;

    if (punkOuter == nullptr)
    {
        CPIFShellExt *pobj = new CPIFShellExt();
        if (pobj)
        {
            hr = pobj->QueryInterface(riid, ppvOut);
            pobj->Release();
        }
        else
        {
            hr = E_OUTOFMEMORY;
        }
    }
    else
    {
        hr = CLASS_E_NOAGGREGATION;
    }
    return hr;
}
