#pragma once
#include <gdiplus.h>
#include <Thumbcache.h>
#include <propsys.h>
#include <atlbase.h>
#include "PIFParser.h"
#include "ImageHelpers.h"

void DllAddRef();
void DllRelease();

/* 70DCDE53-4190-4B4C-8438-3D858EDD45B6 */

const GUID CLSID_PIFShellExtensions = {
    0x70DCDE53, 0x4190, 0x4B4C,
    { 0x84, 0x38, 0x3D, 0x85, 0x8E, 0xDD, 0x45, 0xB6 } };

extern HINSTANCE g_hinst;
extern LONG g_cRef;

/*****************************************************************************
*
*      Portable Image Format Shell Extensions
*
*****************************************************************************/
class CPIFShellExt :
    public IPersistFile,
    public IDataObject,
    public IPropertyStoreFactory,
    public IPropertyStore,
    public IInitializeWithStream,
    public IExtractImage,
    public IThumbnailProvider
{
public:
    CPIFShellExt();

    // IUnknown
    IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ LPVOID* ppvOut);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IPersist
    IFACEMETHODIMP GetClassID(_Out_ CLSID* pclsid);

    // IPersistFile
    IFACEMETHODIMP IsDirty();
    IFACEMETHODIMP Load(LPCOLESTR pszFileName, DWORD dwMode);
    IFACEMETHODIMP Save(LPCOLESTR pszFileName, BOOL fRemember);
    IFACEMETHODIMP SaveCompleted(LPCOLESTR pszFileName);
    IFACEMETHODIMP GetCurFile(_Outptr_ LPOLESTR* ppszFileName);

    // IDataObject
    IFACEMETHODIMP GetData(_In_ FORMATETC *pformatetcIn, _Out_ STGMEDIUM *pmedium);
    IFACEMETHODIMP EnumFormatEtc(_In_ DWORD dwDirFlags, _COM_Outptr_ IEnumFORMATETC** ppief);
    IFACEMETHODIMP QueryGetData(_In_ FORMATETC* pfe);
    IFACEMETHODIMP GetDataHere(FORMATETC*, STGMEDIUM*) { return E_NOTIMPL; }
    IFACEMETHODIMP GetCanonicalFormatEtc(FORMATETC*, FORMATETC*) { return E_NOTIMPL; }
    IFACEMETHODIMP SetData(FORMATETC*, STGMEDIUM*, BOOL) { return E_NOTIMPL; }
    IFACEMETHODIMP DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) { return E_NOTIMPL; }
    IFACEMETHODIMP DUnadvise(DWORD) { return E_NOTIMPL; }
    IFACEMETHODIMP EnumDAdvise(IEnumSTATDATA**) { return E_NOTIMPL; }

    // IPropertyStore
    IFACEMETHODIMP GetCount(_Out_ DWORD* pcProps);
    IFACEMETHODIMP GetAt(_In_ DWORD iProp, _Outptr_ PROPERTYKEY* pkey);
    IFACEMETHODIMP GetValue(_In_ REFPROPERTYKEY key, _Outptr_ PROPVARIANT* pPropVar);
    IFACEMETHODIMP SetValue(_In_ REFPROPERTYKEY key, _In_ REFPROPVARIANT propVar);
    IFACEMETHODIMP Commit();

    // IPropertyStoreFactory
    IFACEMETHODIMP GetPropertyStore(
        _In_ GETPROPERTYSTOREFLAGS flags,
        _In_opt_ IUnknown* punk,
        _In_ REFIID riid,
        _COM_Outptr_ void** ppv);
    IFACEMETHODIMP GetPropertyStoreForKeys(
        _In_reads_opt_(cKeys) const PROPERTYKEY* rgKeys,
        _In_ UINT cKeys,
        _In_ GETPROPERTYSTOREFLAGS flags,
        _In_ REFIID riid,
        _COM_Outptr_ void** ppv);

    // IInitializeWithStream
    IFACEMETHODIMP Initialize(_In_ IStream* pStream, _In_ DWORD grfMode);

    // IExtractImage
    IFACEMETHODIMP GetLocation(
        _Out_writes_(cchMax) PWSTR pszFileName,
        DWORD cchMax,
        _Out_ DWORD* pdwPriority,
        _In_ SIZE const* prgSize,
        DWORD dwRecClrDepth,
        _Inout_ DWORD* pdwFlags);
    IFACEMETHODIMP Extract(_Out_ HBITMAP* phBmpImage);

    // IThumbnailProvider
    IFACEMETHODIMP GetThumbnail(
        _In_ UINT cx,
        _Out_ HBITMAP* phbmp,
        _Out_ WTS_ALPHATYPE* pdwAlpha);

private:
    ~CPIFShellExt();

    HRESULT _EnsurePropertyStore();
    HRESULT _InitializeProperties();

private:
    LONG  _cRef;
    SIZE  _sizeDesired;
    DWORD _dwRecClrDepth;
    DWORD _dwFlags;
    WCHAR _szTarget[MAX_PATH + 1];
    DWORD _grfMode;
    bool  _fPropertiesInitialized;
    CComPtr<IStream> _spStream;
    CComPtr<IPropertyStore> _spPropertyStore;
};

class CPIFFactory : public IClassFactory
{
public:
    CPIFFactory();

    // IUnknown
    IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ LPVOID* ppvOut);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IClassFactory
    IFACEMETHODIMP CreateInstance(_In_ IUnknown* punkOuter, _In_ REFIID riid, _COM_Outptr_ LPVOID* ppvOut);
    IFACEMETHODIMP LockServer(_In_ BOOL fLock);

private:
    ~CPIFFactory();

private:
    LONG _cRef;
};