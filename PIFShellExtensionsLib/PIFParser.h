#pragma once
#include "stdafx.h"

enum PortableImageFormatType
{
    PortableImageFormatType_Invalid = 0,
    PortableImageFormatType_PBMA,
    PortableImageFormatType_PGMA,
    PortableImageFormatType_PPMA,
    PortableImageFormatType_PBMB,
    PortableImageFormatType_PGMB,
    PortableImageFormatType_PPMB,
    PortableImageFormatType_Count
};

class CBaseImageParser
{
public:
    CBaseImageParser();
    virtual ~CBaseImageParser();

    HRESULT InitializeFromStream(_In_ IStream* pStream);
    HRESULT GetBitmap(_Out_ HBITMAP* phBitmap);
    PortableImageFormatType GetImageType() { return m_imageType; }
    UINT GetWidth() { return m_uWidth; }
    UINT GetHeight() { return m_uHeight; }
    virtual UINT GetBitDepth() = 0;

    static HRESULT CreateParserFromStream(_In_ IStream* pStream, _Outptr_ CBaseImageParser** ppImageParser);
    static HRESULT ResizeBitmap(
        _In_ HBITMAP hbmpOriginal,
        _In_ bool fMaintainAspectRatio,
        _In_ SIZE* pSizeDesired,
        _Outptr_ HBITMAP* phbmpResized);

protected:
    virtual size_t GetAllocationSize() = 0;
    virtual HRESULT ReadImageHeaders();
    virtual HRESULT ReadImageData() = 0;
    static HRESULT ReadImageType(_In_ IStream* pStream, _Out_ PortableImageFormatType* pImageType);
    virtual HRESULT ReadBinaryImageData();
    virtual HRESULT ReadAsciiImageData();
    virtual bool IsDataSizeValid() = 0;
    static HRESULT ReadLineFromStream(_In_ IStream* pStream, _Outptr_ char** ppBuffer);
    static HRESULT ParseImageLineForNum(_In_ IStream* pStream, _In_ PCSTR pszFormat, _Out_ UINT* value);
    virtual HRESULT PopulateBitmap(_In_ Gdiplus::Bitmap* pBitmap) = 0;
    static bool IsWhitespace(_In_ char ch);

    UINT m_uWidth;
    UINT m_uHeight;
    UINT m_uMaxValue;
    UINT m_uDataLen;
    BYTE* m_rgData;
    Gdiplus::GdiplusStartupInput _gdiPlusStatupInput;
    ULONG_PTR _gdiPlusToken;
    CComPtr<IStream> m_spStream;
    PortableImageFormatType m_imageType;
};

class CPPMImageParser : public CBaseImageParser
{
public:
    CPPMImageParser(){}
    virtual ~CPPMImageParser(){}

    __override UINT GetBitDepth() { return 24; }

private:
    __override HRESULT PopulateBitmap(_In_ Gdiplus::Bitmap* pBitmap);
    __override HRESULT ReadImageData();
    __override size_t GetAllocationSize();
    __override bool IsDataSizeValid();
};

class CPGMImageParser : public CBaseImageParser
{
public:
    CPGMImageParser(){}
    virtual ~CPGMImageParser(){}

    __override UINT GetBitDepth() { return 8; }

private:
    __override HRESULT PopulateBitmap(_In_ Gdiplus::Bitmap* pBitmap);
    __override HRESULT ReadImageData();
    __override size_t GetAllocationSize();
    __override bool IsDataSizeValid();
};

class CPBMImageParser : public CBaseImageParser
{
public:
    CPBMImageParser(){}
    virtual ~CPBMImageParser(){}

    __override UINT GetBitDepth() { return 1; }

private:
    __override HRESULT PopulateBitmap(_In_ Gdiplus::Bitmap* pBitmap);
    __override HRESULT ReadImageHeaders();
    __override HRESULT ReadImageData();
    __override HRESULT ReadBinaryImageData();
    __override size_t GetAllocationSize();
    __override bool IsDataSizeValid();
};


