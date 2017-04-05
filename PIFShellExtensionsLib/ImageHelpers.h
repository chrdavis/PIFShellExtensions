#pragma once
#include "stdafx.h"

class CPIFImageHelpers
{
public:
    CPIFImageHelpers();
    ~CPIFImageHelpers();

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

    static UINT BitDepthFromPIFType(
        _In_ PortableImageFormatType formatType);
    static HRESULT GetBitmapFromStream(
        _In_ IStream* pStream,
        _Out_ UINT* puWidth,
        _Out_ UINT* puHeight,
        _Out_ HBITMAP* phbmp);
    static HRESULT CreateBitmapFromBytes(
        _In_ BYTE* data,
        _In_ UINT uDataLen,
        _In_ PortableImageFormatType formatType,
        _In_ UINT uWidth,
        _In_ UINT uHeight,
        _Outptr_ HBITMAP *phBitmap);
    static HRESULT ResizeBitmap(
        _In_ HBITMAP hbmpOriginal,
        _In_ bool fMaintainAspectRatio,
        _In_ SIZE* pSizeDesired,
        _Outptr_ HBITMAP* phbmpResized);
    static HRESULT ReadImageHeaders(
        _In_ IStream* pStream,
        _Out_ PortableImageFormatType* pFormatType,
        _Out_ UINT* puWidth,
        _Out_ UINT* puHeight,
        _Out_ UINT* puMax);
    static HRESULT ReadImageFromStream(
        _In_ IStream* pStream,
        _Out_ PortableImageFormatType* pFormatType,
        _Out_ UINT* puWidth,
        _Out_ UINT* puHeight,
        _Out_ UINT* puMax,
        _Out_ UINT* puDataLen,
        _Outptr_ BYTE** ppData);
    static HRESULT ReadLineFromStream(
        _In_ IStream* pStream,
        _Outptr_ char** ppBuffer);
    static HRESULT ParseImageLineForNum(
        _In_ IStream* pStream,
        _In_ PCSTR format,
        _Out_ UINT* value);
    static HRESULT PopulateBitmapFromBytesPPM(
        _In_ BYTE* data,
        _In_ UINT uDataLen,
        _In_ UINT uWidth,
        _In_ UINT uHeight,
        _In_ Gdiplus::Bitmap* pBitmap);
    static HRESULT PopulateBitmapFromBytesPGM(
        _In_ BYTE* data,
        _In_ UINT uDataLen,
        _In_ UINT uWidth,
        _In_ UINT uHeight,
        _In_ Gdiplus::Bitmap* pBitmap);
    static HRESULT PopulateBitmapFromBytesPBM(
        _In_ BYTE* data,
        _In_ UINT uDataLen,
        _In_ UINT uWidth,
        _In_ UINT uHeight,
        _In_ Gdiplus::Bitmap* pBitmap);

    Gdiplus::GdiplusStartupInput _gdiPlusStatupInput;
    ULONG_PTR _gdiPlusToken;
};