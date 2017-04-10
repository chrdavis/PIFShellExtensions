#pragma once
#include "stdafx.h"

class CPIFImageHelpers
{
public:
    CPIFImageHelpers();
    ~CPIFImageHelpers();

    static HRESULT GetBitmapFromStream(
        _In_ IStream* pStream,
        _Out_ HBITMAP* phbmp);
    static HRESULT ResizeBitmap(
        _In_ HBITMAP hbmpOriginal,
        _In_ bool fMaintainAspectRatio,
        _In_ SIZE* pSizeDesired,
        _Outptr_ HBITMAP* phbmpResized);

    Gdiplus::GdiplusStartupInput _gdiPlusStatupInput;
    ULONG_PTR _gdiPlusToken;
};