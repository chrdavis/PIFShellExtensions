#include "stdafx.h"
#include "ImageHelpers.h"
#include <gdiplus.h>

using namespace Gdiplus;

const LARGE_INTEGER c_liZero = { 0, 0 };

CPIFImageHelpers::CPIFImageHelpers()
{
    GdiplusStartupInput gdiPlusStatupInput;
    GdiplusStartup(&_gdiPlusToken, &gdiPlusStatupInput, nullptr);
}

CPIFImageHelpers::~CPIFImageHelpers()
{
    GdiplusShutdown(_gdiPlusToken);
}

UINT CPIFImageHelpers::BitDepthFromPIFType(_In_ PortableImageFormatType formatType)
{
    UINT uBitDepth = 0;
    switch (formatType)
    {
    case PortableImageFormatType_PBMA:
    case PortableImageFormatType_PBMB:
        uBitDepth = 1;
        break;
    case PortableImageFormatType_PGMA:
    case PortableImageFormatType_PGMB:
        uBitDepth = 8;
        break;
    case PortableImageFormatType_PPMA:
    case PortableImageFormatType_PPMB:
        uBitDepth = 24;
        break;
    }
    return uBitDepth;
}

HRESULT CPIFImageHelpers::GetBitmapFromStream(
    _In_ IStream* pStream,
    _Out_ UINT* puWidth,
    _Out_ UINT* puHeight,
    _Out_ HBITMAP* phbmp)
{
    *phbmp = nullptr;
    UINT uMax = 0;
    PortableImageFormatType formatType = PortableImageFormatType_Invalid;
    BYTE* pb = nullptr;
    UINT uDataLen = 0;

    HRESULT hr = ReadImageFromStream(pStream, &formatType, puWidth, puHeight, &uMax, &uDataLen, &pb);
    if (SUCCEEDED(hr))
    {
        hr = CreateBitmapFromBytes(pb, uDataLen, formatType, *puWidth, *puHeight, phbmp);
        delete[] pb;
    }

    return hr;
}

HRESULT CPIFImageHelpers::ResizeBitmap(
    _In_ HBITMAP hbmpOriginal,
    _In_ bool fMaintainAspectRatio,
    _In_ SIZE* pSizeDesired,
    _Outptr_ HBITMAP* phbmpResized)
{
    *phbmpResized = nullptr;

    Bitmap *pBitmap = new Bitmap(hbmpOriginal, nullptr);
    HRESULT hr = pBitmap ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        int sourceWidth = pBitmap->GetWidth();
        int sourceHeight = pBitmap->GetHeight();

        hr = E_FAIL;
        if (sourceWidth > 0 && sourceHeight > 0 && pSizeDesired->cx > 0 && pSizeDesired->cy > 0)
        {
            int desiredWidth = pSizeDesired->cx;
            int desiredHeight = pSizeDesired->cy;

            if (fMaintainAspectRatio)
            {
                // Maintain the aspect ratio of the original image
                float aspect = 0;
                if (sourceHeight > sourceWidth)
                {
                    aspect = (float)sourceWidth / (float)sourceHeight;
                    desiredWidth = min((int)((float)pSizeDesired->cx * aspect), sourceWidth);
                    desiredHeight = min(pSizeDesired->cy, sourceHeight);
                }
                else
                {
                    aspect = (float)sourceHeight / (float)sourceWidth;
                    desiredWidth = min(pSizeDesired->cx, sourceWidth);
                    desiredHeight = min((int)((float)pSizeDesired->cy * aspect), sourceHeight);
                }
            }

            Bitmap *pResizeBitmap = new Bitmap(desiredWidth, desiredHeight, PixelFormat24bppRGB);
            hr = pResizeBitmap ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                pResizeBitmap->SetResolution(pBitmap->GetHorizontalResolution(), pBitmap->GetVerticalResolution());

                Graphics *pgrResize = new Graphics(pResizeBitmap);
                hr = pgrResize ? S_OK : E_OUTOFMEMORY;
                if (SUCCEEDED(hr))
                {
                    pgrResize->Clear(Color(0, 0, 0));
                    pgrResize->SetInterpolationMode(InterpolationModeBicubic);

                    RectF destinationRect((float)0, (float)0, (float)desiredWidth, (float)desiredHeight);

                    pgrResize->DrawImage((Image*)pBitmap, destinationRect);

                    delete pgrResize;
                }

                // Create the HBITMAP
                if (pResizeBitmap->GetHBITMAP(Color(0, 0, 0), phbmpResized) != Gdiplus::Ok)
                {
                    hr = E_FAIL;
                }

                delete pResizeBitmap;
            }
        }

        delete pBitmap;
    }

    return hr;
}

HRESULT CPIFImageHelpers::PopulateBitmapFromBytesPPM(
    _In_ BYTE* data,
    _In_ UINT uDataLen,
    _In_ UINT uWidth,
    _In_ UINT uHeight,
    _In_ Bitmap* pBitmap)
{
    HRESULT hr = (uWidth * uHeight * 3 == uDataLen) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        BYTE *pCurr = data;
        for (UINT row = 0; row < uHeight; row++)
        {
            for (UINT col = 0; col < uWidth; col++)
            {
                // Color RGB values
                pBitmap->SetPixel(col, row, Color(pCurr[0], pCurr[1], pCurr[2]));
                pCurr += 3;
            }
        }
    }

    return hr;
}

HRESULT CPIFImageHelpers::PopulateBitmapFromBytesPGM(
    _In_ BYTE* data,
    _In_ UINT uDataLen,
    _In_ UINT uWidth,
    _In_ UINT uHeight,
    _In_ Bitmap* pBitmap)
{
    HRESULT hr = (uWidth * uHeight == uDataLen) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        BYTE *pCurr = data;
        for (UINT row = 0; row < uHeight; row++)
        {
            for (UINT col = 0; col < uWidth; col++)
            {
                // Grayscale values
                pBitmap->SetPixel(col, row, Color(pCurr[0], pCurr[0], pCurr[0]));
                pCurr += 1;
            }
        }
    }

    return hr;
}

HRESULT CPIFImageHelpers::PopulateBitmapFromBytesPBM(
    _In_ BYTE* data,
    _In_ UINT uDataLen,
    _In_ UINT uWidth,
    _In_ UINT uHeight,
    _In_ Bitmap* pBitmap)
{
    HRESULT hr = (uWidth * uHeight == uDataLen) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        BYTE *pCurr = data;
        for (UINT row = 0; row < uHeight; row++)
        {
            for (UINT col = 0; col < uWidth; col++)
            {
                // In this format 0=white and 1=black.
                BYTE curr = (pCurr[0] == 0 ? 255 : 0);
                pBitmap->SetPixel(col, row, Color(curr, curr, curr));
                pCurr += 1;
            }
        }
    }

    return hr;
}

HRESULT CPIFImageHelpers::CreateBitmapFromBytes(
    _In_ BYTE* data,
    _In_ UINT uDataLen,
    _In_ PortableImageFormatType formatType,
    _In_ UINT uWidth,
    _In_ UINT uHeight,
    _Outptr_ HBITMAP* phBitmap)
{
    *phBitmap = nullptr;

    Bitmap *pBitmap = new Bitmap(uWidth, uHeight, PixelFormat24bppRGB);
    HRESULT hr = pBitmap ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        if ((formatType == PortableImageFormatType_PPMB) ||
            (formatType == PortableImageFormatType_PPMA))
        {
            hr = PopulateBitmapFromBytesPPM(data, uDataLen, uWidth, uHeight, pBitmap);
        }
        else if ((formatType == PortableImageFormatType_PGMB) ||
            (formatType == PortableImageFormatType_PGMA))
        {
            hr = PopulateBitmapFromBytesPGM(data, uDataLen, uWidth, uHeight, pBitmap);
        }
        else if ((formatType == PortableImageFormatType_PBMA) ||
            (formatType == PortableImageFormatType_PBMB))
        {
            hr = PopulateBitmapFromBytesPBM(data, uDataLen, uWidth, uHeight, pBitmap);
        }
        else
        {
            hr = E_INVALIDARG;
        }

        if (SUCCEEDED(hr))
        {
            // Create the HBITMAP
            if (pBitmap->GetHBITMAP(Color(0, 0, 0), phBitmap) != Gdiplus::Ok)
            {
                hr = E_FAIL;
            }
        }

        delete pBitmap;
    }

    return hr;
}

bool IsWhitespace(_In_ char ch)
{
    return (ch == '\r' || ch == '\n' || ch == '\t' || ch == '\v' || ch == ' ');
}

HRESULT CPIFImageHelpers::ReadLineFromStream(
    _In_ IStream* pStream,
    _Outptr_ char** ppBuffer)
{
    *ppBuffer = nullptr;
    // Get the current seek position, so we can go back to the beginning and read it in one
    // call so we know the size of the buffer to create.
    ULARGE_INTEGER uliOriginSeek;
    HRESULT hr = pStream->Seek(c_liZero, STREAM_SEEK_CUR, &uliOriginSeek);
    if (SUCCEEDED(hr))
    {
        char chCurrent = ' ';
        // Walk past any leading whitespace
        while (SUCCEEDED(hr) && IsWhitespace(chCurrent))
        {
            hr = IStream_Read(pStream, (void*)&chCurrent, sizeof(char));
            if (IsWhitespace(chCurrent))
            {
                // If we found leading white space chars we need to update
                // our origin.
                hr = pStream->Seek(c_liZero, STREAM_SEEK_CUR, &uliOriginSeek);
            }
        }

        // Are we a comment?  If so we need to only break on newlines.
        bool isComment = (chCurrent == '#');

        // Walk through the actual content

        // Walk through the actual content. If it is a comment we read everything until a newline.
        // If it is anything else we break on the first whitespace
        while (SUCCEEDED(hr) &&
            ((isComment && chCurrent != '\n') ||
            (!isComment && !IsWhitespace(chCurrent))))
        {
            hr = IStream_Read(pStream, (void*)&chCurrent, sizeof(char));
        }

        if (SUCCEEDED(hr))
        {
            // Verify we are currently at a whitespace char
            hr = IsWhitespace(chCurrent) ? S_OK : E_FAIL;
            if (SUCCEEDED(hr))
            {
                // Get the current stream read position
                ULARGE_INTEGER uliCurrentSeek;
                hr = pStream->Seek(c_liZero, STREAM_SEEK_CUR, &uliCurrentSeek);
                if (SUCCEEDED(hr))
                {
                    hr = (uliCurrentSeek.LowPart > uliOriginSeek.LowPart) ? S_OK : E_FAIL;
                    if (SUCCEEDED(hr))
                    {
                        UINT uBufferSize = uliCurrentSeek.LowPart - uliOriginSeek.LowPart;
                        hr = uBufferSize > 0 ? S_OK : E_FAIL;
                        if (SUCCEEDED(hr))
                        {
                            char* pBuffer = new char[uBufferSize];
                            hr = pBuffer ? S_OK : E_OUTOFMEMORY;
                            if (SUCCEEDED(hr))
                            {
                                // Ensure zero terminated
                                *pBuffer = '\0';
                                // Reset the stream to the origin
                                LARGE_INTEGER seekSet = { 0 };
                                seekSet.LowPart = uliOriginSeek.LowPart;
                                seekSet.HighPart = uliOriginSeek.HighPart;
                                hr = pStream->Seek(seekSet, STREAM_SEEK_SET, nullptr);
                                if (SUCCEEDED(hr))
                                {
                                    hr = IStream_Read(pStream, (void*)pBuffer, sizeof(char) * uBufferSize);
                                }

                                if (SUCCEEDED(hr))
                                {
                                    *ppBuffer = pBuffer;
                                }
                                else
                                {
                                    delete [] pBuffer;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    return hr;
}

HRESULT CPIFImageHelpers::ParseImageLineForNum(
    _In_ IStream* pStream,
    _In_ PCSTR pszFormat,
    _Out_ UINT* value)
{
    *value = 0;
    HRESULT hr = S_OK;
    char* pBuffer = nullptr;
    while (SUCCEEDED(hr) && *value == 0)
    {
        hr = ReadLineFromStream(pStream, &pBuffer);
        if (SUCCEEDED(hr))
        {
            // Skip comment lines
            if (*pBuffer != '#')
            {
                hr = (sscanf_s(pBuffer, pszFormat, value) == 1) ? S_OK : E_FAIL;
            }

            delete [] pBuffer;
            pBuffer = nullptr;
        }
    }

    return hr;
}

HRESULT CPIFImageHelpers::ReadImageHeaders(
    _In_ IStream* pStream,
    _Out_ PortableImageFormatType* pFormatType,
    _Out_ UINT* puWidth,
    _Out_ UINT* puHeight,
    _Out_ UINT* puMax)
{
    *pFormatType = PortableImageFormatType_Invalid;
    *puWidth = 0;
    *puHeight = 0;
    *puMax = 1;

    HRESULT hr = pStream ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        // First get the portable image type
        while (SUCCEEDED(hr) && *pFormatType == PortableImageFormatType_Invalid)
        {
            // Verify type
            UINT type = 0;
            hr = ParseImageLineForNum(pStream, "P%u", &type);
            if (SUCCEEDED(hr))
            {
                hr = (type > PortableImageFormatType_Invalid && type < PortableImageFormatType_Count) ? S_OK : E_FAIL;
                if (SUCCEEDED(hr))
                {
                    *pFormatType = static_cast<PortableImageFormatType>(type);
                }
            }
        }

        if (SUCCEEDED(hr))
        {
            // Parse width, height and max.  PBMA and PBMB don't have max values.
            hr = ParseImageLineForNum(pStream, "%u", puWidth);
            if (SUCCEEDED(hr))
            {
                hr = ParseImageLineForNum(pStream, "%u", puHeight);
                if (SUCCEEDED(hr) &&
                    ((*pFormatType != PortableImageFormatType_PBMA) &&
                    (*pFormatType != PortableImageFormatType_PBMB)))
                {
                    hr = ParseImageLineForNum(pStream, "%u", puMax);
                }
            }

            if (SUCCEEDED(hr))
            {
                // Validate read parameters
                hr = (*puWidth > 0 && *puHeight > 0) ? S_OK : E_FAIL;
            }
        }
    }
    return hr;
}

HRESULT CPIFImageHelpers::ReadImageFromStream(
    _In_ IStream* pStream,
    _Out_ PortableImageFormatType* pFormatType,
    _Out_ UINT* puWidth,
    _Out_ UINT* puHeight,
    _Out_ UINT* puMax,
    _Out_ UINT* puDataLen,
    _Outptr_ BYTE** ppData)
{
    *pFormatType = PortableImageFormatType_Invalid;
    *puWidth = 0;
    *puHeight = 0;
    *puMax = 1;
    *puDataLen = 0;
    *ppData = nullptr;

    HRESULT hr = ReadImageHeaders(pStream, pFormatType, puWidth, puHeight, puMax);
    if (SUCCEEDED(hr))
    {
        char chCurrent = ' ';
        // Walk past any remaining whitespace
        while (SUCCEEDED(hr) && IsWhitespace(chCurrent))
        {
            hr = IStream_Read(pStream, (void*)&chCurrent, sizeof(char));
        }

        if (SUCCEEDED(hr))
        {
            // Move back before the last read char
            ULARGE_INTEGER uliCurrentSeek;
            hr = pStream->Seek(c_liZero, STREAM_SEEK_CUR, &uliCurrentSeek);
            if (SUCCEEDED(hr))
            {
                LARGE_INTEGER seekSet = { 0 };
                seekSet.LowPart = uliCurrentSeek.LowPart - 1;
                seekSet.HighPart = uliCurrentSeek.HighPart;
                hr = pStream->Seek(seekSet, STREAM_SEEK_SET, nullptr);
            }
        }
    }

    if (SUCCEEDED(hr))
    {
        size_t sizeAlloc = 0;

        if ((*pFormatType == PortableImageFormatType_PPMA) ||
            (*pFormatType == PortableImageFormatType_PPMB))
        {
            sizeAlloc = (*puWidth) * (*puHeight) * 3;
        }
        else if ((*pFormatType == PortableImageFormatType_PGMA) ||
            (*pFormatType == PortableImageFormatType_PGMB) ||
            (*pFormatType == PortableImageFormatType_PBMA) ||
            (*pFormatType == PortableImageFormatType_PBMB))
        {
            sizeAlloc = (*puWidth) * (*puHeight);
        }
        else
        {
            hr = E_INVALIDARG;
        }

        if (SUCCEEDED(hr))
        {
            BYTE* pData = new BYTE[sizeAlloc];
            hr = pData ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                ZeroMemory(pData, sizeof(BYTE) * sizeAlloc);
                ULONG cbRead = 0;

                if ((*pFormatType == PortableImageFormatType_PPMB) ||
                    (*pFormatType == PortableImageFormatType_PGMB))
                {
                    // If we are reading a binary file we can read it in one go
                    hr = pStream->Read((void*)pData, sizeof(BYTE) * sizeAlloc, &cbRead);
                }
                else if (*pFormatType == PortableImageFormatType_PBMB)
                {
                    // We are packing bits for black/white pixels, not bytes
                    size_t sizeAllocBits = (((*puWidth) * (*puHeight)) / 8);
                    if ((((*puWidth) * (*puHeight)) % 8) > 0)
                    {
                        // Add an extra byte per row to account for some extra bits
                        // that didn't fit.
                        sizeAllocBits += (*puHeight);
                    }

                    BYTE* pBitsData = new BYTE[sizeAllocBits];
                    hr = pBitsData ? S_OK : E_OUTOFMEMORY;
                    if (SUCCEEDED(hr))
                    {
                        ZeroMemory(pBitsData, sizeof(BYTE) * sizeAllocBits);
                        hr = pStream->Read((void*)pBitsData, sizeof(BYTE) * sizeAllocBits, &cbRead);
                        if (SUCCEEDED(hr))
                        {
                            UINT shift;
                            BYTE currentByte = 0;
                            BYTE* pBytesDataCurr = pData;
                            BYTE* pBitsDataCurr = pBitsData;
                            for (UINT j = 0; j < (*puHeight); j++)
                            {
                                for (UINT i = 0; i < (*puWidth); i++)
                                {
                                    if (i % 8 == 0)
                                    {
                                        currentByte = (*pBitsDataCurr);
                                        pBitsDataCurr += 1;
                                    }
                                    shift = 7 - i % 8;
                                    *pBytesDataCurr = (currentByte >> shift) % 2;
                                    pBytesDataCurr += 1;
                                }
                            }
                        }

                        delete[] pBitsData;
                    }
                }
                else
                {
                    // Reading ascii format
                    ULARGE_INTEGER dataSize;
                    hr = IStream_Size(pStream, &dataSize);
                    if (SUCCEEDED(hr))
                    {
                        // We need to take into account spaces between pixel data.  Typically we will have
                        // to account for > 2.5x sizeAlloc to account for the extra spaces
                        char* pAsciiData = new char[dataSize.LowPart];
                        hr = pAsciiData ? S_OK : E_OUTOFMEMORY;
                        if (SUCCEEDED(hr))
                        {
                            BYTE* pDataCurr = pData;
                            ZeroMemory(pAsciiData, sizeof(char) * dataSize.LowPart);
                            hr = pStream->Read((void*)pAsciiData, sizeof(char) * (dataSize.LowPart - 1), &cbRead);
                            if (SUCCEEDED(hr))
                            {
                                char* pCurrNumStart = pAsciiData;
                                char* pCurrNumEnd = pAsciiData;

                                while (SUCCEEDED(hr) && pCurrNumStart && *pCurrNumStart != '\0')
                                {
                                    // First walk any leading whitespace
                                    while (IsWhitespace(*pCurrNumStart))
                                    {
                                        pCurrNumStart++;
                                        pCurrNumEnd = pCurrNumStart;
                                    }

                                    // Walk until we hit whitespace
                                    while (pCurrNumEnd && *pCurrNumEnd != '\0' &&
                                        !IsWhitespace(*pCurrNumEnd))
                                    {
                                        pCurrNumEnd++;
                                    }

                                    // Read in the current value
                                    if (pCurrNumStart && *pCurrNumStart != '\0' &&
                                        pCurrNumEnd && *pCurrNumEnd != '\0')
                                    {
                                        *pCurrNumEnd = '\0';

                                        hr = ((pDataCurr - pData) < sizeAlloc) ? S_OK : E_FAIL;
                                        if (SUCCEEDED(hr))
                                        {
                                            *pDataCurr = static_cast<BYTE>(atoi(pCurrNumStart));
                                            // Move our insertion pointer for the data
                                            pDataCurr++;
                                        }
                                    }

                                    // Set our two pointers to the next value
                                    if (pCurrNumEnd)
                                    {
                                        pCurrNumEnd++;
                                        pCurrNumStart = pCurrNumEnd;
                                    }
                                }

                                if (SUCCEEDED(hr))
                                {
                                    // Did we read the expected ammount of data?
                                    hr = (((pDataCurr - pData) == sizeAlloc - 1) ||
                                          ((pDataCurr - pData) == sizeAlloc)) ? S_OK : E_FAIL;
                                }
                            }
                            delete[] pAsciiData;
                        }
                    }
                }

                if (SUCCEEDED(hr))
                {
                    *puDataLen = sizeAlloc;
                    *ppData = pData;
                }
                else
                {
                    delete[] pData;
                }
            }
        }
    }

    return hr;
}