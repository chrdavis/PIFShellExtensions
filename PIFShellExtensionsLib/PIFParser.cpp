#include "stdafx.h"
#include "PIFParser.h"
#include <gdiplus.h>

using namespace Gdiplus;

const LARGE_INTEGER c_liZero = { 0, 0 };

//
// CBaseImageParser
//
CBaseImageParser::CBaseImageParser() :
    m_uWidth(0),
    m_uHeight(0),
    m_uMaxValue(0),
    m_uDataLen(0),
    m_spStream(nullptr),
    m_rgData(nullptr),
    m_imageType(PortableImageFormatType_Invalid)
{
    GdiplusStartupInput gdiPlusStatupInput;
    GdiplusStartup(&_gdiPlusToken, &gdiPlusStatupInput, nullptr);
}

CBaseImageParser::~CBaseImageParser()
{
    if (m_rgData)
    {
        delete [] m_rgData;
    }

    GdiplusShutdown(_gdiPlusToken);
}

HRESULT CBaseImageParser::CreateParserFromStream(_In_ IStream* pStream, _Outptr_ CBaseImageParser** ppImageParser)
{
    *ppImageParser = nullptr;

    PortableImageFormatType imageType;
    HRESULT hr = ReadImageType(pStream, &imageType);
    if (SUCCEEDED(hr))
    {
        // Reset the stream after reading the type so we can use it to initialize an image parser
        IStream_Reset(pStream);
        if ((imageType == PortableImageFormatType_PPMA) || (imageType == PortableImageFormatType_PPMB))
        {
            *ppImageParser = new CPPMImageParser();
        }
        else if ((imageType == PortableImageFormatType_PGMA) || (imageType == PortableImageFormatType_PGMB))
        {
            *ppImageParser = new CPGMImageParser();
        }
        else if ((imageType == PortableImageFormatType_PBMA) || (imageType == PortableImageFormatType_PBMB))
        {
            *ppImageParser = new CPBMImageParser();
        }

        hr = (*ppImageParser) ? S_OK : E_OUTOFMEMORY;
        if (SUCCEEDED(hr))
        {
            hr = (*ppImageParser)->InitializeFromStream(pStream);
            if (FAILED(hr))
            {
                delete (*ppImageParser);
            }
        }
    }

    return hr;
}

HRESULT CBaseImageParser::GetBitmap(
    _Out_ HBITMAP* phBitmap)
{
    *phBitmap = nullptr;

    HRESULT hr = (m_uWidth > 0 && m_uHeight > 0) ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        Bitmap *pBitmap = new Bitmap(m_uWidth, m_uHeight, PixelFormat24bppRGB);
        hr = pBitmap ? S_OK : E_OUTOFMEMORY;
        if (SUCCEEDED(hr))
        {
            hr = PopulateBitmap(pBitmap);
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
    }

    return hr;
}

HRESULT CBaseImageParser::ReadImageType(_In_ IStream* pStream, _Out_ PortableImageFormatType* pImageType)
{
    *pImageType = PortableImageFormatType_Invalid;
    HRESULT hr = pStream ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        while (SUCCEEDED(hr) && (*pImageType) == PortableImageFormatType_Invalid)
        {
            // Verify type
            UINT type = 0;
            hr = ParseImageLineForNum(pStream, "P%u", &type);
            if (SUCCEEDED(hr))
            {
                hr = (type > PortableImageFormatType_Invalid && type < PortableImageFormatType_Count) ? S_OK : E_FAIL;
                if (SUCCEEDED(hr))
                {
                    (*pImageType) = static_cast<PortableImageFormatType>(type);
                }
            }
        }
    }

    return hr;
}

HRESULT CBaseImageParser::ReadImageHeaders()
{
    HRESULT hr = m_spStream ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        // First get the portable image type
        hr = ReadImageType(m_spStream, &m_imageType);
        if (SUCCEEDED(hr))
        {
            // Parse width, height and max.  PBMA and PBMB don't have max values.
            hr = ParseImageLineForNum(m_spStream, "%u", &m_uWidth);
            if (SUCCEEDED(hr))
            {
                hr = ParseImageLineForNum(m_spStream, "%u", &m_uHeight);
                if (SUCCEEDED(hr))
                {
                    hr = ParseImageLineForNum(m_spStream, "%u", &m_uMaxValue);
                }
            }

            if (SUCCEEDED(hr))
            {
                // Validate read parameters
                hr = (m_uWidth > 0 && m_uHeight > 0) ? S_OK : E_FAIL;
            }
        }
    }
    return hr;
}

HRESULT CBaseImageParser::ReadBinaryImageData()
{
    HRESULT hr = IsDataSizeValid() ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        size_t sizeAlloc = GetAllocationSize();
        hr = sizeAlloc > 0 ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            delete[] m_rgData;
            m_rgData = new BYTE[sizeAlloc];
            hr = m_rgData ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                m_uDataLen = static_cast<UINT>(sizeAlloc);
                ULONG cbRead = 0;
                hr = m_spStream->Read((void*)m_rgData, sizeof(BYTE) * m_uDataLen, &cbRead);
                if (SUCCEEDED(hr))
                {
                    // Did we read the expected ammount of data?
                    hr = ((static_cast<size_t>(cbRead) == sizeAlloc - 1) ||
                          (static_cast<size_t>(cbRead) == sizeAlloc)) ? S_OK : E_FAIL;
                }

                if (FAILED(hr))
                {
                    delete [] m_rgData;
                    m_rgData = nullptr;
                    m_uDataLen = 0;
                }
            }
        }
    }

    return hr;
}

HRESULT CBaseImageParser::ReadAsciiImageData()
{
    // Reading ascii format
    // We need to take into account spaces between pixel data.  Typically we will have
    // to account for > 2.5x sizeAlloc to account for the extra spaces
    HRESULT hr = IsDataSizeValid() ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        size_t sizeAlloc = GetAllocationSize();
        hr = sizeAlloc > 0 ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            delete[] m_rgData;
            m_rgData = new BYTE[sizeAlloc];
            hr = m_rgData ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                m_uDataLen = static_cast<UINT>(sizeAlloc);
            }
        }
    }

    if (SUCCEEDED(hr))
    {
        ULARGE_INTEGER dataSize = { 0 };
        hr = IStream_Size(m_spStream, &dataSize);
        if (SUCCEEDED(hr))
        {
            char* pAsciiData = new char[dataSize.LowPart];
            hr = pAsciiData ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                BYTE* pDataCurr = m_rgData;
                ZeroMemory(pAsciiData, sizeof(char) * dataSize.LowPart);

                ULONG cbRead = 0;
                hr = m_spStream->Read((void*)pAsciiData, sizeof(char) * (dataSize.LowPart - 1), &cbRead);
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

                            hr = (static_cast<UINT>(pDataCurr - m_rgData) < m_uDataLen) ? S_OK : E_FAIL;
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
                        hr = ((static_cast<UINT>(pDataCurr - m_rgData) == m_uDataLen - 1) ||
                              (static_cast<UINT>(pDataCurr - m_rgData) == m_uDataLen)) ? S_OK : E_FAIL;
                    }
                }
                delete[] pAsciiData;
            }
        }
    }

    if (FAILED(hr))
    {
        delete[] m_rgData;
        m_rgData = nullptr;
        m_uDataLen = 0;
    }

    return hr;
}

HRESULT CBaseImageParser::InitializeFromStream(_In_ IStream* pStream)
{
    m_spStream = pStream;
    HRESULT hr = ReadImageHeaders();
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
        // Do a sniff test to check if sizeAlloc is expected.  Protect against an invalid file type
        // specifying a huge width and/or height causing us to allocate an enormous buffer.
        hr = IsDataSizeValid() ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            hr = ReadImageData();
        }
    }

    return hr;
}

bool CBaseImageParser::IsWhitespace(_In_ char ch)
{
    return (ch == '\r' || ch == '\n' || ch == '\t' || ch == '\v' || ch == ' ');
}

HRESULT CBaseImageParser::ReadLineFromStream(
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
            if (SUCCEEDED(hr) && IsWhitespace(chCurrent))
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
                                    delete[] pBuffer;
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

HRESULT CBaseImageParser::ParseImageLineForNum(
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

            delete[] pBuffer;
            pBuffer = nullptr;
        }
    }

    return hr;
}

HRESULT CBaseImageParser::ResizeBitmap(
    _In_ HBITMAP hbmpOriginal,
    _In_ bool fMaintainAspectRatio,
    _In_ SIZE* pSizeDesired,
    _Outptr_ HBITMAP* phbmpResized)
{
    *phbmpResized = nullptr;

    Bitmap* pBitmap = new Bitmap(hbmpOriginal, nullptr);
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

            Bitmap* pResizeBitmap = new Bitmap(desiredWidth, desiredHeight, PixelFormat24bppRGB);
            hr = pResizeBitmap ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                pResizeBitmap->SetResolution(pBitmap->GetHorizontalResolution(), pBitmap->GetVerticalResolution());

                Graphics* pgrResize = new Graphics(pResizeBitmap);
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

//
// CPPMImageParser
//
__override
size_t CPPMImageParser::GetAllocationSize()
{
    return m_uWidth * m_uHeight * 3;
}

__override
HRESULT CPPMImageParser::ReadImageData()
{
    HRESULT hr = E_FAIL;
    if (m_imageType == PortableImageFormatType_PPMA)
    {
        hr = ReadAsciiImageData();
    }
    else if (m_imageType == PortableImageFormatType_PPMB)
    {
        hr = ReadBinaryImageData();
    }

    return hr;
}

__override
bool CPPMImageParser::IsDataSizeValid()
{
    bool isValid = false;

    ULARGE_INTEGER dataSize = { 0 };
    if (SUCCEEDED(IStream_Size(m_spStream, &dataSize)))
    {
        isValid = (dataSize.LowPart >= static_cast<ULONGLONG>(GetAllocationSize()));
    }

    return isValid;
}

__override
HRESULT CPPMImageParser::PopulateBitmap(_In_ Bitmap* pBitmap)
{
    HRESULT hr = (m_uWidth * m_uHeight * 3 == m_uDataLen) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        BYTE* pCurr = m_rgData;
        for (UINT row = 0; row < m_uHeight; row++)
        {
            for (UINT col = 0; col < m_uWidth; col++)
            {
                // Color RGB values
                pBitmap->SetPixel(col, row, Color(pCurr[0], pCurr[1], pCurr[2]));
                pCurr += 3;
            }
        }
    }

    return hr;
}

//
// CPGMImageParser
//
__override
size_t CPGMImageParser::GetAllocationSize()
{
    return m_uWidth * m_uHeight;
}

__override
HRESULT CPGMImageParser::ReadImageData()
{
    HRESULT hr = E_FAIL;
    if (m_imageType == PortableImageFormatType_PGMA)
    {
        hr = ReadAsciiImageData();
    }
    else if (m_imageType == PortableImageFormatType_PGMB)
    {
        hr = ReadBinaryImageData();
    }
    return hr;
}

__override
bool CPGMImageParser::IsDataSizeValid()
{
    bool isValid = false;

    ULARGE_INTEGER dataSize = { 0 };
    if (SUCCEEDED(IStream_Size(m_spStream, &dataSize)))
    {
        isValid = (dataSize.LowPart >= static_cast<ULONGLONG>(GetAllocationSize()));
    }

    return isValid;
}

__override
HRESULT CPGMImageParser::PopulateBitmap(_In_ Bitmap* pBitmap)
{
    HRESULT hr = (m_uWidth * m_uHeight == m_uDataLen) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        BYTE* pCurr = m_rgData;
        for (UINT row = 0; row < m_uHeight; row++)
        {
            for (UINT col = 0; col < m_uWidth; col++)
            {
                // Grayscale values
                pBitmap->SetPixel(col, row, Color(pCurr[0], pCurr[0], pCurr[0]));
                pCurr += 1;
            }
        }
    }

    return hr;
}

//
// CPBMImageParser
//
__override
size_t CPBMImageParser::GetAllocationSize()
{
    return m_uWidth * m_uHeight;
}

// Provider our own version of this from the base since we don't have a max value field
__override
HRESULT CPBMImageParser::ReadImageHeaders()
{
    HRESULT hr = m_spStream ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        // First get the portable image type
        hr = ReadImageType(m_spStream, &m_imageType);
        if (SUCCEEDED(hr))
        {
            // Parse width and height.  PBM images don't have a max value field
            hr = ParseImageLineForNum(m_spStream, "%u", &m_uWidth);
            if (SUCCEEDED(hr))
            {
                hr = ParseImageLineForNum(m_spStream, "%u", &m_uHeight);
            }

            if (SUCCEEDED(hr))
            {
                // Validate read parameters
                hr = (m_uWidth > 0 && m_uHeight > 0) ? S_OK : E_FAIL;
            }
        }
    }
    return hr;
}

__override
HRESULT CPBMImageParser::ReadImageData()
{
    HRESULT hr = E_FAIL;
    if (m_imageType == PortableImageFormatType_PBMA)
    {
        hr = ReadAsciiImageData();
    }
    else if (m_imageType == PortableImageFormatType_PBMB)
    {
        hr = ReadBinaryImageData();
    }

    return hr;
}

__override
HRESULT CPBMImageParser::ReadBinaryImageData()
{
    HRESULT hr = IsDataSizeValid() ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        size_t sizeAlloc = GetAllocationSize();
        hr = sizeAlloc > 0 ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            delete[] m_rgData;
            m_rgData = new BYTE[sizeAlloc];
            hr = m_rgData ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                m_uDataLen = static_cast<UINT>(sizeAlloc);

                // We are packing bits for black/white pixels, not bytes
                size_t sizeAllocBits = ((m_uWidth * m_uHeight) / 8);
                if (((m_uWidth * m_uHeight) % 8) > 0)
                {
                    // Add an extra byte per row to account for some extra bits
                    // that didn't fit.
                    sizeAllocBits += m_uHeight;
                }

                BYTE* pBitsData = new BYTE[sizeAllocBits];
                hr = pBitsData ? S_OK : E_OUTOFMEMORY;
                if (SUCCEEDED(hr))
                {
                    ZeroMemory(pBitsData, sizeof(BYTE) * sizeAllocBits);
                    ULONG cbRead = 0;
                    hr = m_spStream->Read((void*)pBitsData, static_cast<ULONG>(sizeof(BYTE) * sizeAllocBits), &cbRead);
                    if (SUCCEEDED(hr))
                    {
                        UINT shift;
                        BYTE currentByte = 0;
                        BYTE* pBytesDataCurr = m_rgData;
                        BYTE* pBitsDataCurr = pBitsData;
                        for (UINT j = 0; j < m_uHeight; j++)
                        {
                            for (UINT i = 0; i < m_uWidth; i++)
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

                if (FAILED(hr))
                {
                    delete[] m_rgData;
                    m_rgData = nullptr;
                    m_uDataLen = 0;
                }
            }
        }
    }

    return hr;
}

__override
bool CPBMImageParser::IsDataSizeValid()
{
    bool isValid = false;

    ULARGE_INTEGER dataSize = { 0 };
    if (SUCCEEDED(IStream_Size(m_spStream, &dataSize)))
    {
        // We are packing bits for black/white pixels, not bytes
        size_t sizeAllocBits = ((m_uWidth * m_uHeight) / 8);
        if (((m_uWidth * m_uHeight) % 8) > 0)
        {
            // Add an extra byte per row to account for some extra bits
            // that didn't fit.
            sizeAllocBits += m_uHeight;
        }

        isValid = (dataSize.LowPart >= static_cast<ULONGLONG>(sizeAllocBits - m_uHeight));
    }

    return isValid;
}

__override
HRESULT CPBMImageParser::PopulateBitmap(_In_ Bitmap* pBitmap)
{
    HRESULT hr = (m_uWidth * m_uHeight == m_uDataLen) ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        BYTE* pCurr = m_rgData;
        for (UINT row = 0; row < m_uHeight; row++)
        {
            for (UINT col = 0; col < m_uWidth; col++)
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