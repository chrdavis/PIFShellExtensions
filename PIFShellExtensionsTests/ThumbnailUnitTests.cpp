#include "stdafx.h"
#include "CppUnitTest.h"
#include <PIFShellExtensions.h>
#include <gdiplus.h>
#include <propkey.h>

using namespace Gdiplus;
using namespace Microsoft::VisualStudio::CppUnitTestFramework;

//
// READ THIS READ THIS READ THIS READ THIS READ THIS READ THIS READ THIS READ THIS READ THIS//
// Be sure to move the Test Data folder in this project to a known location like below
// or change the below path.
// READ THIS READ THIS READ THIS READ THIS READ THIS READ THIS READ THIS READ THIS READ THIS//
//

#define SZ_ROOT_TEST_FOLDER L"c:\\Test Data\\"

namespace PIFShellExtensionsTests
{
    PCWSTR g_pszTestOutputFolder = L"%temp%\\PIFShellExtensionsTestsOut";
    wchar_t g_szPathOut[MAX_PATH] = { 0 };

    ULONG_PTR g_gdiplusToken = 0;

    PCWSTR g_rgPBMATestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"PBMA\\fool.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMA\\gerrymander.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMA\\letter_a.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMA\\scs.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMA\\seahorse.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMA\\sierpinski.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMA\\washington.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMA\\circle.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMA\\feep.ascii.pbm"
    };

    PCWSTR g_rgPPMATestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"PPMA\\star_field.ascii.ppm",
        SZ_ROOT_TEST_FOLDER L"PPMA\\blackbuck.ascii.ppm",
        SZ_ROOT_TEST_FOLDER L"PPMA\\feep.ascii.ppm",
        SZ_ROOT_TEST_FOLDER L"PPMA\\pbmlib.ascii.ppm",
        SZ_ROOT_TEST_FOLDER L"PPMA\\sines.ascii.ppm",
        SZ_ROOT_TEST_FOLDER L"PPMA\\snail.ascii.ppm"
    };


    PCWSTR g_rgPGMATestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"PGMA\\apollonian_gasket.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\baboon.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\balloons.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\balloons_noisy.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\barbara.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\body1.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\body2.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\body3.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\brain_398.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\brain_492.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\brain_508.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\brain_604.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\casablanca.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\coins.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\columns.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\dla.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\dragon.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\f14.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\feep.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\foliage.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\fractal_tree.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\galaxy.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\gator.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\glassware_noisy.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\hands.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\handsmat.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\lena.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\marcie.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\mona_lisa.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\mother_daughter.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\mountain.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\pbmlib.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\pepper.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\roi_14.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\saturn.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\screws.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\snap.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\surf.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\totem.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\tracks.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\venus1.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\venus2.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMA\\x31_f18.ascii.pgm",
    };

    PCWSTR g_rgPBMBTestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"PBMB\\apollonian_gasket.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\circle.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\dla.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\feep.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\fool.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\gerrymander.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\letter_a.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\math_emporium.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\pbmlib.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\scs.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\seahorse.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\sierpinski.pbm",
        SZ_ROOT_TEST_FOLDER L"PBMB\\washington.pbm",
    };

    PCWSTR g_rgPPMBTestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"PPMB\\pbmlib.ppm",
        SZ_ROOT_TEST_FOLDER L"PPMB\\underwater_bmx.ppm",
        SZ_ROOT_TEST_FOLDER L"PPMB\\bell_206.ppm",
        SZ_ROOT_TEST_FOLDER L"PPMB\\blackbuck.ppm",
        SZ_ROOT_TEST_FOLDER L"PPMB\\newton.ppm"
    };

    PCWSTR g_rgPGMBTestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"PGMB\\surf.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\baboon.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\barbara.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\bird.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\buffalo.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\casablanca.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\coins.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\columns.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\dewey_defeats_truman.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\feep.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\ladyzhenskaya.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\lena.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\lena_noisy.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\milking.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\mona_lisa.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\pbmlib.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\pepper.pgm",
        SZ_ROOT_TEST_FOLDER L"PGMB\\snap.pgm"
    };

    int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
    {
        UINT  num = 0;          // number of image encoders
        UINT  size = 0;         // size of the image encoder array in bytes

        ImageCodecInfo* pImageCodecInfo = NULL;

        GetImageEncodersSize(&num, &size);
        if (size == 0)
            return -1;  // Failure

        pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
        if (pImageCodecInfo == NULL)
            return -1;  // Failure

        GetImageEncoders(num, size, pImageCodecInfo);

        for (UINT j = 0; j < num && j < size; ++j)
        {
            if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0)
            {
                *pClsid = pImageCodecInfo[j].Clsid;
                free(pImageCodecInfo);
                return j;  // Success
            }
        }
        free(pImageCodecInfo);
        return -1;  // Failure
    }

    void VerifyLoadFileArrayViaExtract(_In_ DWORD dwWidth, _In_ DWORD dwHeight, _In_ DWORD dwCount, _In_ PCWSTR* rgFilePaths)
    {
        CPIFImageExtractor* pPIFImageExtractor = new CPIFImageExtractor();
        Assert::IsNotNull(pPIFImageExtractor);

        CComPtr<IPersistFile> spPersistFile;
        Assert::IsTrue(pPIFImageExtractor->QueryInterface(IID_PPV_ARGS(&spPersistFile)) == S_OK);

        for (DWORD i = 0; i < dwCount; i++)
        {
            Assert::IsTrue(spPersistFile->Load(rgFilePaths[i], 0) == S_OK);

            CComPtr<IExtractImage> spImageExtractor;
            Assert::IsTrue(pPIFImageExtractor->QueryInterface(IID_PPV_ARGS(&spImageExtractor)) == S_OK);

            DWORD dwPriority = 0;
            DWORD dwColorDepth = 32;
            DWORD dwFlags = IEIFLAG_ASPECT;
            WCHAR szPath[MAX_PATH] = {};
            SIZE sizeImage = { (LONG)dwWidth, (LONG)dwHeight };

            Assert::IsTrue(spImageExtractor->GetLocation(
                szPath,
                ARRAYSIZE(szPath),
                &dwPriority,
                &sizeImage,
                dwColorDepth,
                &dwFlags) == S_OK);

            HBITMAP hBitmap;
            Assert::IsTrue(spImageExtractor->Extract(&hBitmap) == S_OK);
            Assert::IsNotNull(hBitmap);

            Bitmap *image = new Bitmap(hBitmap, NULL);
            Assert::IsNotNull(image);

            CLSID myClsId;
            GetEncoderClsid(L"image/bmp", &myClsId);

            wchar_t pathOutputFile[MAX_PATH] = { 0 };
            StringCchPrintf(pathOutputFile, ARRAYSIZE(pathOutputFile), L"%s\\%s_Extract.bmp", g_szPathOut, PathFindFileName(rgFilePaths[i]));
            image->Save(pathOutputFile, &myClsId, NULL);
            delete image;

            DeleteObject(hBitmap);
        }

        pPIFImageExtractor->Release();
    }

    void VerifyLoadFileArrayViaGetThumbnail(_In_ DWORD dwWidth, _In_ DWORD dwHeight, _In_ DWORD dwCount, _In_ PCWSTR* rgFilePaths)
    {
        CreateDirectory(g_pszTestOutputFolder, nullptr);

        CPIFImageExtractor* pPIFImageExtractor = new CPIFImageExtractor();
        Assert::IsNotNull(pPIFImageExtractor);

        for (DWORD i = 0; i < dwCount; i++)
        {
            CComPtr<IStream> spStream;
            Assert::IsTrue(SHCreateStreamOnFileEx(rgFilePaths[i], STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream) == S_OK);

            CComPtr<IInitializeWithStream> spInitWithStream;
            Assert::IsTrue(pPIFImageExtractor->QueryInterface(IID_PPV_ARGS(&spInitWithStream)) == S_OK);

            Assert::IsTrue(spInitWithStream->Initialize(spStream, STGM_READ) == S_OK);

            CComPtr<IThumbnailProvider> spThumbnailProvider;
            Assert::IsTrue(pPIFImageExtractor->QueryInterface(IID_PPV_ARGS(&spThumbnailProvider)) == S_OK);

            HBITMAP hBitmap;
            WTS_ALPHATYPE wts_alphatype;
            Assert::IsTrue(spThumbnailProvider->GetThumbnail(max(dwWidth, dwHeight), &hBitmap, &wts_alphatype) == S_OK);

            Assert::IsNotNull(hBitmap);

            Bitmap *image = new Bitmap(hBitmap, NULL);
            Assert::IsNotNull(image);

            CLSID myClsId;
            int retVal = GetEncoderClsid(L"image/bmp", &myClsId);

            wchar_t pathOutputFile[MAX_PATH] = { 0 };
            StringCchPrintf(pathOutputFile, ARRAYSIZE(pathOutputFile), L"%s\\%s_Thumbnail.bmp", g_szPathOut, PathFindFileName(rgFilePaths[i]));
            image->Save(pathOutputFile, &myClsId, NULL);
            delete image;

            DeleteObject(hBitmap);
        }

        pPIFImageExtractor->Release();
    }

    TEST_CLASS(ThumbnailUnitTests)
    {
    public:
        TEST_CLASS_INITIALIZE(Setup)
        {
            ExpandEnvironmentStrings(g_pszTestOutputFolder, g_szPathOut, ARRAYSIZE(g_szPathOut));
            // Create the test directory
            CreateDirectory(g_szPathOut, nullptr);

            // Init GDIPlus
            GdiplusStartupInput gdiplusStartupInput;
            GdiplusStartup(&g_gdiplusToken, &gdiplusStartupInput, NULL);
        }

        TEST_CLASS_CLEANUP(Cleanup)
        {
            // Delete the test directory
            SHFILEOPSTRUCT fileStruct = { 0 };
            fileStruct.pFrom = g_szPathOut;
            fileStruct.wFunc = FO_DELETE;
            fileStruct.fFlags = FOF_SILENT|FOF_NOERRORUI|FOF_NO_UI;
            SHFileOperation(&fileStruct);

            // Shutdown GDIPlus
            GdiplusShutdown(g_gdiplusToken);
        }


        TEST_METHOD(SimplePBMAExtractThumbnailTest)
        {
            VerifyLoadFileArrayViaExtract(500, 500, ARRAYSIZE(g_rgPBMATestFiles), g_rgPBMATestFiles);
        }

        TEST_METHOD(SimplePPMAExtractThumbnailTest)
        {
            VerifyLoadFileArrayViaExtract(500, 500, ARRAYSIZE(g_rgPPMATestFiles), g_rgPPMATestFiles);
        }

        TEST_METHOD(SimplePGMAExtractThumbnailTest)
        {
            VerifyLoadFileArrayViaExtract(500, 500, ARRAYSIZE(g_rgPGMATestFiles), g_rgPGMATestFiles);
        }

        TEST_METHOD(SimplePBMBExtractThumbnailTest)
        {
            VerifyLoadFileArrayViaExtract(500, 500, ARRAYSIZE(g_rgPBMBTestFiles), g_rgPBMBTestFiles);
        }

        TEST_METHOD(SimplePPMBExtractThumbnailTest)
        {
            VerifyLoadFileArrayViaExtract(500, 500, ARRAYSIZE(g_rgPPMBTestFiles), g_rgPPMBTestFiles);
        }

        TEST_METHOD(SimplePGMBExtractThumbnailTest)
        {
            VerifyLoadFileArrayViaGetThumbnail(500, 500, ARRAYSIZE(g_rgPGMBTestFiles), g_rgPGMBTestFiles);
        }

        TEST_METHOD(SimplePBMAGetThumbnailTest)
        {
            VerifyLoadFileArrayViaGetThumbnail(500, 500, ARRAYSIZE(g_rgPBMATestFiles), g_rgPBMATestFiles);
        }

        TEST_METHOD(SimplePPMAGetThumbnailTest)
        {
            VerifyLoadFileArrayViaGetThumbnail(500, 500, ARRAYSIZE(g_rgPPMATestFiles), g_rgPPMATestFiles);
        }

        TEST_METHOD(SimplePGMAGetThumbnailTest)
        {
            VerifyLoadFileArrayViaGetThumbnail(500, 500, ARRAYSIZE(g_rgPGMATestFiles), g_rgPGMATestFiles);
        }

        TEST_METHOD(SimplePBMBGetThumbnailTest)
        {
            VerifyLoadFileArrayViaGetThumbnail(500, 500, ARRAYSIZE(g_rgPBMBTestFiles), g_rgPBMBTestFiles);
        }

        TEST_METHOD(SimplePPMBGetThumbnailTest)
        {
            VerifyLoadFileArrayViaGetThumbnail(500, 500, ARRAYSIZE(g_rgPPMBTestFiles), g_rgPPMBTestFiles);
        }

        TEST_METHOD(SimplePGMBGetThumbnailTest)
        {
            VerifyLoadFileArrayViaGetThumbnail(500, 500, ARRAYSIZE(g_rgPGMBTestFiles), g_rgPGMBTestFiles);
        }

        TEST_METHOD(VerifyGetPropertyStore)
        {
            CComPtr<IStream> spStream;
            Assert::IsTrue(SHCreateStreamOnFileEx(SZ_ROOT_TEST_FOLDER L"PBMA\\fool.ascii.pbm", STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream) == S_OK);

            CPIFImageExtractor* pPIFImageExtractor = new CPIFImageExtractor();
            Assert::IsNotNull(pPIFImageExtractor);

            CComPtr<IInitializeWithStream> spInitWithStream;
            Assert::IsTrue(pPIFImageExtractor->QueryInterface(IID_PPV_ARGS(&spInitWithStream)) == S_OK);

            Assert::IsTrue(spInitWithStream->Initialize(spStream, STGM_READ) == S_OK);

            CComPtr<IPropertyStore> spPropStore;
            Assert::IsTrue(pPIFImageExtractor->QueryInterface(IID_PPV_ARGS(&spPropStore)) == S_OK);

            PROPVARIANT pv = {};
            Assert::IsTrue(spPropStore->GetValue(PKEY_Image_Dimensions, &pv) == S_OK);
        }

        // TODO:
        // Invalid file types
        // Verify sizes of returned images
        // Invalid resize size
    };
}