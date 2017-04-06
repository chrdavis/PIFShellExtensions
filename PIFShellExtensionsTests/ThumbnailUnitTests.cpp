#include "stdafx.h"
#include "CppUnitTest.h"
#include <PIFShellExtensions.h>
#include <gdiplus.h>
#include <propkey.h>

using namespace Gdiplus;
using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace PIFShellExtensionsTests
{
    PCWSTR g_pszTestOutputFolder = L"%temp%\\PIFShellExtensionsTestsOut";
    wchar_t g_szPathOut[MAX_PATH] = { 0 };

    ULONG_PTR g_gdiplusToken = 0;

    PCWSTR g_rgPBMATestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\fool.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\gerrymander.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\letter_a.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\scs.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\seahorse.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\sierpinski.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\washington.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\circle.ascii.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\feep.ascii.pbm"
    };

    PCWSTR g_rgPPMATestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMA\\star_field.ascii.ppm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMA\\blackbuck.ascii.ppm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMA\\feep.ascii.ppm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMA\\pbmlib.ascii.ppm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMA\\sines.ascii.ppm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMA\\snail.ascii.ppm"
    };

    PCWSTR g_rgPGMATestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\apollonian_gasket.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\baboon.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\balloons.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\balloons_noisy.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\barbara.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\body1.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\body2.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\body3.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\brain_398.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\brain_492.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\brain_508.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\brain_604.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\casablanca.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\coins.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\columns.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\dla.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\dragon.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\f14.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\feep.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\foliage.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\fractal_tree.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\galaxy.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\gator.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\glassware_noisy.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\hands.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\handsmat.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\lena.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\marcie.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\mona_lisa.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\mother_daughter.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\mountain.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\pbmlib.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\pepper.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\roi_14.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\saturn.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\screws.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\snap.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\surf.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\totem.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\tracks.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\venus1.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\venus2.ascii.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMA\\x31_f18.ascii.pgm",
    };

    PCWSTR g_rgPBMBTestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\apollonian_gasket.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\circle.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\dla.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\feep.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\fool.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\gerrymander.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\letter_a.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\math_emporium.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\pbmlib.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\scs.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\seahorse.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\sierpinski.pbm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PBMB\\washington.pbm",
    };

    PCWSTR g_rgPPMBTestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMB\\pbmlib.ppm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMB\\underwater_bmx.ppm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMB\\bell_206.ppm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMB\\blackbuck.ppm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PPMB\\newton.ppm"
    };

    PCWSTR g_rgPGMBTestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\surf.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\baboon.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\barbara.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\bird.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\buffalo.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\casablanca.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\coins.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\columns.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\dewey_defeats_truman.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\feep.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\ladyzhenskaya.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\lena.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\lena_noisy.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\milking.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\mona_lisa.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\pbmlib.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\pepper.pgm",
        SZ_ROOT_TEST_FOLDER L"Valid\\PGMB\\snap.pgm"
    };

    struct TestFilePathWithResult
    {
        PCWSTR path;
        HRESULT result;
    };

    TestFilePathWithResult g_rgInvalidTestFilesWithResult[] =
    {
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMA\\circle.ascii_invalid_type.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMA\\feep.ascii_too_little_data.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMA\\gerrymander.ascii_large_numbers.pbm", S_OK },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMA\\letter_a.ascii_neg_one_width_height.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMA\\letter_a.ascii_too_large_width.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMA\\letter_a.ascii_zero_width_height.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMA\\scs.ascii_no_image_data.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMB\\apollonian_gasket_to_much_data.pbm", S_OK },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMB\\circle_too_little_data.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMB\\dla_invalid_image_type.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMB\\feep_no_image_data.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMB\\letter_a_width_height_too_large.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PBMB\\scs_zero_width_zero_height.pbm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMA\\apollonian_gasket.ascii_invalid_image_type.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMA\\baboon.ascii_invalid_height.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMA\\balloons.ascii_zero_width_and_height.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMA\\balloons_noisy.ascii_too_little_data.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMA\\barbara.ascii_too_much_data.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMA\\body1.ascii_no_image_data.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMB\\baboon_invalid_image_type.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMB\\barbara_zero_width_and_height.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMB\\bird_too_little_data.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMB\\buffalo_too_much_data.pgm", S_OK },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMB\\casablanca_width_height_too_large.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMB\\coins_no_image_data.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PGMB\\columns_width_height_plus_1.pgm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMB\\bell_206_invalid_image_type.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMB\\blackbuck_zero_width_height.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMB\\newton_width_height_too_large.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMB\\pbmlib_no_image_data.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMB\\underwater_bmx_no_max_value.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\sines.ascii_reallylongcomments.ppm", S_OK },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\snail.ascii_invalidP9.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\blackbuck.ascii_too_little_image_data.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\blackbuck.ascii_too_much_data.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\feep.ascii_justcomment.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\feep.ascii_multiple_image_types.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\feep.ascii_nowidthheightmaxdata.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\noextensionppm", S_OK },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\pbmlib.ascii_noimagedata.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\sines.ascii_incorrect_max_value.ppm", S_OK },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\sines.ascii_incorrect_width_height_too_large.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\sines.ascii_incorrect_width_height_too_small.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\sines.ascii_incorrect_width_height_zero.ppm", E_FAIL },
        { SZ_ROOT_TEST_FOLDER L"Invalid\\PPMA\\sines.ascii_invalid_file_type.ppm", E_FAIL },
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
        CPIFShellExt* pPIFShellExt = new CPIFShellExt();
        Assert::IsNotNull(pPIFShellExt);

        CComPtr<IPersistFile> spPersistFile;
        Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spPersistFile)) == S_OK);

        for (DWORD i = 0; i < dwCount; i++)
        {
            Assert::IsTrue(spPersistFile->Load(rgFilePaths[i], 0) == S_OK);

            CComPtr<IExtractImage> spImageExtractor;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spImageExtractor)) == S_OK);

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

            // Verify Width and Height
            Assert::IsTrue(image->GetWidth() <= dwWidth && image->GetHeight() <= dwHeight);

            CLSID myClsId;
            GetEncoderClsid(L"image/bmp", &myClsId);

            wchar_t pathOutputFile[MAX_PATH] = { 0 };
            StringCchPrintf(pathOutputFile, ARRAYSIZE(pathOutputFile), L"%s\\%s_Extract.bmp", g_szPathOut, PathFindFileName(rgFilePaths[i]));
            image->Save(pathOutputFile, &myClsId, NULL);
            delete image;

            DeleteObject(hBitmap);
        }

        pPIFShellExt->Release();
    }

    void VerifyLoadFileArrayViaGetThumbnail(_In_ DWORD dwWidth, _In_ DWORD dwHeight, _In_ DWORD dwCount, _In_ PCWSTR* rgFilePaths)
    {
        CPIFShellExt* pPIFShellExt = new CPIFShellExt();
        Assert::IsNotNull(pPIFShellExt);

        for (DWORD i = 0; i < dwCount; i++)
        {
            CComPtr<IStream> spStream;
            Assert::IsTrue(SHCreateStreamOnFileEx(rgFilePaths[i], STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream) == S_OK);

            CComPtr<IInitializeWithStream> spInitWithStream;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spInitWithStream)) == S_OK);

            Assert::IsTrue(spInitWithStream->Initialize(spStream, STGM_READ) == S_OK);

            CComPtr<IThumbnailProvider> spThumbnailProvider;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spThumbnailProvider)) == S_OK);

            HBITMAP hBitmap;
            WTS_ALPHATYPE wts_alphatype;
            Assert::IsTrue(spThumbnailProvider->GetThumbnail(max(dwWidth, dwHeight), &hBitmap, &wts_alphatype) == S_OK);

            Assert::IsNotNull(hBitmap);

            Bitmap *image = new Bitmap(hBitmap, NULL);
            Assert::IsNotNull(image);

            // Verify Width and Height
            Assert::IsTrue(image->GetWidth() <= dwWidth && image->GetHeight() <= dwHeight);

            CLSID myClsId;
            GetEncoderClsid(L"image/bmp", &myClsId);

            wchar_t pathOutputFile[MAX_PATH] = { 0 };
            StringCchPrintf(pathOutputFile, ARRAYSIZE(pathOutputFile), L"%s\\%s_Thumbnail.bmp", g_szPathOut, PathFindFileName(rgFilePaths[i]));
            image->Save(pathOutputFile, &myClsId, NULL);
            delete image;

            DeleteObject(hBitmap);
        }

        pPIFShellExt->Release();
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

        // TODO: small size, zero size, different size
        TEST_METHOD(OnePixelThumbnailTest)
        {
            CComPtr<IStream> spStream;
            Assert::IsTrue(SHCreateStreamOnFileEx(SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\fool.ascii.pbm", STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream) == S_OK);

            CPIFShellExt* pPIFShellExt = new CPIFShellExt();
            Assert::IsNotNull(pPIFShellExt);

            CComPtr<IInitializeWithStream> spInitWithStream;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spInitWithStream)) == S_OK);

            Assert::IsTrue(spInitWithStream->Initialize(spStream, STGM_READ) == S_OK);

            CComPtr<IThumbnailProvider> spThumbnailProvider;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spThumbnailProvider)) == S_OK);

            HBITMAP hBitmap;
            WTS_ALPHATYPE wts_alphatype;
            Assert::IsFalse(spThumbnailProvider->GetThumbnail(1, &hBitmap, &wts_alphatype) == S_OK);

            pPIFShellExt->Release();
        }

        TEST_METHOD(ZeroPixelThumbnailTest)
        {
            CComPtr<IStream> spStream;
            Assert::IsTrue(SHCreateStreamOnFileEx(SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\fool.ascii.pbm", STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream) == S_OK);

            CPIFShellExt* pPIFShellExt = new CPIFShellExt();
            Assert::IsNotNull(pPIFShellExt);

            CComPtr<IInitializeWithStream> spInitWithStream;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spInitWithStream)) == S_OK);

            Assert::IsTrue(spInitWithStream->Initialize(spStream, STGM_READ) == S_OK);

            CComPtr<IThumbnailProvider> spThumbnailProvider;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spThumbnailProvider)) == S_OK);

            HBITMAP hBitmap;
            WTS_ALPHATYPE wts_alphatype;
            Assert::IsFalse(spThumbnailProvider->GetThumbnail(0, &hBitmap, &wts_alphatype) == S_OK);

            pPIFShellExt->Release();
        }

        TEST_METHOD(InvalidFileTests)
        {
            CPIFShellExt* pPIFShellExt = new CPIFShellExt();
            Assert::IsNotNull(pPIFShellExt);

            for (DWORD i = 0; i < ARRAYSIZE(g_rgInvalidTestFilesWithResult); i++)
            {
                CComPtr<IStream> spStream;
                Assert::IsTrue(SHCreateStreamOnFileEx(g_rgInvalidTestFilesWithResult[i].path, STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream) == S_OK);

                CComPtr<IInitializeWithStream> spInitWithStream;
                Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spInitWithStream)) == S_OK);

                Assert::IsTrue(spInitWithStream->Initialize(spStream, STGM_READ) == S_OK);

                CComPtr<IThumbnailProvider> spThumbnailProvider;
                Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spThumbnailProvider)) == S_OK);

                HBITMAP hBitmap;
                WTS_ALPHATYPE wts_alphatype;
                Assert::IsTrue(spThumbnailProvider->GetThumbnail(500, &hBitmap, &wts_alphatype) == g_rgInvalidTestFilesWithResult[i].result);

                if (FAILED(g_rgInvalidTestFilesWithResult[i].result))
                {
                    Assert::IsNull(hBitmap);
                }
                else
                {
                    Assert::IsNotNull(hBitmap);
                    DeleteObject(hBitmap);
                }
            }

            pPIFShellExt->Release();
        }
    };
}