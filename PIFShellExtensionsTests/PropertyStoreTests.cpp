#include "stdafx.h"
#include "CppUnitTest.h"
#include <PIFShellExtensions.h>
#include <propkey.h>

using namespace Gdiplus;
using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace PIFShellExtensionsTests
{
    TEST_CLASS(PropertyStoreTests)
    {
    public:
        TEST_CLASS_INITIALIZE(Setup)
        {
        }

        TEST_CLASS_CLEANUP(Cleanup)
        {
        }

        TEST_METHOD(VerifyGetPropertyStore)
        {
            CComPtr<IStream> spStream;
            Assert::IsTrue(SHCreateStreamOnFileEx(SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\fool.ascii.pbm", STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream) == S_OK);

            CPIFShellExt* pPIFShellExt = new CPIFShellExt();
            Assert::IsNotNull(pPIFShellExt);

            CComPtr<IInitializeWithStream> spInitWithStream;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spInitWithStream)) == S_OK);

            Assert::IsTrue(spInitWithStream->Initialize(spStream, STGM_READ) == S_OK);

            CComPtr<IPropertyStore> spPropStore;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spPropStore)) == S_OK);

            PROPVARIANT pv = {};
            Assert::IsTrue(spPropStore->GetValue(PKEY_Image_Dimensions, &pv) == S_OK);

            pPIFShellExt->Release();
        }

        TEST_METHOD(VerifyInvalidFileGetPropertyStore)
        {
            CComPtr<IStream> spStream;
            Assert::IsTrue(SHCreateStreamOnFileEx(SZ_ROOT_TEST_FOLDER L"Invalid\\PGMB\\barbara_zero_width_and_height.pgm", STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream) == S_OK);

            CPIFShellExt* pPIFShellExt = new CPIFShellExt();
            Assert::IsNotNull(pPIFShellExt);

            CComPtr<IInitializeWithStream> spInitWithStream;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spInitWithStream)) == S_OK);

            Assert::IsTrue(spInitWithStream->Initialize(spStream, STGM_READ) == S_OK);

            CComPtr<IPropertyStore> spPropStore;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spPropStore)) == S_OK);

            PROPVARIANT pv = {};
            Assert::IsFalse(spPropStore->GetValue(PKEY_Image_Dimensions, &pv) == S_OK);

            pPIFShellExt->Release();
        }


        // TODO:
        // Verify dimensions are correct
        // Verify with invalid file types
    };
}