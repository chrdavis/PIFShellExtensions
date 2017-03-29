#include "stdafx.h"
#include "CppUnitTest.h"
#include <PIFShellExtensions.h>
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
        // Verify dimensions are correct
        // Verify with invalid file types
    };
}