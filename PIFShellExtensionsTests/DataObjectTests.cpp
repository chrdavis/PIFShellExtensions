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

        TEST_METHOD(VerifyDataObjectSimple)
        {
            CComPtr<IStream> spStream;
            Assert::IsTrue(SHCreateStreamOnFileEx(SZ_ROOT_TEST_FOLDER L"Valid\\PBMA\\fool.ascii.pbm", STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &spStream) == S_OK);

            CPIFShellExt* pPIFShellExt = new CPIFShellExt();
            Assert::IsNotNull(pPIFShellExt);

            CComPtr<IInitializeWithStream> spInitWithStream;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spInitWithStream)) == S_OK);

            Assert::IsTrue(spInitWithStream->Initialize(spStream, STGM_READ) == S_OK);

            CComPtr<IDataObject> spDataObject;
            Assert::IsTrue(pPIFShellExt->QueryInterface(IID_PPV_ARGS(&spDataObject)) == S_OK);

            pPIFShellExt->Release();
        }


        // TODO:
        // Verify Clipboard format for Bitmap and the bitmap is valid
        // Check case where we get the data object but have not initialized yet
    };
}