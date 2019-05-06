/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiTestImage.cpp
**
**	Description:
**		VsiImage Test
**
********************************************************************************/
#include "stdafx.h"
#include <ATLComTime.h>
#include <VsiUnitTestHelper.h>
#include <VsiPdmModule.h>
#include <VsiBModeModule.h>
#include <VsiStudyModule.h>

using namespace System;
using namespace System::Text;
using namespace System::Collections::Generic;
using namespace Microsoft::VisualStudio::TestTools::UnitTesting;
using namespace System::IO;

namespace VsiStudyModuleTest
{
	[TestClass]
	public ref class VsiTestImage
	{
	private:
		TestContext^ testContextInstance;
		ULONG_PTR m_cookie;

	public: 
		/// <summary>
		///Gets or sets the test context which provides
		///information about and functionality for the current test run.
		///</summary>
		property Microsoft::VisualStudio::TestTools::UnitTesting::TestContext^ TestContext
		{
			Microsoft::VisualStudio::TestTools::UnitTesting::TestContext^ get()
			{
				return testContextInstance;
			}
			System::Void set(Microsoft::VisualStudio::TestTools::UnitTesting::TestContext^ value)
			{
				testContextInstance = value;
			}
		};

		#pragma region Additional test attributes
		//
		//You can use the following additional attributes as you write your tests:
		//
		//Use ClassInitialize to run code before running the first test in the class
		//[ClassInitialize()]
		//static void MyClassInitialize(TestContext^ testContext) {};
		//
		//Use ClassCleanup to run code after all tests in a class have run
		//[ClassCleanup()]
		//static void MyClassCleanup() {};
		//
		//Use TestInitialize to run code before running each test
		[TestInitialize()]
		void MyTestInitialize()
		{
			m_cookie = 0;

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			ACTCTX actCtx;
			memset((void*)&actCtx, 0, sizeof(ACTCTX));
			actCtx.cbSize = sizeof(ACTCTX);
			actCtx.dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID | ACTCTX_FLAG_RESOURCE_NAME_VALID;
			actCtx.lpSource = L"VsiStudyModuleTest.dll";
			actCtx.lpAssemblyDirectory = pszTestDir;
			actCtx.lpResourceName = MAKEINTRESOURCE(2);

			HANDLE hCtx = CreateActCtx(&actCtx);
			Assert::IsTrue(INVALID_HANDLE_VALUE != hCtx, "CreateActCtx failed.");

			ULONG_PTR cookie;
			BOOL bRet = ActivateActCtx(hCtx, &cookie);
			Assert::IsTrue(FALSE != bRet, "ActivateActCtx failed.");

			m_cookie = cookie;
		};
		//
		//Use TestCleanup to run code after each test has run
		[TestCleanup()]
		void MyTestCleanup()
		{
			if (0 != m_cookie)
				DeactivateActCtx(0, m_cookie);
		};
		//
		#pragma endregion 

		//test methods in IVsiImage
		[TestMethod]
		void TestCreate()
		{
			CComPtr<IVsiImage> pImage;

			HRESULT hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
		};

		/// <summary>
		/// <step>Create a Image with following attributes:</step>
		/// <step>Data path: "Test\UnitTestImage.xml"</step>
		/// <step>Image ID: "01234567890123456789"</step>
		/// <step>Image name: "Image 0123456789"</step>
		/// <step>Clone image</step>
		/// <step>Compare cloned image attributes<results>Images attributes are identical</results></step>
		/// </summary>
		[TestMethod]
		void TestClone()
		{
			CComPtr<IVsiImage> pImage;

			HRESULT hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestImage.vxml", pszTestDir);

			hr = pImage->SetDataPath(strPath);
			Assert::AreEqual(S_OK, hr, "SetDataPath failed.");

			// ID
			CComVariant vSetId(L"01234567890123456789");
			hr = pImage->SetProperty(VSI_PROP_IMAGE_ID, &vSetId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Name
			CComVariant vSetName(L"Image 0123456789");
			hr = pImage->SetProperty(VSI_PROP_IMAGE_NAME, &vSetName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Clone
			CComPtr<IVsiImage> pImageOut;
			hr = pImage->Clone(&pImageOut);
			Assert::AreEqual(S_OK, hr, "Clone failed.");

			// Path
			CComHeapPtr<OLECHAR> pszGetPath;
			hr = pImageOut->GetDataPath(&pszGetPath);
			Assert::AreEqual(S_OK, hr, "GetDataPath failed.");
			CString strGetPath(pszGetPath);
			Assert::IsTrue(strPath == strGetPath, "Clone failed.");

			// ID
			CComVariant vGetId;
			hr = pImageOut->GetProperty(VSI_PROP_IMAGE_ID, &vGetId);
			Assert::IsTrue(vSetId == vGetId, "Clone failed.");

			// Name
			CComVariant vGetName;
			hr = pImageOut->GetProperty(VSI_PROP_IMAGE_NAME, &vGetName);
			Assert::IsTrue(vSetName == vGetName, "Clone failed.");
		}

		/// <summary>
		/// <step>Create a Image</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_IMAGE_ERROR_NONE</results></step>
		/// <step>Load a good image data file</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_IMAGE_ERROR_NONE</results></step>
		/// <step>Load a invalid image data file</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_IMAGE_ERROR_LOAD_XML</results></step>
		/// <step>Load a version incompatible image data file</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_IMAGE_ERROR_VERSION_INCOMPATIBLE</results></step>
		/// </summary>
		[TestMethod]
		void TestGetErrorCode()
		{
			// App
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			// Study Manager
			CComPtr<IVsiStudyManager> pStudyMgr;

			HRESULT hr = pStudyMgr.CoCreateInstance(__uuidof(VsiStudyManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pStudyMgr->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Series
			CComPtr<IVsiSeries> pSeries;

			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Image
			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pImage->SetSeries(pSeries);
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			CComQIPtr<IVsiPersistImage> pPersist(pImage);
			Assert::IsTrue(NULL != pPersist, "IVsiPersistImage missing.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			// Error code
			VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_LOAD_XML);
			hr = pImage->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_IMAGE_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

			// Test good series
			CString strPath;
			strPath.Format(L"%s\\UnitTestImage.vxml", pszTestDir);

			hr = pPersist->LoadImageData(strPath, NULL, NULL, 0);
			Assert::AreEqual(S_OK, hr, "LoadImageData failed.");

			hr = pImage->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_IMAGE_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

			// Test bad series
			strPath.Format(L"%s\\UnitTestImageInvalid.vxml", pszTestDir);

			hr = pPersist->LoadImageData(strPath, NULL, NULL, 0);
			Assert::IsTrue(FAILED(hr), "LoadImageData failed.");

			hr = pImage->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(0 < (VSI_IMAGE_ERROR_LOAD_XML & dwErrorCode), "GetErrorCode failed.");

			// Test version incompatible series
			strPath.Format(L"%s\\UnitTestImageVersionIncompatible.vxml", pszTestDir);

			hr = pPersist->LoadImageData(strPath, NULL, NULL, 0);
			Assert::AreEqual(S_OK, hr, "LoadImageData failed.");

			hr = pImage->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(0 < (VSI_IMAGE_ERROR_VERSION_INCOMPATIBLE & dwErrorCode), "GetErrorCode failed.");

			// Clean up
			hr = pStudyMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Set data path: "TestPath\UnitTestImage.vxml"</step>
		/// <step>Retrieve data path<results>The data path is "TestPath\UnitTestImage.vxml"</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetDataPath()
		{
			// Image
			CComPtr<IVsiImage> pImage;

			HRESULT hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestImage.vxml", pszTestDir);

			hr = pImage->SetDataPath(strPath);
			Assert::AreEqual(S_OK, hr, "SetDataPath failed.");

			// Path
			CComHeapPtr<OLECHAR> pszGetPath;
			hr = pImage->GetDataPath(&pszGetPath);
			Assert::AreEqual(S_OK, hr, "GetDataPath failed.");

			Assert::IsTrue(strPath == pszGetPath, "Paths not match.");
		}


		/// <summary>
		/// <step>Set image property</step>
		/// <step>Retrieve image property<results>The image property that is retrieved must match one that is set before</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetProperties()
		{
			// Image
			CComPtr<IVsiImage> pImage;

			HRESULT hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			VSI_PROP_IMAGE propsBstr[] = {
				VSI_PROP_IMAGE_ID,						// VT_BSTR
				VSI_PROP_IMAGE_ID_SERIES,				// VT_BSTR
				VSI_PROP_IMAGE_ID_STUDY,				// VT_BSTR
				VSI_PROP_IMAGE_NS,						// VT_BSTR - Namespace
				VSI_PROP_IMAGE_NAME,					// VT_BSTR
				VSI_PROP_IMAGE_MODE,					// VT_BSTR - Mode name
				VSI_PROP_IMAGE_MODE_DISPLAY,			// VT_BSTR - Mode name for display
				VSI_PROP_IMAGE_ACQID,					// VT_BSTR
				VSI_PROP_IMAGE_KEY,						// VT_BSTR
				VSI_PROP_IMAGE_VERSION_REQUIRED,		// VT_BSTR
				VSI_PROP_IMAGE_VERSION,					// VT_BSTR
		
				VSI_PROP_IMAGE_VEVOSTRAIN_SERIES_NAME,	// VT_BSTR - series name at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOSTRAIN_STUDY_NAME,	// VT_BSTR - study name at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOSTRAIN_NAME,			// VT_BSTR - image label at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOSTRAIN_ANIMAL_ID,	// VT_BSTR - Animal ID at the time that VevoStrain was processed

				VSI_PROP_IMAGE_VEVOCQ_SERIES_NAME,		// VT_BSTR - series name at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOCQ_STUDY_NAME,		// VT_BSTR - study name at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOCQ_NAME,				// VT_BSTR - image label at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOCQ_ANIMAL_ID,		// VT_BSTR - Animal ID at the time that VevoStrain was processed
			};

			for (int i = 0; i < _countof(propsBstr); ++i)
			{
				CString strSet;
				strSet.Format(L"%d%d%d", i, i, i);

				CComVariant vSet(strSet.GetString());
				hr = pImage->SetProperty(propsBstr[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				CComVariant vGet;
				hr = pImage->GetProperty(propsBstr[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");
			}

			VSI_PROP_IMAGE propsBool[] = {
				VSI_PROP_IMAGE_FRAME_TYPE,				// VT_BOOL
				VSI_PROP_IMAGE_RF_DATA,					// VT_BOOL - true if RF data
				VSI_PROP_IMAGE_VEVOSTRAIN,				// VT_BOOL - true if VEVOSTRAIN processing has been applied to this image
				VSI_PROP_IMAGE_VEVOSTRAIN_UPDATED,		// VT_BOOL - true if image has been updated since VEVOSTRAIN processing
				VSI_PROP_IMAGE_VEVOCQ,					// VT_BOOL - true if VEVOSTRAIN processing has been applied to this image
				VSI_PROP_IMAGE_VEVOCQ_UPDATED,			// VT_BOOL - true if image has been updated since VEVOSTRAIN processing
			};

			for (int i = 0; i < _countof(propsBool); ++i)
			{
				CComVariant vSet(true);
				hr = pImage->SetProperty(propsBool[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				CComVariant vGet;
				hr = pImage->GetProperty(propsBool[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");

				vSet = false;
				hr = pImage->SetProperty(propsBool[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				vGet.Clear();
				hr = pImage->GetProperty(propsBool[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");
			}

			VSI_PROP_IMAGE propsDate[] = {
				VSI_PROP_IMAGE_ACQUIRED_DATE,			// VT_DATE
				VSI_PROP_IMAGE_CREATED_DATE,			// VT_DATE
				VSI_PROP_IMAGE_MODIFIED_DATE,			// VT_DATE
			};

			for (int i = 0; i < _countof(propsDate); ++i)
			{
				SYSTEMTIME st;
				GetSystemTime(&st);

				COleDateTime date(st);

				CComVariant vSet;
				V_VT(&vSet) = VT_DATE;
				V_DATE(&vSet) = DATE(date);

				hr = pImage->SetProperty(propsDate[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				CComVariant vGet;
				hr = pImage->GetProperty(propsDate[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");
			}

			VSI_PROP_IMAGE propsR8[] = {
				VSI_PROP_IMAGE_LENGTH,					// VT_R8
				VSI_PROP_IMAGE_THUMBNAIL_FRAME,			// VT_R8
			};

			for (int i = 0; i < _countof(propsR8); ++i)
			{
				double dbVal = 1234;

				CComVariant vSet(dbVal);
				hr = pImage->SetProperty(propsR8[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				CComVariant vGet;
				hr = pImage->GetProperty(propsR8[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");
			}

		}

		/// <summary>
		/// <step>For each slot</step>
		/// <step>SetSeries</step>
		/// <step>GetSeries<results>Series are identical</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetSeries()
		{
			// Image
			CComPtr<IVsiImage> pImage;

			HRESULT hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Create series
			CComPtr<IVsiSeries> pSeriesIn;
			hr = pSeriesIn.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
			{
				hr = pImage->SetSeries(pSeriesIn);
				Assert::AreEqual(S_OK, hr, "SetSeries failed.");

				CComPtr<IVsiSeries> pSeriesOut;
				hr = pImage->GetSeries(&pSeriesOut);
				Assert::AreEqual(S_OK, hr, "GetSeries failed.");

				Assert::IsTrue(pSeriesIn == pSeriesOut, "GetSeries failed.");
			}
		}

		/// <summary>
		/// <step>SetStudy</step>
		/// <step>GetStudy<results>Studies are identical</results></step>
		/// </summary>
		[TestMethod]
		void TestGetStudy()
		{
			// Create series
			CComPtr<IVsiSeries> pSeries;
			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Create study
			CComPtr<IVsiStudy> pStudyIn;
			hr = pStudyIn.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Image
			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
	
			CComVariant vSetId(L"01234567890123456789");
			hr = pImage->SetProperty(VSI_PROP_IMAGE_ID, &vSetId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			hr = pSeries->AddImage(pImage);
			Assert::AreEqual(S_OK, hr, "Add Image failed.");

			hr = pSeries->SetStudy(pStudyIn);
			Assert::AreEqual(S_OK, hr, "SetStudy failed.");

			CComPtr<IVsiStudy> pStudyOut;
			hr = pImage->GetStudy(&pStudyOut);
			Assert::AreEqual(S_OK, hr, "GetStudy failed.");

			Assert::IsTrue(pStudyIn == pStudyOut, "GetStudy failed.");
		}

		/// <summary>
		/// <step>Set different acquired date/time for 2 images</step>
		/// <step>Compare<results>acquisition data/time for the 2 images are not identical</results></step>
		/// </summary>
		[TestMethod]
		void TestCompare()
		{
			// Image
			CComPtr<IVsiImage> pImage1;

			HRESULT hr = pImage1.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
	
			CComPtr<IVsiImage> pImage2;

			hr = pImage2.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
	
			SYSTEMTIME st;
			GetSystemTime(&st);

			COleDateTime date(st);
			COleDateTime date1;
			COleDateTime date2;

			date1.SetDateTime(2011, 01, 10, 8, 0, 0);
			date2.SetDateTime(2011, 01, 10, 7, 0, 0);

			CComVariant vSet1;
			V_VT(&vSet1) = VT_DATE;
			V_DATE(&vSet1) = DATE(date1);

			CComVariant vSet2;
			V_VT(&vSet2) = VT_DATE;
			V_DATE(&vSet2) = DATE(date2);

			hr = pImage1->SetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vSet1);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			hr = pImage2->SetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vSet2);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			int nCmp = 0;
			hr = pImage1->Compare(pImage2, VSI_PROP_IMAGE_ACQUIRED_DATE, &nCmp);
			Assert::AreEqual(S_OK, hr, "Compare failed.");

			Assert::IsTrue(nCmp > 0, "The compare is not correct.");			
		}

		/// <summary>
		/// <step>SetDataPath</step>
		/// <step>GetThumbnailFile<results>Path are identical</results></step>
		/// </summary>
		[TestMethod]
		void TestGetThumbnailFile()
		{
			// Image
			CComPtr<IVsiImage> pImage;

			HRESULT hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
	
			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestImage.vxml", pszTestDir);
			CString strPathThumbnail;
			strPathThumbnail.Format(L"%s\\UnitTestImage.png", pszTestDir);

			hr = pImage->SetDataPath(strPath);
			Assert::AreEqual(S_OK, hr, "SetDataPath failed.");

			CComHeapPtr<OLECHAR> pFile;
			hr = pImage->GetThumbnailFile(&pFile);
			Assert::AreEqual(S_OK, hr, "GetThumbnailFile failed.");
			
			Assert::IsTrue(strPathThumbnail.Compare(CComBSTR(pFile)) == 0, "Wrong Thumbnail file path.");
		}

		//test methods in IVsiPersitImage
		/// <summary>
		/// <step>Create a Image object</step>
		/// <step>Load a image from a test file</step>
		/// <step>Check all properties against expected values<results>The call must succeed and the properties must match</results></step>
		/// </summary>
		[TestMethod]
		void TestLoadImageData()
		{
			// App
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			// Study Manager
			CComPtr<IVsiStudyManager> pStudyMgr;

			HRESULT hr = pStudyMgr.CoCreateInstance(__uuidof(VsiStudyManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pStudyMgr->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Series
			CComPtr<IVsiSeries> pSeries;

			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Image
			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pImage->SetSeries(pSeries);
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			CComQIPtr<IVsiPersistImage> pPersist(pImage);
			Assert::IsTrue(NULL != pPersist, "IVsiPersistImage missing.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestImage.vxml", pszTestDir);

			hr = pPersist->LoadImageData(strPath, NULL, NULL, 0);
			Assert::AreEqual(S_OK, hr, "LoadImageData failed.");

			VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_LOAD_XML);
			hr = pImage->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_IMAGE_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

			CComVariant vExpected(L"0123456789ABCDEFGHIJKLMON");
			CComVariant vGet;
			hr = pImage->GetProperty(VSI_PROP_IMAGE_ID, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"0123456789ABCDEFGHIJKLMON");
			hr = pImage->GetProperty(VSI_PROP_IMAGE_ID_SERIES, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"0123456789ABCDEFGHIJKLMON");
			hr = pImage->GetProperty(VSI_PROP_IMAGE_ID_STUDY, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"Image Test");
			hr = pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			SYSTEMTIME st;
			st.wYear = 2010;
			st.wMonth = 03;
			st.wDay = 15;
			st.wHour = 19;
			st.wMinute = 19;
			st.wSecond = 45;
			st.wMilliseconds = 0;
			COleDateTime date(st);
			V_VT(&vExpected) = VT_DATE;
			V_DATE(&vExpected) = DATE(date);
			hr = pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			st.wYear = 2010;
			st.wMonth = 03;
			st.wDay = 15;
			st.wHour = 19;
			st.wMinute = 19;
			st.wSecond = 56;
			st.wMilliseconds = 0;
			date = COleDateTime(st);
			V_VT(&vExpected) = VT_DATE;
			V_DATE(&vExpected) = DATE(date);
			hr = pImage->GetProperty(VSI_PROP_IMAGE_CREATED_DATE, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			st.wYear = 2010;
			st.wMonth = 03;
			st.wDay = 15;
			st.wHour = 19;
			st.wMinute = 19;
			st.wSecond = 56;
			st.wMilliseconds = 0;
			date = COleDateTime(st);
			V_VT(&vExpected) = VT_DATE;
			V_DATE(&vExpected) = DATE(date);
			hr = pImage->GetProperty(VSI_PROP_IMAGE_MODIFIED_DATE, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"B-Mode");
			hr = pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"B-Mode");
			hr = pImage->GetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(100.0);
			hr = pImage->GetProperty(VSI_PROP_IMAGE_LENGTH, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"0123456789ABCDEFGHIJKLMON");
			hr = pImage->GetProperty(VSI_PROP_IMAGE_ACQID, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"0123456789ABCDEFGHIJKLMON");
			hr = pImage->GetProperty(VSI_PROP_IMAGE_KEY, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"1.0.0.0");
			hr = pImage->GetProperty(VSI_PROP_IMAGE_VERSION_REQUIRED, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(false);
			hr = pImage->GetProperty(VSI_PROP_IMAGE_FRAME_TYPE, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"2");
			hr = pImage->GetProperty(VSI_PROP_IMAGE_VERSION, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(100.0);
			hr = pImage->GetProperty(VSI_PROP_IMAGE_THUMBNAIL_FRAME, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(false);
			hr = pImage->GetProperty(VSI_PROP_IMAGE_RF_DATA, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(true);
			hr = pImage->GetProperty(VSI_PROP_IMAGE_VEVOSTRAIN_UPDATED, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(true);
			hr = pImage->GetProperty(VSI_PROP_IMAGE_VEVOCQ_UPDATED, &vGet); // SonoVevoUpdated
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			// Clean up
			hr = pStudyMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}
		
		//This test case needs couple more hours future developments from a senior developer.
		///// <summary>
		///// <step>Create a Image object</step>
		///// <step>Set some properties to the object</step>
		///// <step>Save a image to a test file</step>
		///// <step>Load a image from a test file</step>
		///// <step>Check all properties against expected values<results>The call must succeed and the properties must match</results></step>
		///// </summary>
		//[TestMethod]
		//void TestSaveImageData()
		//{
		//	// App
		//	CComPtr<IVsiApp> pApp;
		//	CreateApp(&pApp);
		//	Assert::IsTrue(pApp != NULL, "Create App failed.");

		//	// Study Manager
		//	CComPtr<IVsiStudyManager> pStudyMgr;

		//	HRESULT hr = pStudyMgr.CoCreateInstance(__uuidof(VsiStudyManager));
		//	Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

		//	hr = pStudyMgr->Initialize(pApp);
		//	Assert::AreEqual(S_OK, hr, "Initialize failed.");

		//	// Series
		//	CComPtr<IVsiSeries> pSeries;

		//	hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
		//	Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

		//	// Image
		//	CComPtr<IVsiImage> pImage;

		//	hr = pImage.CoCreateInstance(__uuidof(VsiImage));
		//	Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

		//	hr = pImage->SetSeries(pSeries);
		//	Assert::AreEqual(S_OK, hr, "SetSeries failed.");

		//	CComQIPtr<IVsiPersistImage> pPersist(pImage);
		//	Assert::IsTrue(NULL != pPersist, "IVsiPersistImage missing.");

		//	// Data path
		//	String^ strDir = testContextInstance->TestDeploymentDir;
		//	pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

		//	CComVariant vSetImageId(L"0123456789ABCDEFGHIJKLMON");
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_ID, &vSetImageId);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageIdSeries(L"0123456789ABCDEFGHIJKLMON");
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_ID_SERIES, &vSetImageIdSeries);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageIdStudy(L"0123456789ABCDEFGHIJKLMON");
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_ID_STUDY, &vSetImageIdStudy);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageName(L"Image Test");
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_NAME, &vSetImageName);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	SYSTEMTIME st;
		//	GetSystemTime(&st);
		//	COleDateTime date(st);

		//	CComVariant vSetImageAcquiredDate;
		//	V_VT(&vSetImageAcquiredDate) = VT_DATE;
		//	V_DATE(&vSetImageAcquiredDate) = DATE(date);
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vSetImageAcquiredDate);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageCreatedDate;
		//	V_VT(&vSetImageCreatedDate) = VT_DATE;
		//	V_DATE(&vSetImageCreatedDate) = DATE(date);
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_CREATED_DATE, &vSetImageCreatedDate);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageModifiedDate;
		//	V_VT(&vSetImageModifiedDate) = VT_DATE;
		//	V_DATE(&vSetImageModifiedDate) = DATE(date);
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_MODIFIED_DATE, &vSetImageModifiedDate);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageMode(L"B-Mode");
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_MODE, &vSetImageMode);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageModeDisplay(L"B-Mode");
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vSetImageModeDisplay);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageLength(100.0);
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_LENGTH, &vSetImageLength);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageAcqId(L"0123456789ABCDEFGHIJKLMON");
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_ACQID, &vSetImageAcqId);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageKey(L"0123456789ABCDEFGHIJKLMON");
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_KEY, &vSetImageKey);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageVersionRequired(L"1.0.0.0");
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_VERSION_REQUIRED, &vSetImageVersionRequired);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageFrameType(false);
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_FRAME_TYPE, &vSetImageFrameType);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageVersion(L"2");
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_VERSION, &vSetImageVersion);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageThumbnailFrame(100.0);
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_THUMBNAIL_FRAME, &vSetImageThumbnailFrame);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageRFData(false);
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_RF_DATA, &vSetImageRFData);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageVevoStrainUpdated(true);
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_VEVOSTRAIN_UPDATED, &vSetImageVevoStrainUpdated);
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComVariant vSetImageVevoCQUpdated(true);
		//	hr = pImage->SetProperty(VSI_PROP_IMAGE_VEVOCQ_UPDATED, &vSetImageVevoCQUpdated); // SonoVevoUpdated
		//	Assert::AreEqual(S_OK, hr, "SetProperty failed.");

		//	CComPtr<IVsiPdm> pPdm;

		//	hr = pPdm.CoCreateInstance(__uuidof(VsiPdm));
		//	Assert::AreEqual(S_OK, hr, "PDM CoCreateInstance failed.");

		//	hr = pPdm->Initialize(FALSE);
		//	Assert::AreEqual(S_OK, hr, "PDM initialize failed.");

		//	CComPtr<IVsiBMode> pBMode;

		//	hr = pBMode.CoCreateInstance(__uuidof(VsiBMode));
		//	Assert::AreEqual(S_OK, hr, "B-Mode CoCreateInstance failed.");

		//	CComQIPtr<IVsiMode> pMode(pBMode);
		//	Assert::IsTrue(pMode != NULL, "B-Mode initialize failed.");

		//	hr = pMode->Initialize(pApp, VSI_MODE_INIT_CREATE_DATA_VIEW);
		//	Assert::AreEqual(S_OK, hr, "B-Mode CoCreateInstance failed."); // it will returns E_FAIL here, it need to initialize much more things, need future development

		//	CString strPath;
		//	strPath.Format(L"%s\\TestSaveImage.vxml", pszTestDir);
		//	hr = pPersist->SaveImageData(strPath, pMode, pPdm, VSI_DATA_TYPE_MSMNT_PARAMS_ONLY);
		//	Assert::AreEqual(S_OK, hr, "SaveImageData failed.");

		//	hr = pPdm->Uninitialize();
		//	Assert::AreEqual(S_OK, hr, "PDM un-initialize failed.");


		//	CComPtr<IVsiImage> pImageLoaded;

		//	hr = pImageLoaded.CoCreateInstance(__uuidof(VsiImage));
		//	Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
		//	
		//	CComQIPtr<IVsiPersistImage> pPersistLoaded(pImageLoaded);
		//	Assert::IsTrue(NULL != pPersistLoaded, "IVsiPersistSeries missing.");

		//	hr = pPersistLoaded->LoadImageData(strPath, NULL, NULL, 0);
		//	Assert::AreEqual(S_OK, hr, "LoadImageData failed.");

		//	VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_LOAD_XML);
		//	hr = pImageLoaded->GetErrorCode(&dwErrorCode);
		//	Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
		//	Assert::IsTrue(VSI_IMAGE_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

		//	CComVariant vGetImageId;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_ID, &vGetImageId);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageId == vGetImageId, "Property value not match.");

		//	CComVariant vGetImageIdSeries;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_ID_SERIES, &vGetImageIdSeries);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageIdSeries == vGetImageIdSeries, "Property value not match.");

		//	CComVariant vGetImageIdStudy;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_ID_STUDY, &vGetImageIdStudy);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageIdStudy == vGetImageIdStudy, "Property value not match.");

		//	CComVariant vGetImageName;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_NAME, &vGetImageName);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageName == vGetImageName, "Property value not match.");

		//	CComVariant vGetImageAcquiredDate;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vGetImageAcquiredDate);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageAcquiredDate == vGetImageAcquiredDate, "Property value not match.");

		//	CComVariant vGetImageCreatedDate;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_CREATED_DATE, &vGetImageCreatedDate);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageCreatedDate == vGetImageCreatedDate, "Property value not match.");

		//	CComVariant vGetImageModifiedDate;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_MODIFIED_DATE, &vGetImageModifiedDate);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageModifiedDate == vGetImageModifiedDate, "Property value not match.");

		//	CComVariant vGetImageMode;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_MODE, &vGetImageMode);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageMode == vGetImageMode, "Property value not match.");

		//	CComVariant vGetImageModeDisplay;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vGetImageModeDisplay);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageModeDisplay == vGetImageModeDisplay, "Property value not match.");

		//	CComVariant vGetImageLength;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_LENGTH, &vGetImageLength);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageLength == vGetImageLength, "Property value not match.");

		//	CComVariant vGetImageAcqId;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_ACQID, &vGetImageAcqId);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageAcqId == vGetImageAcqId, "Property value not match.");

		//	CComVariant vGetImageKey;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_KEY, &vGetImageKey);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageKey == vGetImageKey, "Property value not match.");

		//	CComVariant vGetImageVersionRequired;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_VERSION_REQUIRED, &vGetImageVersionRequired);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageVersionRequired == vGetImageVersionRequired, "Property value not match.");

		//	CComVariant vGetImageFrameType;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_FRAME_TYPE, &vGetImageFrameType);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageFrameType == vGetImageFrameType, "Property value not match.");

		//	CComVariant vGetImageVersion;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_VERSION, &vGetImageVersion);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageVersion == vGetImageVersion, "Property value not match.");

		//	CComVariant vGetImageThumbnailFrame;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_THUMBNAIL_FRAME, &vGetImageThumbnailFrame);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageThumbnailFrame == vGetImageThumbnailFrame, "Property value not match.");

		//	CComVariant vGetImageRFData;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_RF_DATA, &vGetImageRFData);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageRFData == vGetImageRFData, "Property value not match.");

		//	CComVariant vGetImageVevoStrainUpdated;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_VEVOSTRAIN_UPDATED, &vGetImageVevoStrainUpdated);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageVevoStrainUpdated == vGetImageVevoStrainUpdated, "Property value not match.");

		//	CComVariant vGetImageVevoCQUpdated;
		//	hr = pImageLoaded->GetProperty(VSI_PROP_IMAGE_VEVOCQ_UPDATED, &vGetImageVevoCQUpdated);
		//	Assert::AreEqual(S_OK, hr, "GetProperty failed.");
		//	Assert::IsTrue(vSetImageVevoCQUpdated == vGetImageVevoCQUpdated, "Property value not match.");

		//	// Clean up
		//	hr = pStudyMgr->Uninitialize();
		//	Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

		//	CloseApp(pApp);

		//	// Clean Up
		//	String^ strDirTest = gcnew String(strDir);
		//	strDirTest += "\\TestSaveImage.vxml";
		//	File::Delete(strDirTest);
		//}

		/// <summary>
		/// <step>Create an image object</step>
		/// <step>Set many properties using SetProperty</step>
		/// <step>Call ResaveProperties to a predefined image file</step>
		/// <step>load image data from the image file</step>
		/// <step>Check all attributes on the root element<results>all attributes should be override correctly</results></step>
		/// </summary>
		[TestMethod]
		void TestResaveProperties()
		{
			// App
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			// Study Manager
			CComPtr<IVsiStudyManager> pStudyMgr;

			HRESULT hr = pStudyMgr.CoCreateInstance(__uuidof(VsiStudyManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pStudyMgr->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Series
			CComPtr<IVsiSeries> pSeries;

			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Image
			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
			
			hr = pImage->SetSeries(pSeries);
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			CComQIPtr<IVsiPersistImage> pPersist(pImage);
			Assert::IsTrue(NULL != pPersist, "IVsiPersistImage missing.");			

			VSI_PROP_IMAGE propsBstr[] = {
				VSI_PROP_IMAGE_ID,						// VT_BSTR
				VSI_PROP_IMAGE_ID_SERIES,				// VT_BSTR
				VSI_PROP_IMAGE_ID_STUDY,				// VT_BSTR
				VSI_PROP_IMAGE_NAME,					// VT_BSTR
				VSI_PROP_IMAGE_MODE,					// VT_BSTR - Mode name
				VSI_PROP_IMAGE_MODE_DISPLAY,			// VT_BSTR - Mode name for display
				VSI_PROP_IMAGE_ACQID,					// VT_BSTR
				VSI_PROP_IMAGE_KEY,						// VT_BSTR
				VSI_PROP_IMAGE_VERSION_REQUIRED,		// VT_BSTR
				VSI_PROP_IMAGE_VERSION,					// VT_BSTR
		
				VSI_PROP_IMAGE_VEVOSTRAIN_SERIES_NAME,	// VT_BSTR - series name at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOSTRAIN_STUDY_NAME,	// VT_BSTR - study name at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOSTRAIN_NAME,			// VT_BSTR - image label at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOSTRAIN_ANIMAL_ID,	// VT_BSTR - Animal ID at the time that VevoStrain was processed

				VSI_PROP_IMAGE_VEVOCQ_SERIES_NAME,		// VT_BSTR - series name at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOCQ_STUDY_NAME,		// VT_BSTR - study name at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOCQ_NAME,				// VT_BSTR - image label at the time that VevoStrain was processed
				VSI_PROP_IMAGE_VEVOCQ_ANIMAL_ID,		// VT_BSTR - Animal ID at the time that VevoStrain was processed
			};

			for (int i = 0; i < _countof(propsBstr); ++i)
			{
				CString strSet;
				strSet.Format(L"%d%d%d", i, i, i);

				CComVariant vSet(strSet.GetString());
				hr = pImage->SetProperty(propsBstr[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			}

			VSI_PROP_IMAGE propsBool[] = {
				VSI_PROP_IMAGE_FRAME_TYPE,				// VT_BOOL
				VSI_PROP_IMAGE_RF_DATA,					// VT_BOOL - true if RF data
				VSI_PROP_IMAGE_VEVOSTRAIN_UPDATED,		// VT_BOOL - true if image has been updated since VEVOSTRAIN processing
				VSI_PROP_IMAGE_VEVOCQ_UPDATED,			// VT_BOOL - true if image has been updated since VEVOSTRAIN processing
			};

			for (int i = 0; i < _countof(propsBool); ++i)
			{
				CComVariant vSet(false);
				hr = pImage->SetProperty(propsBool[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			}

			VSI_PROP_IMAGE propsDate[] = {
				VSI_PROP_IMAGE_ACQUIRED_DATE,			// VT_DATE
				VSI_PROP_IMAGE_CREATED_DATE,			// VT_DATE
				VSI_PROP_IMAGE_MODIFIED_DATE,			// VT_DATE
			};

			for (int i = 0; i < _countof(propsDate); ++i)
			{
				SYSTEMTIME st;
				GetSystemTime(&st);

				COleDateTime date(st);

				CComVariant vSet;
				V_VT(&vSet) = VT_DATE;
				V_DATE(&vSet) = DATE(date);

				hr = pImage->SetProperty(propsDate[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			}

			VSI_PROP_IMAGE propsR8[] = {
				VSI_PROP_IMAGE_LENGTH,					// VT_R8
				VSI_PROP_IMAGE_THUMBNAIL_FRAME,			// VT_R8
			};

			for (int i = 0; i < _countof(propsR8); ++i)
			{
				double dbVal = 1234;

				CComVariant vSet(dbVal);
				hr = pImage->SetProperty(propsR8[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			}

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestImage.vxml", pszTestDir);

			hr = pPersist->ResaveProperties(strPath);
			Assert::AreEqual(S_OK, hr, "ResaveProperties failed.");

			hr = pPersist->LoadImageData(strPath, NULL, NULL, 0);
			Assert::AreEqual(S_OK, hr, "LoadImageData failed.");

			VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_LOAD_XML);
			hr = pImage->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_IMAGE_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

			for (int i = 0; i < _countof(propsBstr); ++i)
			{
				CString strExpected;
				strExpected.Format(L"%d%d%d", i, i, i);

				CComVariant vExpected(strExpected.GetString());
				CComVariant vGet;
				hr = pImage->GetProperty(propsBstr[i], &vGet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");
	
				Assert::IsTrue(vGet == vExpected, "Property value not match.");

				vExpected.Clear();
				vGet.Clear();

			}

			for (int i = 0; i < _countof(propsBool); ++i)
			{
				CComVariant vExpected(false);
				CComVariant vGet;
				hr = pImage->GetProperty(propsBool[i], &vGet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				Assert::IsTrue(vGet == vExpected, "Property value not match.");

				vExpected.Clear();
				vGet.Clear();
			}

			for (int i = 0; i < _countof(propsDate); ++i)
			{
				SYSTEMTIME st;
				GetSystemTime(&st);

				COleDateTime date(st);

				CComVariant vExpected;
				V_VT(&vExpected) = VT_DATE;
				V_DATE(&vExpected) = DATE(date);

				CComVariant vGet;
				hr = pImage->GetProperty(propsDate[i], &vGet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				Assert::IsTrue(vGet == vExpected, "Property value not match.");

				vExpected.Clear();
				vGet.Clear();
			}

			for (int i = 0; i < _countof(propsR8); ++i)
			{
				double dbVal = 1234;

				CComVariant vExpected(dbVal);
				CComVariant vGet;
				hr = pImage->GetProperty(propsR8[i], &vGet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				Assert::IsTrue(vGet == vExpected, "Property value not match.");

				vExpected.Clear();
				vGet.Clear();
			}

			// Clean up
			hr = pStudyMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		//test methods in IVsiEnumImage

	private:
		void CreateApp(IVsiApp** ppApp)
		{
			CComPtr<IVsiTestApp> pTestApp;

			HRESULT hr = pTestApp.CoCreateInstance(__uuidof(VsiTestApp));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance app failed.");

			hr = pTestApp->InitServices(VSI_TESTAPP_SERVICE_PDM | VSI_TESTAPP_SERVICE_MODE_MANAGER);
			Assert::AreEqual(S_OK, hr, "InitServices failed.");

			CComQIPtr<IVsiApp> pApp(pTestApp);
			Assert::IsTrue(NULL != pApp, "Get IVsiApp failed.");

			*ppApp = pApp.Detach();
		}

		void CloseApp(IVsiApp* pApp)
		{
			CComQIPtr<IVsiTestApp> pTestApp(pApp);

			HRESULT hr = pTestApp->ShutdownServices();
			Assert::AreEqual(S_OK, hr, "ShutdownServices failed.");
		}
	};
}
