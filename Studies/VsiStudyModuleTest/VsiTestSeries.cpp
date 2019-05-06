/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiTestSeries.cpp
**
**	Description:
**		VsiSeries Test
**
********************************************************************************/
#include "stdafx.h"
#include <ATLComTime.h>
#include <VsiUnitTestHelper.h>
#include <VsiStudyModule.h>

using namespace System;
using namespace System::Text;
using namespace System::Collections::Generic;
using namespace Microsoft::VisualStudio::TestTools::UnitTesting;
using namespace System::IO;

namespace VsiStudyModuleTest
{
	[TestClass]
	public ref class VsiTestSeries
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

		//test methods in IVsiSeries
		[TestMethod]
		void TestCreate()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
		};

		/// <summary>
		/// <step>Create a Series with following attributes:</step>
		/// <step>Data path: "Test\UnitTestSeries.xml"</step>
		/// <step>Series ID: "01234567890123456789"</step>
		/// <step>Series name: "Series 0123456789"</step>
		/// <step>Clone series</step>
		/// <step>Compare cloned series attributes<results>Series attributes are identical</results></step>
		/// </summary>
		[TestMethod]
		void TestClone()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestSeries.vxml", pszTestDir);

			hr = pSeries->SetDataPath(strPath);
			Assert::AreEqual(S_OK, hr, "SetDataPath failed.");

			// ID
			CComVariant vSetId(L"01234567890123456789");
			hr = pSeries->SetProperty(VSI_PROP_SERIES_ID, &vSetId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Name
			CComVariant vSetName(L"Series 0123456789");
			hr = pSeries->SetProperty(VSI_PROP_SERIES_NAME, &vSetName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Clone
			CComPtr<IVsiSeries> pSeriesOut;
			hr = pSeries->Clone(&pSeriesOut);
			Assert::AreEqual(S_OK, hr, "Clone failed.");

			// Path
			CComHeapPtr<OLECHAR> pszGetPath;
			hr = pSeriesOut->GetDataPath(&pszGetPath);
			Assert::AreEqual(S_OK, hr, "GetDataPath failed.");
			CString strGetPath(pszGetPath);
			Assert::IsTrue(strPath == strGetPath, "Clone failed.");

			// ID
			CComVariant vGetId;
			hr = pSeriesOut->GetProperty(VSI_PROP_SERIES_ID, &vGetId);
			Assert::IsTrue(vSetId == vGetId, "Clone failed.");

			// Name
			CComVariant vGetName;
			hr = pSeriesOut->GetProperty(VSI_PROP_SERIES_NAME, &vGetName);
			Assert::IsTrue(vSetName == vGetName, "Clone failed.");
		}

		/// <summary>
		/// <step>Create a Series</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_SERIES_ERROR_NONE</results></step>
		/// <step>Load a good series data file</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_SERIES_ERROR_NONE</results></step>
		/// <step>Load a invalid series data file</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_SERIES_ERROR_LOAD_XML</results></step>
		/// <step>Load a version incompatible series data file</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_SERIES_ERROR_VERSION_INCOMPATIBLE</results></step>
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

			CComQIPtr<IVsiPersistSeries> pPersist(pSeries);
			Assert::IsTrue(NULL != pPersist, "IVsiPersistSeries missing.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			// Error code
			VSI_SERIES_ERROR dwErrorCode(VSI_SERIES_ERROR_LOAD_XML);
			hr = pSeries->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

			// Test good series
			CString strPath;
			strPath.Format(L"%s\\UnitTestSeries.vxml", pszTestDir);

			hr = pPersist->LoadSeriesData(strPath, VSI_DATA_TYPE_SERIES_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadSeriesData failed.");

			hr = pSeries->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

			// Test bad series
			strPath.Format(L"%s\\UnitTestSeriesInvalid.vxml", pszTestDir);

			hr = pPersist->LoadSeriesData(strPath, VSI_DATA_TYPE_SERIES_LIST, NULL);
			Assert::IsTrue(FAILED(hr), "LoadSeriesData failed.");

			hr = pSeries->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(0 < (VSI_SERIES_ERROR_LOAD_XML & dwErrorCode), "GetErrorCode failed.");

			// Test version incompatible series
			strPath.Format(L"%s\\UnitTestSeriesVersionIncompatible.vxml", pszTestDir);

			hr = pPersist->LoadSeriesData(strPath, VSI_DATA_TYPE_IMAGE_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadSeriesData failed.");

			hr = pSeries->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(0 < (VSI_SERIES_ERROR_VERSION_INCOMPATIBLE & dwErrorCode), "GetErrorCode failed.");

			// Clean up
			hr = pStudyMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Set data path: "TestPath\UnitTestSeries.vxml"</step>
		/// <step>Retrieve data path<results>The data path is "TestPath\UnitTestSeries.vxml"</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetDataPath()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestSeries.vxml", pszTestDir);

			hr = pSeries->SetDataPath(strPath);
			Assert::AreEqual(S_OK, hr, "SetDataPath failed.");

			// Path
			CComHeapPtr<OLECHAR> pszGetPath;
			hr = pSeries->GetDataPath(&pszGetPath);
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
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			VSI_PROP_SERIES propsBstr[] = {
				VSI_PROP_SERIES_ID,
				VSI_PROP_SERIES_ID_STUDY,
				VSI_PROP_SERIES_NS,
				VSI_PROP_SERIES_NAME,
				VSI_PROP_SERIES_ACQUIRED_BY,
				VSI_PROP_SERIES_VERSION_REQUIRED,
				VSI_PROP_SERIES_VERSION_CREATED,
				VSI_PROP_SERIES_NOTES,
				VSI_PROP_SERIES_APPLICATION,
				VSI_PROP_SERIES_MSMNT_PACKAGE,
				VSI_PROP_SERIES_ANIMAL_ID,
				VSI_PROP_SERIES_ANIMAL_COLOR,
				VSI_PROP_SERIES_STRAIN,
				VSI_PROP_SERIES_SOURCE,
				VSI_PROP_SERIES_WEIGHT,
				VSI_PROP_SERIES_TYPE,
				VSI_PROP_SERIES_HEART_RATE,
				VSI_PROP_SERIES_BODY_TEMP,
				VSI_PROP_SERIES_SEX,
				VSI_PROP_SERIES_CUSTOM1,
				VSI_PROP_SERIES_CUSTOM2,
				VSI_PROP_SERIES_CUSTOM3,
				VSI_PROP_SERIES_CUSTOM4,
				VSI_PROP_SERIES_CUSTOM5,
				VSI_PROP_SERIES_CUSTOM6,
				VSI_PROP_SERIES_CUSTOM7,
				VSI_PROP_SERIES_CUSTOM8,
				VSI_PROP_SERIES_CUSTOM9,
				VSI_PROP_SERIES_CUSTOM10,
				VSI_PROP_SERIES_CUSTOM11,
				VSI_PROP_SERIES_CUSTOM12,
				VSI_PROP_SERIES_CUSTOM13,
				VSI_PROP_SERIES_CUSTOM14,
			};

			for (int i = 0; i < _countof(propsBstr); ++i)
			{
				CString strSet;
				strSet.Format(L"%d%d%d", i, i, i);

				CComVariant vSet(strSet.GetString());
				hr = pSeries->SetProperty(propsBstr[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				CComVariant vGet;
				hr = pSeries->GetProperty(propsBstr[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");
			}

			VSI_PROP_SERIES propsBool[] = {
				VSI_PROP_SERIES_EXPORTED,
				VSI_PROP_SERIES_PREGNANT,
			};

			for (int i = 0; i < _countof(propsBool); ++i)
			{
				CComVariant vSet(true);
				hr = pSeries->SetProperty(propsBool[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				CComVariant vGet;
				hr = pSeries->GetProperty(propsBool[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");

				vSet = false;
				hr = pSeries->SetProperty(propsBool[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				vGet.Clear();
				hr = pSeries->GetProperty(propsBool[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");
			}

			VSI_PROP_SERIES propsDate[] = {
				VSI_PROP_SERIES_CREATED_DATE,
				VSI_PROP_SERIES_DATE_OF_BIRTH,
				VSI_PROP_SERIES_DATE_MATED,
				VSI_PROP_SERIES_DATE_PLUGGED,
			};

			for (int i = 0; i < _countof(propsDate); ++i)
			{
				SYSTEMTIME st;
				GetSystemTime(&st);

				COleDateTime date(st);

				CComVariant vSet;
				V_VT(&vSet) = VT_DATE;
				V_DATE(&vSet) = DATE(date);

				hr = pSeries->SetProperty(propsDate[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				CComVariant vGet;
				hr = pSeries->GetProperty(propsDate[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");
			}
		}
		
		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>LoadPropertyLabel from a test file<results>The call must succeed</results></step>
		/// <step>GetPropertyLabel for all properties against expected values<results>The call must succeed and the properties must match</results></step>
		/// </summary>
		[TestMethod]
		void TestGetPropertyLabel()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistSeries> pPersist(pSeries);
			Assert::IsTrue(NULL != pPersist, "Get IVsiPersistSeries failed.");

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestSeriesFields.xml", pszTestDir);

			hr = pPersist->LoadPropertyLabel(strPath);
			Assert::AreEqual(S_OK, hr, "LoadPropertyLabel failed.");

			CComHeapPtr<OLECHAR> pszSetLabel1(L"Anesthetic Type");
			CComHeapPtr<OLECHAR> pszGetLabel1;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM1, &pszGetLabel1);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel1, pszGetLabel1), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel2(L"Anesthetic On");
			CComHeapPtr<OLECHAR> pszGetLabel2;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM2, &pszGetLabel2);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel2, pszGetLabel2), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel3(L"Anesthetic Off");
			CComHeapPtr<OLECHAR> pszGetLabel3;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM3, &pszGetLabel3);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel3, pszGetLabel3), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel4(L"Protocol ID");
			CComHeapPtr<OLECHAR> pszGetLabel4;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM4, &pszGetLabel4);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel4, pszGetLabel4), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel5(L"Protocol Name");
			CComHeapPtr<OLECHAR> pszGetLabel5;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM5, &pszGetLabel5);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel5, pszGetLabel5), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel6(L"Injectable");
			CComHeapPtr<OLECHAR> pszGetLabel6;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM6, &pszGetLabel6);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel6, pszGetLabel6), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel7(L"Injection Site");
			CComHeapPtr<OLECHAR> pszGetLabel7;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM7, &pszGetLabel7);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel7, pszGetLabel7), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel8(L"Injection Amount");
			CComHeapPtr<OLECHAR> pszGetLabel8;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM8, &pszGetLabel8);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel8, pszGetLabel8), "Property value not match.");
		}

		/// <summary>
		/// <step>Create a Series</step>
		/// <step>Create a Study</step>
		/// <step>Set a Study for the Series</step>
		/// <step>Get a Study for the Series<results>The Study that is retrieved from the Series must match one that is set before</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetStudy()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "VsiSeries CoCreateInstance failed.");

			CComPtr<IVsiStudy> pStudy;
			hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "VsiStudy CoCreateInstance failed.");
			
			hr = pSeries->SetStudy(pStudy);
			Assert::AreEqual(S_OK, hr, "SetStudy failed.");

			CComPtr<IVsiStudy> pStudyRet;
			
			hr = pSeries->GetStudy(&pStudyRet);
			Assert::AreEqual(S_OK, hr, "GetStudy failed.");

			Assert::IsTrue(pStudy == pStudyRet, "Study objects do not match.");
		}

		/// <summary>
		/// <step>Create a Series</step>
		/// <step>Set properties relevant for Compare function</step>
		/// <step>Compare using all relevant properties<results>Series objects must match</results></step>
		/// </summary>
		[TestMethod]
		void TestCompare()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
			// Name
			CComVariant vSetName(L"Series 0123456789");
			hr = pSeries->SetProperty(VSI_PROP_SERIES_NAME, &vSetName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			// Acquired By
			CComVariant vSetAcq(L"Venkataratnam Narasimha Rattaiah");
			hr = pSeries->SetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vSetAcq);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			// Date
			SYSTEMTIME st;
			GetSystemTime(&st);

			COleDateTime date(st);

			CComVariant vSet;
			V_VT(&vSet) = VT_DATE;
			V_DATE(&vSet) = DATE(date);

			hr = pSeries->SetProperty(VSI_PROP_SERIES_CREATED_DATE, &vSet);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Clone
			CComPtr<IVsiSeries> pSeriesOut;
			hr = pSeries->Clone(&pSeriesOut);
			Assert::AreEqual(S_OK, hr, "Clone failed.");
			
			int iCmp;
			hr = pSeries->Compare(pSeriesOut, VSI_PROP_SERIES_NAME, &iCmp);
			Assert::AreEqual(S_OK, hr, "Compare call failed.");
			Assert::AreEqual(iCmp, 0, "Series are different.");

			hr = pSeries->Compare(pSeriesOut, VSI_PROP_SERIES_ACQUIRED_BY, &iCmp);
			Assert::AreEqual(S_OK, hr, "Compare call failed.");
			Assert::AreEqual(iCmp, 0, "Series are different.");

			hr = pSeries->Compare(pSeriesOut, VSI_PROP_SERIES_CREATED_DATE, &iCmp);
			Assert::AreEqual(S_OK, hr, "Compare call failed.");
			Assert::AreEqual(iCmp, 0, "Series are different.");
		}

		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>Create a Image object</step>
		/// <step>Set a dummy Id for the image object</step>
		/// <step>Add image object to series<results>The call must succeed</results></step>
		/// </summary>
		[TestMethod]
		void TestAddImage()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComVariant vId(L"45245634573565463456245625");
			hr = pImage->SetProperty(VSI_PROP_IMAGE_ID, &vId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			
			hr = pSeries->AddImage(pImage);
			Assert::AreEqual(S_OK, hr, "AddImage failed.");
		}

		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>Create a Image object</step>
		/// <step>Set a dummy Id for the image object</step>
		/// <step>Add image object to the series</step>
		/// <step>Remove image object from the series<results>The call must succeed</results></step>
		/// </summary>
		[TestMethod]
		void TestRemoveImage()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CString strId(L"45245634573565463456245625");
			CComVariant vId(strId);
			hr = pImage->SetProperty(VSI_PROP_IMAGE_ID, &vId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			
			hr = pSeries->AddImage(pImage);
			Assert::AreEqual(S_OK, hr, "AddImage failed.");
			
			hr = pSeries->RemoveImage(strId);
			Assert::AreEqual(S_OK, hr, "RemoveImage failed.");
		}

		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>Create 2 Image objects</step>
		/// <step>Set a dummy Id for each of the image objects</step>
		/// <step>Add image objects to the series</step>
		/// <step>Count images from the series<results>count should be two</results></step>
		/// <step>Remove all images from the series<results>The call must succeed</results></step>
		/// <step>Count images from the series<results>count should be zero</results></step>
		/// </summary>
		[TestMethod]
		void TestRemoveAllImages()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiImage> pImage1;

			hr = pImage1.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CString strId(L"45245634573565463456245625");
			CComVariant vId(strId);
			hr = pImage1->SetProperty(VSI_PROP_IMAGE_ID, &vId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComPtr<IVsiImage> pImage2;

			hr = pImage2.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			strId = L"45245634573565463456245626";
			vId = strId;
			hr = pImage2->SetProperty(VSI_PROP_IMAGE_ID, &vId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			
			hr = pSeries->AddImage(pImage1);
			Assert::AreEqual(S_OK, hr, "AddImage failed.");
			
			hr = pSeries->AddImage(pImage2);
			Assert::AreEqual(S_OK, hr, "AddImage failed.");

			LONG lCnt;
			hr = pSeries->GetImageCount(&lCnt, NULL);
			Assert::AreEqual(S_OK, hr, "GetImageCount failed.");

			Assert::IsTrue(lCnt == 2, "Image count different from expected.");
			
			hr = pSeries->RemoveAllImages();
			Assert::AreEqual(S_OK, hr, "RemoveAllImages failed.");

			hr = pSeries->GetImageCount(&lCnt, NULL);
			Assert::AreEqual(S_OK, hr, "GetImageCount failed.");

			Assert::IsTrue(lCnt == 0, "Image count different from expected.");
		}

		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>Create a Image object</step>
		/// <step>Set a dummy Id for the image object</step>
		/// <step>Add image object to the series</step>
		/// <step>Get image count from the series<results>Count must be 1</results></step>
		/// <step>Remove image object from the series</step>
		/// <step>Get image count from the series<results>Count must be 0</results></step>
		/// </summary>
		[TestMethod]
		void TestGetImageCount()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CString strId(L"45245634573565463456245625");
			CComVariant vId(strId);
			hr = pImage->SetProperty(VSI_PROP_IMAGE_ID, &vId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			
			hr = pSeries->AddImage(pImage);
			Assert::AreEqual(S_OK, hr, "AddImage failed.");
			
			LONG lCnt;
			hr = pSeries->GetImageCount(&lCnt, NULL);
			Assert::AreEqual(S_OK, hr, "GetImageCount failed.");

			Assert::IsTrue(lCnt == 1, "Image count different from expected.");
			
			hr = pSeries->RemoveImage(strId);
			Assert::AreEqual(S_OK, hr, "RemoveImage failed.");

			hr = pSeries->GetImageCount(&lCnt, NULL);
			Assert::AreEqual(S_OK, hr, "GetImageCount failed.");

			Assert::IsTrue(lCnt == 0, "Image count different from expected.");
		}

		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>Create a Image object</step>
		/// <step>Set a dummy Id for the image object</step>
		/// <step>Add image object to the series</step>
		/// <step>Get image enumerator from the series<results>The call must succeed and enumerator must be valid</results></step>
		/// </summary>
		[TestMethod]
		void TestGetImageEnumerator()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			// Study Manager
			CComPtr<IVsiStudyManager> pStudyMgr;

			HRESULT hr = pStudyMgr.CoCreateInstance(__uuidof(VsiStudyManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pStudyMgr->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComPtr<IVsiSeries> pSeries;

			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CString strId(L"45245634573565463456245625");
			CComVariant vId(strId);
			hr = pImage->SetProperty(VSI_PROP_IMAGE_ID, &vId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			
			hr = pSeries->AddImage(pImage);
			Assert::AreEqual(S_OK, hr, "AddImage failed.");

			CComPtr<IVsiEnumImage> pEnum;
			hr = pSeries->GetImageEnumerator(VSI_PROP_IMAGE_ID, TRUE, &pEnum, NULL);
			Assert::AreEqual(S_OK, hr, "GetImageEnumerator failed.");

			Assert::IsTrue(pEnum != 0, "Invalid enumerator returned.");

			// Clean up
			hr = pStudyMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>Create 2 Image objects</step>
		/// <step>Set a dummy Id for each of the image objects</step>
		/// <step>Add image objects to the series</step>
		/// <step>Get image 0 from the series with descending order<results>The call must succeed and the image must be equal to the first image set before</results></step>
		/// </summary>
		[TestMethod]
		void TestGetImageFromIndex()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			// Study Manager
			CComPtr<IVsiStudyManager> pStudyMgr;

			HRESULT hr = pStudyMgr.CoCreateInstance(__uuidof(VsiStudyManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pStudyMgr->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComPtr<IVsiSeries> pSeries;

			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiImage> pImage1;

			hr = pImage1.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CString strId(L"45245634573565463456245625");
			CComVariant vId(strId);
			hr = pImage1->SetProperty(VSI_PROP_IMAGE_ID, &vId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			CComVariant vName(L"TestImage45245634573565463456245625");
			hr = pImage1->SetProperty(VSI_PROP_IMAGE_NAME, &vName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComPtr<IVsiImage> pImage2;

			hr = pImage2.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			strId = L"45245634573565463456245626";
			vId = strId;
			hr = pImage2->SetProperty(VSI_PROP_IMAGE_ID, &vId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			vName = L"TestImage45245634573565463456245626";
			hr = pImage2->SetProperty(VSI_PROP_IMAGE_NAME, &vName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			
			hr = pSeries->AddImage(pImage1);
			Assert::AreEqual(S_OK, hr, "AddImage failed.");

			hr = pSeries->AddImage(pImage2);
			Assert::AreEqual(S_OK, hr, "AddImage failed.");

			CComPtr<IVsiImage> pImageRet;
			hr = pSeries->GetImageFromIndex(0, VSI_PROP_IMAGE_NAME, TRUE, &pImageRet); // descending order
			Assert::AreEqual(S_OK, hr, "GetImageFromIndex failed.");

			int iCmp;
			pImage1->Compare(pImageRet, VSI_PROP_IMAGE_NAME, &iCmp);
			Assert::IsTrue(iCmp == 0, "The compare is not correct.");
			pImage2->Compare(pImageRet, VSI_PROP_IMAGE_NAME, &iCmp);
			Assert::IsTrue(iCmp > 0, "The compare is not correct."); //pImage2 > pImageRet

			Assert::IsTrue(pImage1 == pImageRet, "Invalid image returned."); // Index 0 is the end of the image list (pImage2, pImage1)

			// Clean up
			hr = pStudyMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>Create a Image object</step>
		/// <step>Set a dummy Id for the image object</step>
		/// <step>Add image object to the series</step>
		/// <step>Get image from the series<results>The call must succeed and the image must be equal to one set before</results></step>
		/// </summary>
		[TestMethod]
		void TestGetImage()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CString strId(L"45245634573565463456245625");
			CComVariant vId(strId);
			hr = pImage->SetProperty(VSI_PROP_IMAGE_ID, &vId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			
			hr = pSeries->AddImage(pImage);
			Assert::AreEqual(S_OK, hr, "AddImage failed.");

			CComPtr<IVsiImage> pImageRet;
			hr = pSeries->GetImage(strId, &pImageRet);
			Assert::AreEqual(S_OK, hr, "GetImage failed.");

			Assert::IsTrue(pImage == pImageRet, "Invalid image returned.");
		}

		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>Create a Image object</step>
		/// <step>Set a dummy Id for the image object</step>
		/// <step>Add image object to the series</step>
		/// <step>Get item from the series with the image Id<results>The call must succeed and the image must be equal to one set before</results></step>
		/// </summary>
		[TestMethod]
		void TestGetItem()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CString strId(L"45245634573565463456245625");
			CComVariant vId(strId);
			hr = pImage->SetProperty(VSI_PROP_IMAGE_ID, &vId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			
			hr = pSeries->AddImage(pImage);
			Assert::AreEqual(S_OK, hr, "AddImage failed.");

			CComPtr<IUnknown> pImageRet;
			hr = pSeries->GetItem(strId, &pImageRet);
			Assert::AreEqual(S_OK, hr, "GetImage failed.");

			Assert::IsTrue(pImage == pImageRet, "Invalid image returned.");
		}

		//test methods in IVsiPersistSeries
		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>Load a series from a test file</step>
		/// <step>Check all properties against expected values<results>The call must succeed and the properties must match</results></step>
		/// </summary>
		[TestMethod]
		void TestLoadSeriesData()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
			
			CComQIPtr<IVsiPersistSeries> pPersist(pSeries);
			Assert::IsTrue(NULL != pPersist, "IVsiPersistSeries missing.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestSeries.vxml", pszTestDir);

			hr = pPersist->LoadSeriesData(strPath, VSI_DATA_TYPE_SERIES_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadSeriesData failed.");

			VSI_SERIES_ERROR dwErrorCode(VSI_SERIES_ERROR_LOAD_XML);
			hr = pSeries->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

			CComVariant vExpected(L"0123456789ABCDEFGHIJKLMON");
			CComVariant vGet;
			hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"0123456789ABCDEFGHIJKLMON");
			hr = pSeries->GetProperty(VSI_PROP_SERIES_ID_STUDY, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"SeriesNew");
			hr = pSeries->GetProperty(VSI_PROP_SERIES_NAME, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			SYSTEMTIME st;
			st.wYear = 2005;
			st.wMonth = 1;
			st.wDay = 26;
			st.wHour = 19;
			st.wMinute = 0;
			st.wSecond = 2;
			st.wMilliseconds = 687;
			COleDateTime date(st);
			V_VT(&vExpected) = VT_DATE;
			V_DATE(&vExpected) = DATE(date);
			hr = pSeries->GetProperty(VSI_PROP_SERIES_CREATED_DATE, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"Sierra Developer");
			hr = pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"1.0");
			hr = pSeries->GetProperty(VSI_PROP_SERIES_VERSION_REQUIRED, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");
		}

		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>Set some properties to the object</step>
		/// <step>Save a series to a test file</step>
		/// <step>Load a series from a test file</step>
		/// <step>Check all properties against expected values<results>The call must succeed and the properties must match</results></step>
		/// </summary>
		[TestMethod]
		void TestSaveSeriesData()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
			
			CComQIPtr<IVsiPersistSeries> pPersist(pSeries);
			Assert::IsTrue(NULL != pPersist, "IVsiPersistSeries missing.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CComVariant vSetSerId(L"01E67435-7295-44D7-8116-483DC5A08102");
			hr = pSeries->SetProperty(VSI_PROP_SERIES_ID, &vSetSerId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComVariant vSetStuId(L"A94E4564-14B5-46AE-8231-B1B1954A77D8");
			hr = pSeries->SetProperty(VSI_PROP_SERIES_ID_STUDY, &vSetStuId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComVariant vSetName(L"SeriesNew");
			hr = pSeries->SetProperty(VSI_PROP_SERIES_NAME, &vSetName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			SYSTEMTIME st;
			GetSystemTime(&st);

			COleDateTime date(st);

			CComVariant vSetDate;
			V_VT(&vSetDate) = VT_DATE;
			V_DATE(&vSetDate) = DATE(date);

			hr = pSeries->SetProperty(VSI_PROP_SERIES_CREATED_DATE, &vSetDate);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComVariant vSetAcqBy(L"Sierra Developer");
			hr = pSeries->SetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vSetAcqBy);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComVariant vSetVer(L"1.0");
			hr = pSeries->SetProperty(VSI_PROP_SERIES_VERSION_REQUIRED, &vSetVer);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CString strPath;
			strPath.Format(L"%s\\TestSaveSeries.vxml", pszTestDir);
			hr = pPersist->SaveSeriesData(strPath, VSI_DATA_TYPE_SERIES_LIST);
			Assert::AreEqual(S_OK, hr, "SaveSeriesData failed.");

			CComPtr<IVsiSeries> pSeriesLoaded;

			hr = pSeriesLoaded.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
			
			CComQIPtr<IVsiPersistSeries> pPersistLoaded(pSeriesLoaded);
			Assert::IsTrue(NULL != pPersistLoaded, "IVsiPersistSeries missing.");

			hr = pPersistLoaded->LoadSeriesData(strPath, VSI_DATA_TYPE_SERIES_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadSeriesData failed.");

			CComVariant vGetSerId;
			hr = pSeriesLoaded->GetProperty(VSI_PROP_SERIES_ID, &vGetSerId);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetSerId == vGetSerId, "Property value not match.");

			CComVariant vGetStuId;
			hr = pSeriesLoaded->GetProperty(VSI_PROP_SERIES_ID_STUDY, &vGetStuId);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetStuId == vGetStuId, "Property value not match.");

			CComVariant vGetName;
			hr = pSeriesLoaded->GetProperty(VSI_PROP_SERIES_NAME, &vGetName);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetName == vGetName, "Property value not match.");

			CComVariant vGetDate;
			hr = pSeriesLoaded->GetProperty(VSI_PROP_SERIES_CREATED_DATE, &vGetDate);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetDate == vGetDate, "Property value not match.");

			CComVariant vGetAcqBy;
			hr = pSeriesLoaded->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vGetAcqBy);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetAcqBy == vGetAcqBy, "Property value not match.");

			CComVariant vGetVer;
			hr = pSeriesLoaded->GetProperty(VSI_PROP_SERIES_VERSION_REQUIRED, &vGetVer);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetVer == vGetVer, "Property value not match.");

			// Clean Up
			String^ strDirTest = gcnew String(strDir);
			strDirTest += "\\TestSaveSeries.vxml";
			File::Delete(strDirTest);
		}

		/// <summary>
		/// <step>Create a Series object</step>
		/// <step>LoadPropertyLabel from a test file<results>The call must succeed</results></step>
		/// <step>GetPropertyLabel for all properties against expected values<results>The call must succeed and the properties must match</results></step>
		/// </summary>
		[TestMethod]
		void TestLoadPropertyLabel()
		{
			CComPtr<IVsiSeries> pSeries;

			HRESULT hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistSeries> pPersist(pSeries);
			Assert::IsTrue(NULL != pPersist, "Get IVsiPersistSeries failed.");

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestSeriesFields.xml", pszTestDir);

			hr = pPersist->LoadPropertyLabel(strPath);
			Assert::AreEqual(S_OK, hr, "LoadPropertyLabel failed.");

			CComHeapPtr<OLECHAR> pszSetLabel1(L"Anesthetic Type");
			CComHeapPtr<OLECHAR> pszGetLabel1;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM1, &pszGetLabel1);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel1, pszGetLabel1), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel2(L"Anesthetic On");
			CComHeapPtr<OLECHAR> pszGetLabel2;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM2, &pszGetLabel2);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel2, pszGetLabel2), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel3(L"Anesthetic Off");
			CComHeapPtr<OLECHAR> pszGetLabel3;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM3, &pszGetLabel3);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel3, pszGetLabel3), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel4(L"Protocol ID");
			CComHeapPtr<OLECHAR> pszGetLabel4;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM4, &pszGetLabel4);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel4, pszGetLabel4), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel5(L"Protocol Name");
			CComHeapPtr<OLECHAR> pszGetLabel5;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM5, &pszGetLabel5);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel5, pszGetLabel5), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel6(L"Injectable");
			CComHeapPtr<OLECHAR> pszGetLabel6;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM6, &pszGetLabel6);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel6, pszGetLabel6), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel7(L"Injection Site");
			CComHeapPtr<OLECHAR> pszGetLabel7;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM7, &pszGetLabel7);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel7, pszGetLabel7), "Property value not match.");

			CComHeapPtr<OLECHAR> pszSetLabel8(L"Injection Amount");
			CComHeapPtr<OLECHAR> pszGetLabel8;
			hr = pSeries->GetPropertyLabel(VSI_PROP_SERIES_CUSTOM8, &pszGetLabel8);
			Assert::AreEqual(S_OK, hr, "GetPropertyLabel failed.");
			Assert::IsTrue(0 == lstrcmp(pszSetLabel8, pszGetLabel8), "Property value not match.");
		}

		//test methods in IVsiEnumSeries

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
