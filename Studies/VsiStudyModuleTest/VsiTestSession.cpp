/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiTestSession.cpp
**
**	Description:
**		VsiSession Test
**
********************************************************************************/
#include "stdafx.h"
#include <VsiUnitTestHelper.h>
#include <VsiServiceKey.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiStudyModule.h>
#include <VsiBModeModule.h>
#include <VsiModeModule.h>

using namespace System;
using namespace System::Text;
using namespace System::Collections::Generic;
using namespace Microsoft::VisualStudio::TestTools::UnitTesting;
using namespace System::IO;

namespace VsiStudyModuleTest
{
	[TestClass]
	public ref class VsiTestSession
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

		// Test IVsiSession
		[TestMethod]
		void TestCreate()
		{
			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
		};

		/// <summary>
		/// <step>Initializes Session</step>
		/// </summary>
		/// <results>Session initialized successfully</results>
		[TestMethod]
		void TestInitializeUninitialize()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		};

		/// <summary>
		/// <step>SetSessionMode</step>
		/// <step>GetSessionMode</step>
		/// </summary>
		/// <results>Session Modes are identical</results>
		[TestMethod]
		void TestSetGetSessionMode()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Create Session Mode quick
			VSI_SESSION_MODE pSessionModeIn(VSI_SESSION_MODE_QUICK);

			hr = pSession->SetSessionMode(pSessionModeIn);
			Assert::AreEqual(S_OK, hr, "SetSessionMode failed.");

			VSI_SESSION_MODE pSessionModeOut(VSI_SESSION_MODE_NONE);
			hr = pSession->GetSessionMode(&pSessionModeOut);
			Assert::AreEqual(S_OK, hr, "GetSessionMode failed.");

			Assert::IsTrue(pSessionModeIn == pSessionModeOut, "GetSessionMode failed.");

			// Create Session Mode new study
			pSessionModeIn = VSI_SESSION_MODE_NEW_STUDY;

			hr = pSession->SetSessionMode(pSessionModeIn);
			Assert::AreEqual(S_OK, hr, "SetSessionMode failed.");

			hr = pSession->GetSessionMode(&pSessionModeOut);
			Assert::AreEqual(S_OK, hr, "GetSessionMode failed.");

			Assert::IsTrue(pSessionModeIn == pSessionModeOut, "GetSessionMode failed.");

			// Create Session Mode new series
			pSessionModeIn = VSI_SESSION_MODE_NEW_SERIES;

			hr = pSession->SetSessionMode(pSessionModeIn);
			Assert::AreEqual(S_OK, hr, "SetSessionMode failed.");

			hr = pSession->GetSessionMode(&pSessionModeOut);
			Assert::AreEqual(S_OK, hr, "GetSessionMode failed.");

			Assert::IsTrue(pSessionModeIn == pSessionModeOut, "GetSessionMode failed.");

			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}
		
		/// <summary>
		/// <step>SetPrimaryStudy</step>
		/// <step>GetPrimaryStudy</step>
		/// </summary>
		/// <results>Studies are identical</results>
		[TestMethod]
		void TestSetGetPrimaryStudy()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Create study
			CComPtr<IVsiStudy> pStudyIn;
			hr = pStudyIn.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->SetPrimaryStudy(pStudyIn, FALSE);
			Assert::AreEqual(S_OK, hr, "SetPrimaryStudy failed.");

			CComPtr<IVsiStudy> pStudyOut;
			hr = pSession->GetPrimaryStudy(&pStudyOut);
			Assert::AreEqual(S_OK, hr, "GetPrimaryStudy failed.");

			Assert::IsTrue(pStudyIn == pStudyOut, "GetPrimaryStudy failed.");

			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>SetPrimarySeries</step>
		/// <step>GetPrimarySeries</step>
		/// </summary>
		/// <results>Series are identical</results>
		[TestMethod]
		void TestSetGetPrimarySeries()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Create series
			CComPtr<IVsiSeries> pSeriesIn;
			hr = pSeriesIn.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->SetPrimarySeries(pSeriesIn, FALSE);
			Assert::AreEqual(S_OK, hr, "SetPrimaryStudy failed.");

			CComPtr<IVsiSeries> pSeriesOut;
			hr = pSession->GetPrimarySeries(&pSeriesOut);
			Assert::AreEqual(S_OK, hr, "GetPrimarySeries failed.");

			Assert::IsTrue(pSeriesIn == pSeriesOut, "GetPrimarySeries failed.");

			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Gets active session slot <results>Active slot is 0</results></step>
		/// <step>Sets active session slot to 1</step>
		/// <step>Gets active session slot <results>Active slot is 1</results></step>
		/// <step>Sets active session slot to 0</step>
		/// <step>Gets active session slot <results>Active slot is 0</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetActiveSlot()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			int iSlot = -1;

			hr = pSession->GetActiveSlot(&iSlot);
			Assert::AreEqual(S_OK, hr, "GetActiveSlot failed.");
			Assert::AreEqual(0, iSlot, "GetActiveSlot failed.");

			// To 1
			hr = pSession->SetActiveSlot(1);
			Assert::AreEqual(S_OK, hr, "SetActiveSlot failed.");

			hr = pSession->GetActiveSlot(&iSlot);
			Assert::AreEqual(S_OK, hr, "GetActiveSlot failed.");
			Assert::AreEqual(1, iSlot, "GetActiveSlot failed.");

			// To 0
			hr = pSession->SetActiveSlot(0);
			Assert::AreEqual(S_OK, hr, "SetActiveSlot failed.");

			hr = pSession->GetActiveSlot(&iSlot);
			Assert::AreEqual(S_OK, hr, "GetActiveSlot failed.");
			Assert::AreEqual(0, iSlot, "GetActiveSlot failed.");

			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>For each slot</step>
		/// <step>SetStudy</step>
		/// <step>GetStudy<results>Studies are identical</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetStudy()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Create study
			CComPtr<IVsiStudy> pStudyIn;
			hr = pStudyIn.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
			{
				hr = pSession->SetStudy(i, pStudyIn, FALSE);
				Assert::AreEqual(S_OK, hr, "SetStudy failed.");

				CComPtr<IVsiStudy> pStudyOut;
				hr = pSession->GetStudy(i, &pStudyOut);
				Assert::AreEqual(S_OK, hr, "GetStudy failed.");

				Assert::IsTrue(pStudyIn == pStudyOut, "GetStudy failed.");
			}

			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>For each slot</step>
		/// <step>SetSeries</step>
		/// <step>GetSeries<results>Series are identical</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetSeries()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Create series
			CComPtr<IVsiSeries> pSeriesIn;
			hr = pSeriesIn.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
			{
				hr = pSession->SetSeries(i, pSeriesIn, FALSE);
				Assert::AreEqual(S_OK, hr, "SetSeries failed.");

				CComPtr<IVsiSeries> pSeriesOut;
				hr = pSession->GetSeries(i, &pSeriesOut);
				Assert::AreEqual(S_OK, hr, "GetSeries failed.");

				Assert::IsTrue(pSeriesIn == pSeriesOut, "GetSeries failed.");
			}

			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>For each slot</step>
		/// <step>SetImage</step>
		/// <step>GetImage<results>Images are identical</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetImage()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Create image
			CComPtr<IVsiImage> pImageIn;
			hr = pImageIn.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComVariant vNsImage(L"1234/1234/1234");
			hr = pImageIn->SetProperty(VSI_PROP_IMAGE_NS, &vNsImage);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
			{
				hr = pSession->SetImage(i, pImageIn, FALSE, FALSE);
				Assert::AreEqual(S_OK, hr, "SetImage failed.");

				CComPtr<IVsiImage> pImageOut;
				hr = pSession->GetImage(i, &pImageOut);
				Assert::AreEqual(S_OK, hr, "GetImage failed.");

				Assert::IsTrue(pImageIn == pImageOut, "GetImage failed.");
			}

			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		// After we make TestSaveImageData() from VsiTestImage.cpp works, then we can make this test case work.
		///// <summary>
		///// <step>For each slot</step>
		///// <step>SetMode</step>
		///// <step>GetMode<results>Modes are identical</results></step>
		///// </summary>
		//[TestMethod]
		//void TestSetGetMode()
		//{
		//	CComPtr<IVsiApp> pApp;
		//	CreateApp(&pApp);
		//	Assert::IsTrue(pApp != NULL, "Create App failed.");

		//	CComPtr<IVsiSession> pSession;

		//	HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
		//	Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

		//	hr = pSession->Initialize(pApp);
		//	Assert::AreEqual(S_OK, hr, "Initialize failed.");

		//	// Create Mode
		//	CComPtr<IVsiBMode> pBMode;

		//	hr = pBMode.CoCreateInstance(__uuidof(VsiBMode));
		//	Assert::AreEqual(S_OK, hr, "B-Mode CoCreateInstance failed.");

		//	CComQIPtr<IVsiMode> pModeIn(pBMode);
		//	Assert::IsTrue(pModeIn != NULL, "B-Mode initialize failed.");

		//	hr = pModeIn->Initialize(pApp, VSI_MODE_INIT_CREATE_DATA_VIEW);
		//	Assert::AreEqual(S_OK, hr, "B-Mode CoCreateInstance failed."); // it will returns E_FAIL here, it need to initialize much more things, need future development
		//	
		//	for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
		//	{
		//		hr = pSession->SetMode(i, pModeIn);
		//		Assert::AreEqual(S_OK, hr, "SetMode failed.");

		//		CComPtr<IUnknown> pModeOut;
		//		hr = pSession->GetMode(i, &pModeOut);
		//		Assert::AreEqual(S_OK, hr, "GetMode failed.");

		//		Assert::IsTrue(pModeIn == pModeOut, "GetMode failed.");
		//	}

		//	hr = pSession->Uninitialize();
		//	Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
		//	
		//	CloseApp(pApp);
		//}

		/// <summary>
		/// <step>For each slot</step>
		/// <step>SetModeView</step>
		/// <step>GetModeView<results>Mode views are identical</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetModeView()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Create Mode View
			CComPtr<IVsiBModeView> pBModeView;

			hr = pBModeView.CoCreateInstance(__uuidof(VsiBModeView));
			Assert::AreEqual(S_OK, hr, "B-Mode view CoCreateInstance failed.");

			CComQIPtr<IVsiModeView> pModeViewIn(pBModeView);
			Assert::IsTrue(pModeViewIn != NULL, "B-Mode view initialize failed.");

			for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
			{
				hr = pSession->SetModeView(i, pModeViewIn);
				Assert::AreEqual(S_OK, hr, "SetModeView failed.");

				CComPtr<IUnknown> pModeViewOut;
				hr = pSession->GetModeView(i, &pModeViewOut);
				Assert::AreEqual(S_OK, hr, "GetModeView failed.");

				Assert::IsTrue(pModeViewIn == pModeViewOut, "GetModeView failed.");
			}

			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create a Study object</step>
		/// <step>Load a study from a test file</step>
		/// <step>Create a Session object</step>
		/// <step>Connect the study and the session</step>
		/// <step>Check if the study is loaded from the session<results>piSlot value should be the same as set</results></step>
		/// </summary>
		[TestMethod]
		void TestGetIsStudyLoaded()
		{
			// Create Study & LoadStudyData
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
			Assert::IsTrue(NULL != pPersist, "Get IVsiPersistStudy failed.");

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestStudy.vxml", pszTestDir);

			hr = pPersist->LoadStudyData(strPath, VSI_DATA_TYPE_STUDY_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadStudyData failed.");

			VSI_STUDY_ERROR dwErrorCode(VSI_STUDY_ERROR_LOAD_XML);
			hr = pStudy->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

			// Create Session & Set Study to Session
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			hr = pSession->SetStudy(0, pStudy, FALSE); //set iSlot == 0
			Assert::AreEqual(S_OK, hr, "SetStudy failed.");

			// GetIsStudyLoaded
			CString strStudyId(L"0123456789ABCDEFGHIJKLMON");
			int iSlot(-1);

			hr = pSession->GetIsStudyLoaded(strStudyId, &iSlot);
			Assert::AreEqual(S_OK, hr, "GetIsStudyLoaded failed.");
			Assert::IsTrue(iSlot == 0, "GetIsStudyLoaded failed.");

			// Clean Up
			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create a Study object</step>
		/// <step>Load a study from a test file</step>
		/// <step>Create a Series object</step>
		/// <step>Load a series from a test file</step>
		/// <step>Create a Session object</step>
		/// <step>Connect the series & the study & the session</step>
		/// <step>Check if the series is loaded from the session<results>piSlot value should be the same as set</results></step>
		/// </summary>
		[TestMethod]
		void TestGetIsSeriesLoaded()
		{
			// Create Study & LoadStudyData
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistStudy> pPersistStudy(pStudy);
			Assert::IsTrue(NULL != pPersistStudy, "Get IVsiPersistStudy failed.");

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestStudy.vxml", pszTestDir);

			hr = pPersistStudy->LoadStudyData(strPath, VSI_DATA_TYPE_STUDY_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadStudyData failed.");

			VSI_STUDY_ERROR dwStudyErrorCode(VSI_STUDY_ERROR_LOAD_XML);
			hr = pStudy->GetErrorCode(&dwStudyErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwStudyErrorCode, "GetErrorCode failed.");

			// Create Series & LoadSeriesData
			CComPtr<IVsiSeries> pSeries;

			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistSeries> pPersistSeries(pSeries);
			Assert::IsTrue(NULL != pPersistSeries, "Get IVsiPersistStudy failed.");

			strPath.Format(L"%s\\UnitTestSeries.vxml", pszTestDir);

			hr = pPersistSeries->LoadSeriesData(strPath, VSI_DATA_TYPE_SERIES_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadSeriesData failed.");

			VSI_SERIES_ERROR dwSeriesErrorCode(VSI_SERIES_ERROR_LOAD_XML);
			hr = pSeries->GetErrorCode(&dwSeriesErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwSeriesErrorCode, "GetErrorCode failed.");

			// Create Session
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Connect the series & the study & the session
			hr = pSeries->SetStudy(pStudy);
			Assert::AreEqual(S_OK, hr, "SetStudy failed.");

			hr = pSession->SetSeries(0, pSeries, FALSE); //set iSlot == 0 //no need to pSession->SetStudy
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			// GetIsSeriesLoaded
			CString strStudySeriesId(L"0123456789ABCDEFGHIJKLMON/0123456789ABCDEFGHIJKLMON");
			int iSlot(-1);

			hr = pSession->GetIsSeriesLoaded(strStudySeriesId, &iSlot);
			Assert::AreEqual(S_OK, hr, "GetIsSeriesLoaded failed.");
			Assert::IsTrue(iSlot == 0, "GetIsSeriesLoaded failed.");

			// Clean Up
			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create a Study object</step>
		/// <step>Load a study from a test file</step>
		/// <step>Create a Series object</step>
		/// <step>Load a series from a test file</step>
		/// <step>Create a Image object</step>
		/// <step>Load a series from a test file</step>
		/// <step>Create a Session object</step>
		/// <step>Connect the image & the series & the study & the session</step>
		/// <step>Check if the image is loaded from the session<results>piSlot value should be the same as set</results></step>
		/// </summary>
		[TestMethod]
		void TestGetIsImageLoaded()
		{
			// Create Study & LoadStudyData
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistStudy> pPersistStudy(pStudy);
			Assert::IsTrue(NULL != pPersistStudy, "Get IVsiPersistStudy failed.");

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestStudy.vxml", pszTestDir);

			hr = pPersistStudy->LoadStudyData(strPath, VSI_DATA_TYPE_STUDY_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadStudyData failed.");

			VSI_STUDY_ERROR dwStudyErrorCode(VSI_STUDY_ERROR_LOAD_XML);
			hr = pStudy->GetErrorCode(&dwStudyErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwStudyErrorCode, "GetErrorCode failed.");

			// Create Series & LoadSeriesData
			CComPtr<IVsiSeries> pSeries;

			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistSeries> pPersistSeries(pSeries);
			Assert::IsTrue(NULL != pPersistSeries, "Get IVsiPersistStudy failed.");

			strPath.Format(L"%s\\UnitTestSeries.vxml", pszTestDir);

			hr = pPersistSeries->LoadSeriesData(strPath, VSI_DATA_TYPE_SERIES_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadSeriesData failed.");

			VSI_SERIES_ERROR dwSeriesErrorCode(VSI_SERIES_ERROR_LOAD_XML);
			hr = pSeries->GetErrorCode(&dwSeriesErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwSeriesErrorCode, "GetErrorCode failed.");

			// Create Image & LoadImageData
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiStudyManager> pStudyMgr;

			hr = pStudyMgr.CoCreateInstance(__uuidof(VsiStudyManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pStudyMgr->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pImage->SetSeries(pSeries); //Need to Connect the image and the series here, otherwise LoadImageData will return E_FAIL
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			CComQIPtr<IVsiPersistImage> pPersistImage(pImage);
			Assert::IsTrue(NULL != pPersistImage, "IVsiPersistImage missing.");

			strPath.Format(L"%s\\UnitTestImage.vxml", pszTestDir);

			hr = pPersistImage->LoadImageData(strPath, NULL, NULL, 0);
			Assert::AreEqual(S_OK, hr, "LoadImageData failed.");

			VSI_IMAGE_ERROR dwImageErrorCode(VSI_IMAGE_ERROR_LOAD_XML);
			hr = pImage->GetErrorCode(&dwImageErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_IMAGE_ERROR_NONE == dwImageErrorCode, "GetErrorCode failed.");

			// Create Session
			CComPtr<IVsiSession> pSession;

			hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Connect the series & the study & the session
			// the image and the series is already connected at the "Create Image & LoadImageData" section

			hr = pSeries->SetStudy(pStudy);
			Assert::AreEqual(S_OK, hr, "SetStudy failed.");

			hr = pSession->SetSeries(0, pSeries, FALSE); //set iSlot == 0 //no need to pSession->SetStudy
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			hr = pSession->SetImage(0, pImage, FALSE, FALSE);  //set iSlot == 0 //must pSession->SetImage
			Assert::AreEqual(S_OK, hr, "SetImage failed.");

			// GetIsSeriesLoaded
			CString strStudySeriesImageId(L"0123456789ABCDEFGHIJKLMON/0123456789ABCDEFGHIJKLMON/0123456789ABCDEFGHIJKLMON");
			int iSlot(-1);

			hr = pSession->GetIsImageLoaded(strStudySeriesImageId, &iSlot);
			Assert::AreEqual(S_OK, hr, "GetIsImageLoaded failed.");
			Assert::IsTrue(iSlot == 0, "GetIsImageLoaded failed.");

			// Clean Up
			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create a Study object</step>
		/// <step>Load a study from a test file</step>
		/// <step>Create a Series object</step>
		/// <step>Load a series from a test file</step>
		/// <step>Create a Image object</step>
		/// <step>Load a series from a test file</step>
		/// <step>Create a Session object</step>
		/// <step>Connect the image & the series & the study & the session</step>
		/// <step>Check if the image item is loaded from the session<results>piSlot value should be the same as set</results></step>
		/// </summary>
		[TestMethod]
		void TestGetIsItemLoaded()
		{
			// Create Study & LoadStudyData
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistStudy> pPersistStudy(pStudy);
			Assert::IsTrue(NULL != pPersistStudy, "Get IVsiPersistStudy failed.");

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestStudy.vxml", pszTestDir);

			hr = pPersistStudy->LoadStudyData(strPath, VSI_DATA_TYPE_STUDY_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadStudyData failed.");

			VSI_STUDY_ERROR dwStudyErrorCode(VSI_STUDY_ERROR_LOAD_XML);
			hr = pStudy->GetErrorCode(&dwStudyErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwStudyErrorCode, "GetErrorCode failed.");

			// Create Series & LoadSeriesData
			CComPtr<IVsiSeries> pSeries;

			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistSeries> pPersistSeries(pSeries);
			Assert::IsTrue(NULL != pPersistSeries, "Get IVsiPersistStudy failed.");

			strPath.Format(L"%s\\UnitTestSeries.vxml", pszTestDir);

			hr = pPersistSeries->LoadSeriesData(strPath, VSI_DATA_TYPE_SERIES_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadSeriesData failed.");

			VSI_SERIES_ERROR dwSeriesErrorCode(VSI_SERIES_ERROR_LOAD_XML);
			hr = pSeries->GetErrorCode(&dwSeriesErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwSeriesErrorCode, "GetErrorCode failed.");

			// Create Image & LoadImageData
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiStudyManager> pStudyMgr;

			hr = pStudyMgr.CoCreateInstance(__uuidof(VsiStudyManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pStudyMgr->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pImage->SetSeries(pSeries); //Need to connect the image and the series here, otherwise LoadImageData will return E_FAIL
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			CComQIPtr<IVsiPersistImage> pPersistImage(pImage);
			Assert::IsTrue(NULL != pPersistImage, "IVsiPersistImage missing.");

			strPath.Format(L"%s\\UnitTestImage.vxml", pszTestDir);

			hr = pPersistImage->LoadImageData(strPath, NULL, NULL, 0);
			Assert::AreEqual(S_OK, hr, "LoadImageData failed.");

			VSI_IMAGE_ERROR dwImageErrorCode(VSI_IMAGE_ERROR_LOAD_XML);
			hr = pImage->GetErrorCode(&dwImageErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_IMAGE_ERROR_NONE == dwImageErrorCode, "GetErrorCode failed.");

			// Create Session
			CComPtr<IVsiSession> pSession;

			hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Connect the series & the study & the session
			// the image and the series is already connected at the "Create Image & LoadImageData" section

			hr = pSeries->SetStudy(pStudy);
			Assert::AreEqual(S_OK, hr, "SetStudy failed.");

			hr = pSession->SetSeries(0, pSeries, FALSE); //set iSlot == 0 //no need to pSession->SetStudy
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			hr = pSession->SetImage(0, pImage, FALSE, FALSE);  //set iSlot == 0 //must pSession->SetImage
			Assert::AreEqual(S_OK, hr, "SetImage failed.");

			// GetIsItemLoaded
			CString strStudySeriesImageId(L"0123456789ABCDEFGHIJKLMON/0123456789ABCDEFGHIJKLMON/0123456789ABCDEFGHIJKLMON");
			int iSlot(-1);

			hr = pSession->GetIsItemLoaded(strStudySeriesImageId, &iSlot);
			Assert::AreEqual(S_OK, hr, "GetIsItemLoaded failed.");
			Assert::IsTrue(iSlot == 0, "GetIsItemLoaded failed.");

			// Clean Up
			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create a Study object</step>
		/// <step>Load a study from a test file</step>
		/// <step>Create a Series object</step>
		/// <step>Load a series from a test file</step>
		/// <step>Create a Image object</step>
		/// <step>Load a series from a test file</step>
		/// <step>Create a Session object</step>
		/// <step>Connect the image & the series & the study & the session</step>
		/// <step>Check if the image is reviewed in the session<results>S_OK: Loaded before; S_FALSE: Not reviewed before</results></step>
		/// </summary>
		[TestMethod]
		void TestGetIsItemReviewed()
		{
			// Create Study & LoadStudyData
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistStudy> pPersistStudy(pStudy);
			Assert::IsTrue(NULL != pPersistStudy, "Get IVsiPersistStudy failed.");

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestStudy.vxml", pszTestDir);

			hr = pPersistStudy->LoadStudyData(strPath, VSI_DATA_TYPE_STUDY_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadStudyData failed.");

			VSI_STUDY_ERROR dwStudyErrorCode(VSI_STUDY_ERROR_LOAD_XML);
			hr = pStudy->GetErrorCode(&dwStudyErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwStudyErrorCode, "GetErrorCode failed.");

			// Create Series & LoadSeriesData
			CComPtr<IVsiSeries> pSeries;

			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistSeries> pPersistSeries(pSeries);
			Assert::IsTrue(NULL != pPersistSeries, "Get IVsiPersistStudy failed.");

			strPath.Format(L"%s\\UnitTestSeries.vxml", pszTestDir);

			hr = pPersistSeries->LoadSeriesData(strPath, VSI_DATA_TYPE_SERIES_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadSeriesData failed.");

			VSI_SERIES_ERROR dwSeriesErrorCode(VSI_SERIES_ERROR_LOAD_XML);
			hr = pSeries->GetErrorCode(&dwSeriesErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_SERIES_ERROR_NONE == dwSeriesErrorCode, "GetErrorCode failed.");

			// Create Image & LoadImageData
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiStudyManager> pStudyMgr;

			hr = pStudyMgr.CoCreateInstance(__uuidof(VsiStudyManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pStudyMgr->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComPtr<IVsiImage> pImage;

			hr = pImage.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pImage->SetSeries(pSeries); //Need to connect the image and the series here, otherwise LoadImageData will return E_FAIL
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			CComQIPtr<IVsiPersistImage> pPersistImage(pImage);
			Assert::IsTrue(NULL != pPersistImage, "IVsiPersistImage missing.");

			strPath.Format(L"%s\\UnitTestImage.vxml", pszTestDir);

			hr = pPersistImage->LoadImageData(strPath, NULL, NULL, 0);
			Assert::AreEqual(S_OK, hr, "LoadImageData failed.");

			VSI_IMAGE_ERROR dwImageErrorCode(VSI_IMAGE_ERROR_LOAD_XML);
			hr = pImage->GetErrorCode(&dwImageErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_IMAGE_ERROR_NONE == dwImageErrorCode, "GetErrorCode failed.");

			// Create Session
			CComPtr<IVsiSession> pSession;

			hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Connect the series & the study & the session
			// the image and the series is already connected at the "Create Image & LoadImageData" section

			hr = pSeries->SetStudy(pStudy);
			Assert::AreEqual(S_OK, hr, "SetStudy failed.");

			hr = pSession->SetSeries(0, pSeries, FALSE); //set iSlot == 0 //no need to pSession->SetStudy
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			// situation 1: set image as not reviewed
			hr = pSession->SetImage(0, pImage, FALSE, FALSE);  //set iSlot == 0 //must pSession->SetImage
			Assert::AreEqual(S_OK, hr, "SetImage failed.");

			// GetIsItemReviewed
			CString strStudySeriesImageId(L"0123456789ABCDEFGHIJKLMON/0123456789ABCDEFGHIJKLMON/0123456789ABCDEFGHIJKLMON");
			int iSlot(-1);

			hr = pSession->GetIsItemReviewed(strStudySeriesImageId);
			Assert::AreEqual(S_FALSE, hr, "GetIsItemReviewed failed."); //S_OK: Loaded before; S_FALSE: Not reviewed before

			// situation 2: set image as reviewed
			hr = pSession->ReleaseSlot(0);  //release slot before test situation 2
			//hr = pSession->SetImage(0, NULL, FALSE, FALSE); //same as hr = pSession->ReleaseSlot(0);

			hr = pSession->SetImage(0, pImage, TRUE, FALSE);
			Assert::AreEqual(S_OK, hr, "SetImage failed.");

			// GetIsItemReviewed
			hr = pSession->GetIsItemReviewed(strStudySeriesImageId);
			Assert::AreEqual(S_OK, hr, "GetIsItemReviewed failed."); //S_OK: Loaded before; S_FALSE: Not reviewed before

			// Clean Up
			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create a session object</step>
		/// <step>SetModeView</step>
		/// <step>GetModeView<results>returns S_OK and Mode views are identical</results></step>
		/// <step>ReleaseSlot</step>
		/// <step>GetModeView<results>returns S_FALSE</results></step>
		/// <step>SetMode</step> // after we make TestSetGetMode work, then we can make this part work
		/// <step>GetMode<results>returns S_OK and Modes are identical</results></step>
		/// <step>ReleaseSlot</step>
		/// <step>GetMode<results>returns S_FALSE</results></step>
		/// <step>SetImage</step>
		/// <step>GetImage<results>returns S_OK and Images are identical</results></step>
		/// <step>ReleaseSlot</step>
		/// <step>GetImage<results>returns S_FALSE</results></step>
		/// <step>SetSeries</step>
		/// <step>GetSeries<results>returns S_OK and Series are identical</results></step>
		/// <step>ReleaseSlot</step>
		/// <step>GetSeries<results>returns S_FALSE</results></step>
		/// <step>SetStudy</step>
		/// <step>GetStudy<results>returns S_OK and Studies are identical</results></step>
		/// <step>ReleaseSlot</step>
		/// <step>GetStudy<results>returns S_FALSE</results></step>
		/// </summary>
		[TestMethod]
		void TestReleaseSlot()
		{
			// Create Session
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Create Mode View and set it to the session
			CComPtr<IVsiBModeView> pBModeView;

			hr = pBModeView.CoCreateInstance(__uuidof(VsiBModeView));
			Assert::AreEqual(S_OK, hr, "B-Mode view CoCreateInstance failed.");

			CComQIPtr<IVsiModeView> pModeViewIn(pBModeView);
			Assert::IsTrue(pModeViewIn != NULL, "B-Mode view initialize failed.");

			hr = pSession->SetModeView(0, pModeViewIn);
			Assert::AreEqual(S_OK, hr, "SetModeView failed.");

			// GetModeView returns S_OK
			CComPtr<IUnknown> pModeViewOut;
			hr = pSession->GetModeView(0, &pModeViewOut);
			Assert::AreEqual(S_OK, hr, "GetModeView failed.");
			Assert::IsTrue(pModeViewIn == pModeViewOut, "GetModeView failed.");

			// Release Slot
			hr = pSession->ReleaseSlot(0);

			// GetModeView returns S_FALSE
			pModeViewOut.Release();
			hr = pSession->GetModeView(0, &pModeViewOut);
			Assert::AreEqual(S_FALSE, hr, "GetModeView succeed which is unexpected.");

			// after we make TestSetGetMode work, then we can make this part work
			/// <step>SetMode</step> 
			/// <step>GetMode<results>returns S_OK and Modes are identical</results></step>
			/// <step>ReleaseSlot</step>
			/// <step>GetMode<results>returns S_FALSE</results></step>

			// Create image and set it to the session
			CComPtr<IVsiImage> pImageIn;
			hr = pImageIn.CoCreateInstance(__uuidof(VsiImage));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComVariant vNsImage(L"1234/1234/1234");
			hr = pImageIn->SetProperty(VSI_PROP_IMAGE_NS, &vNsImage);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			hr = pSession->SetImage(1, pImageIn, FALSE, FALSE);
			Assert::AreEqual(S_OK, hr, "SetImage failed.");

			// GetImage returns S_OK
			CComPtr<IVsiImage> pImageOut;
			hr = pSession->GetImage(1, &pImageOut);
			Assert::AreEqual(S_OK, hr, "GetImage failed.");
			Assert::IsTrue(pImageIn == pImageOut, "GetImage failed.");

			// Release Slot
			hr = pSession->ReleaseSlot(1);

			// GetImage returns S_FALSE
			pImageOut.Release();
			hr = pSession->GetImage(1, &pImageOut);
			Assert::AreEqual(S_FALSE, hr, "GetImage succeed which is unexpected.");

			// Create series and set it to the session
			CComPtr<IVsiSeries> pSeriesIn;
			hr = pSeriesIn.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->SetSeries(0, pSeriesIn, FALSE);
			Assert::AreEqual(S_OK, hr, "SetSeries failed.");

			// GetSeries returns S_OK
			CComPtr<IVsiSeries> pSeriesOut;
			hr = pSession->GetSeries(0, &pSeriesOut);
			Assert::AreEqual(S_OK, hr, "GetSeries failed.");
			Assert::IsTrue(pSeriesIn == pSeriesOut, "GetSeries failed.");

			// Release Slot
			hr = pSession->ReleaseSlot(0);

			// GetSeries returns S_FALSE
			pSeriesOut.Release();
			hr = pSession->GetSeries(0, &pSeriesOut);
			Assert::AreEqual(S_FALSE, hr, "GetSeries succeed which is unexpected.");

			// Create study and set it to the session
			CComPtr<IVsiStudy> pStudyIn;
			hr = pStudyIn.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->SetStudy(1, pStudyIn, FALSE);
			Assert::AreEqual(S_OK, hr, "SetStudy failed.");

			// GetStudy returns S_OK
			CComPtr<IVsiStudy> pStudyOut;
			hr = pSession->GetStudy(1, &pStudyOut);
			Assert::AreEqual(S_OK, hr, "GetStudy failed.");
			Assert::IsTrue(pStudyIn == pStudyOut, "GetStudy failed.");

			// Release Slot
			hr = pSession->ReleaseSlot(1);

			// GetStudy returns S_FALSE
			pStudyOut.Release();
			hr = pSession->GetStudy(1, &pStudyOut);
			Assert::AreEqual(S_FALSE, hr, "GetStudy succeed which is unexpected.");

			// Clean up
			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

		/// <summary>
		/// <step>When ActiveSeries.xml does not exist under TestDeploymentDir</step>
		/// <step>CheckActiveSeries<results>hr == S_FALSE</results></step>
		/// <step>Create ActiveSeries.xml under TestDeploymentDir</step>
		/// <step>CheckActiveSeries<results>hr == S_OK</results></step>
		/// </summary>
		[TestMethod]
		void TestCheckActiveSeries()
		{
			// Create Session
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// When ActiveSeries.xml does not exist under TestDeploymentDir
			hr = pSession->CheckActiveSeries(FALSE, NULL, NULL); // strDir is passed through pdm when CreateApp
			Assert::AreEqual(S_FALSE, hr, "CheckActiveSeries failed."); //S_FALSE if ActiveSeries.xml does not exist under strDir

			// Create ActiveSeries.xml under TestDeploymentDir
			String^ strDir = testContextInstance->TestDeploymentDir;
			String^ strDirActiveSeries = gcnew String(strDir);
			strDirActiveSeries += "\\ActiveSeries.xml";
			FileStream^ fs = File::Create(strDirActiveSeries);
			Assert::IsTrue(File::Exists(strDirActiveSeries), "Create ActiveSeries.xml failed.");
			fs->Close();

			String^ strContentsActiveSeries(L"<?xml version=\"1.0\" encoding=\"utf-8\"?>"); // \"for"
			strContentsActiveSeries += "\r\n"; // change line
			strContentsActiveSeries += "<activeSeries version=\"1\" ";
			strContentsActiveSeries += "ns=\"0123456789ABCDEFGHIJKLMON/0123456789ABCDEFGHIJKLMON\" ";
			strContentsActiveSeries += "file=\"";
			strContentsActiveSeries += strDir;
			strContentsActiveSeries += "\\Study\\Series\\Series.vxml\"/>";
			File::WriteAllText(strDirActiveSeries, strContentsActiveSeries);

			hr = pSession->CheckActiveSeries(FALSE, NULL, NULL); // strDir is passed through pdm when CreateApp
			Assert::AreEqual(S_OK, hr, "CheckActiveSeries failed."); //S_OK if ActiveSeries.xml does exist under strDir

			// Clean Up
			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);

			File::Delete(strDirActiveSeries);
		}

		// Test IVsiSessionJournal
		/// <summary>
		/// <step>case VSI_SESSION_EVENT_SERIES_NEW</step>
		/// <step>Record<results>ActiveSeries.xml will be created and its content will be written</results></step>
		/// <step>case VSI_SESSION_EVENT_SERIES_CLOSE</step>
		/// <step>Record<results>ActiveSeries.xml will be deleted</results></step>
		/// </summary>
		[TestMethod]
		void TestRecord()
		{
			// Create Session
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiSession> pSession;

			HRESULT hr = pSession.CoCreateInstance(__uuidof(VsiSession));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pSession->Initialize(pApp);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComQIPtr<IVsiSessionJournal> pSessionJournal(pSession);
			Assert::IsTrue(NULL != pSessionJournal, "IVsiSessionJournal missing.");

			// case VSI_SESSION_EVENT_SERIES_NEW
			String^ strDir = testContextInstance->TestDeploymentDir;
			String^ strDirActiveSeries = gcnew String(strDir);
			strDirActiveSeries += "\\ActiveSeries.xml";
			Assert::IsTrue(!File::Exists(strDirActiveSeries), "ActiveSeries.xml exists which is unexpected.");

			CString strNS(L"0123456789ABCDEFGHIJKLMON/0123456789ABCDEFGHIJKLMON");
			CComVariant vNS(strNS);

			String^ strFiletemp= gcnew String(strDir);
			strFiletemp += "\\Study\\Series\\Series.vxml";
			CString strFile(strFiletemp);
			CComVariant vFile(strFile);

			hr = pSessionJournal->Record(VSI_SESSION_EVENT_SERIES_NEW, &vNS, &vFile); // strDir is passed through pdm when CreateApp
			Assert::AreEqual(S_OK, hr, "Record failed.");
			Assert::IsTrue(File::Exists(strDirActiveSeries), "Record failed.");

			String^ strGetContentsActiveSeries = File::ReadAllText(strDirActiveSeries);
			String^ strExpectedContentsActiveSeries(L"<?xml version=\"1.0\" encoding=\"utf-8\"?>"); // \"for"
			strExpectedContentsActiveSeries += "\r\n"; // change line
			strExpectedContentsActiveSeries += "<activeSeries version=\"1\" ";
			strExpectedContentsActiveSeries += "ns=\"0123456789ABCDEFGHIJKLMON/0123456789ABCDEFGHIJKLMON\" ";
			strExpectedContentsActiveSeries += "file=\"";
			strExpectedContentsActiveSeries += strDir;
			strExpectedContentsActiveSeries += "\\Study\\Series\\Series.vxml\"/>\r\n";
			Assert::AreEqual(strGetContentsActiveSeries, strExpectedContentsActiveSeries, "Record failed.");

			// VSI_SESSION_EVENT_SERIES_CLOSE
			hr = pSessionJournal->Record(VSI_SESSION_EVENT_SERIES_CLOSE, NULL, NULL); // strDir is passed through pdm when CreateApp
			Assert::AreEqual(S_OK, hr, "Record failed.");
			Assert::IsTrue(!File::Exists(strDirActiveSeries), "Record failed.");

			// Clean Up
			hr = pSession->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			
			CloseApp(pApp);
		}

	private:
		void CreateApp(IVsiApp** ppApp)
		{
			CComPtr<IVsiTestApp> pTestApp;

			HRESULT hr = pTestApp.CoCreateInstance(__uuidof(VsiTestApp));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance app failed.");

			hr = pTestApp->InitServices(VSI_TESTAPP_SERVICE_PDM | VSI_TESTAPP_SERVICE_OPERATOR_MGR);
			Assert::AreEqual(S_OK, hr, "InitServices failed.");

			CComQIPtr<IVsiServiceProvider> pServiceProvider(pTestApp);
			Assert::IsTrue(NULL != pServiceProvider, "Get IVsiServiceProvider failed.");

			CComPtr<IVsiPdm> pPdm;
			hr = pServiceProvider->QueryService(
				VSI_SERVICE_PDM,
				__uuidof(IVsiPdm),
				(IUnknown**)&pPdm);
			Assert::AreEqual(S_OK, hr, "Get PDM failed.");

			hr = pPdm->AddRoot(VSI_PDM_ROOT_APP, NULL);
			Assert::AreEqual(S_OK, hr, "AddRoot failed.");

			hr = pPdm->AddGroup(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS);
			Assert::AreEqual(S_OK, hr, "AddGroup failed.");

			CComPtr<IVsiParameter> pParam;
			hr = pParam.CoCreateInstance(__uuidof(VsiParameterBoolean));
			Assert::AreEqual(S_OK, hr, "Create parameter failed.");

			hr = pPdm->AddParameter(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
				VSI_PARAMETER_EVENTS_GENERAL_OPERATOR_LIST_UPDATE,
				pParam);
			Assert::AreEqual(S_OK, hr, "AddParameter failed.");

			hr = pPdm->AddGroup(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS);
			Assert::AreEqual(S_OK, hr, "AddGroup failed.");

			hr = pPdm->AddGroup(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL);
			Assert::AreEqual(S_OK, hr, "AddGroup failed.");

			CComPtr<IVsiParameter> pParamSecureMode;
			hr = pParamSecureMode.CoCreateInstance(__uuidof(VsiParameterBoolean));
			Assert::AreEqual(S_OK, hr, "Create parameter failed.");

			hr = pPdm->AddParameter(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_SECURE_MODE,
				pParamSecureMode);
			Assert::AreEqual(S_OK, hr, "AddParameter failed.");

			CComPtr<IVsiParameter> pParamRange;
            hr = pParamRange.CoCreateInstance(__uuidof(VsiParameterRange));
            Assert::AreEqual(S_OK, hr, "Create parameter failed.");

			hr = pParamRange->SetFlags(VSI_PARAM_FLAG_GLOBAL);
            Assert::AreEqual(S_OK, hr, "Create parameter failed.");

			CComQIPtr<IVsiParameterRange> pRange(pParamRange);
            hr = pRange->SetValueType(VT_BSTR);
            Assert::AreEqual(S_OK, hr, "Create parameter failed.");
			String^ strDir = testContextInstance->TestDeploymentDir;
			CString strPath(strDir);
			CComVariant vDataPath(strPath);
            hr = pRange->SetValue(vDataPath, FALSE);
            Assert::AreEqual(S_OK, hr, "Create parameter failed.");
                         
            hr = pPdm->AddParameter(
                VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
                VSI_PARAMETER_SYS_PATH_DATA,
                pParamRange);
            Assert::AreEqual(S_OK, hr, "AddParameter failed.");

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
