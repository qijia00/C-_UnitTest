/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiTestStudy.cpp
**
**	Description:
**		VsiStudy Test
**
********************************************************************************/

#include "stdafx.h"
#include <ATLComTime.h>
#include <VsiUnitTestHelper.h>
#include <VsiStudyModule.h>

using namespace System;
using namespace System::Text;
using namespace System::Collections::Generic;
using namespace	Microsoft::VisualStudio::TestTools::UnitTesting;
using namespace System::IO;

namespace VsiStudyModuleTest
{
	[TestClass]
	public ref class VsiTestStudy
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

		//test methods in IVsiStudy
		[TestMethod]
		void TestCreate()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
		};

		/// <summary>
		/// <step>Create a Study with following attributes:</step>
		/// <step>Data path: "Test\UnitTestStudy.xml"</step>
		/// <step>Study ID: "01234567890123456789"</step>
		/// <step>Study name: "Study 0123456789"</step>
		/// <step>Clone study</step>
		/// <step>Compare cloned study attributes<results>Study attributes are identical</results></step>
		/// </summary>
		[TestMethod]
		void TestClone()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestStudy.vxml", pszTestDir);

			hr = pStudy->SetDataPath(strPath);
			Assert::AreEqual(S_OK, hr, "SetDataPath failed.");

			// ID
			CComVariant vSetId(L"01234567890123456789");
			hr = pStudy->SetProperty(VSI_PROP_STUDY_ID, &vSetId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Name
			CComVariant vSetName(L"Study 0123456789");
			hr = pStudy->SetProperty(VSI_PROP_STUDY_NAME, &vSetName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Clone
			CComPtr<IVsiStudy> pStudyOut;
			hr = pStudy->Clone(&pStudyOut);
			Assert::AreEqual(S_OK, hr, "Clone failed.");

			// Path
			CComHeapPtr<OLECHAR> pszGetPath;
			hr = pStudyOut->GetDataPath(&pszGetPath);
			Assert::AreEqual(S_OK, hr, "GetDataPath failed.");
			CString strGetPath(pszGetPath);
			Assert::IsTrue(strPath == strGetPath, "Clone failed.");

			// ID
			CComVariant vGetId;
			hr = pStudyOut->GetProperty(VSI_PROP_STUDY_ID, &vGetId);
			Assert::IsTrue(vSetId == vGetId, "Clone failed.");

			// Name
			CComVariant vGetName;
			hr = pStudyOut->GetProperty(VSI_PROP_STUDY_NAME, &vGetName);
			Assert::IsTrue(vSetName == vGetName, "Clone failed.");
		}

		/// <summary>
		/// <step>Create a Study</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_STUDY_ERROR_NONE</results></step>
		/// <step>Load a good study data file</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_STUDY_ERROR_NONE</results></step>
		/// <step>Load a invalid study data file</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_STUDY_ERROR_LOAD_XML</results></step>
		/// <step>Load a version incompatible study data file</step>
		/// <step>Get error code<results>Error code retrieved successfully with VSI_STUDY_ERROR_VERSION_INCOMPATIBLE</results></step>
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

			// Study
			CComPtr<IVsiStudy> pStudy;

			hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
			Assert::IsTrue(NULL != pPersist, "IVsiPersistStudy missing.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			// Error code
			VSI_STUDY_ERROR dwErrorCode(VSI_STUDY_ERROR_LOAD_XML);
			hr = pStudy->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_STUDY_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

			// Test good study
			CString strPath;
			strPath.Format(L"%s\\UnitTestStudy.vxml", pszTestDir);

			hr = pPersist->LoadStudyData(strPath, VSI_DATA_TYPE_SERIES_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadStudyData failed.");

			hr = pStudy->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(VSI_STUDY_ERROR_NONE == dwErrorCode, "GetErrorCode failed.");

			// Test bad study
			strPath.Format(L"%s\\UnitTestStudyInvalid.vxml", pszTestDir);

			hr = pPersist->LoadStudyData(strPath, VSI_DATA_TYPE_SERIES_LIST | VSI_DATA_TYPE_NO_SESSION_LINK, NULL);
			Assert::IsTrue(FAILED(hr), "LoadStudyData failed.");

			hr = pStudy->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(0 < (VSI_STUDY_ERROR_LOAD_XML & dwErrorCode), "GetErrorCode failed.");

			// Test version incompatible study
			strPath.Format(L"%s\\UnitTestStudyVersionIncompatible.vxml", pszTestDir);

			hr = pPersist->LoadStudyData(strPath, VSI_DATA_TYPE_STUDY_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadStudyData failed.");

			hr = pStudy->GetErrorCode(&dwErrorCode);
			Assert::AreEqual(S_OK, hr, "GetErrorCode failed.");
			Assert::IsTrue(0 < (VSI_STUDY_ERROR_VERSION_INCOMPATIBLE & dwErrorCode), "GetErrorCode failed.");

			// Clean up
			hr = pStudyMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Set data path: "TestPath\UnitTestStudy.vxml"</step>
		/// <step>Retrieve data path<results>The data path is "TestPath\UnitTestStudy.vxml"</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetDataPath()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strPath;
			strPath.Format(L"%s\\UnitTestStudy.vxml", pszTestDir);

			hr = pStudy->SetDataPath(strPath);
			Assert::AreEqual(S_OK, hr, "SetDataPath failed.");

			// Path
			CComHeapPtr<OLECHAR> pszGetPath;
			hr = pStudy->GetDataPath(&pszGetPath);
			Assert::AreEqual(S_OK, hr, "GetDataPath failed.");

			Assert::IsTrue(strPath == pszGetPath, "Paths not match.");
		}

		/// <summary>
		/// <step>Create a Study</step>
		/// <step>Set properties relevant for Compare function</step>
		/// <step>Compare using all relevant properties<results>Study objects must match</results></step>
		/// </summary>
		[TestMethod]
		void TestCompare()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
			// Name
			CComVariant vSetName(L"Study 0123456789");
			hr = pStudy->SetProperty(VSI_PROP_STUDY_NAME, &vSetName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			// Owner
			CComVariant vSetOwner(L"JJ");
			hr = pStudy->SetProperty(VSI_PROP_STUDY_OWNER, &vSetOwner);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");
			// Date
			SYSTEMTIME st;
			GetSystemTime(&st);

			COleDateTime date(st);

			CComVariant vSet;
			V_VT(&vSet) = VT_DATE;
			V_DATE(&vSet) = DATE(date);

			hr = pStudy->SetProperty(VSI_PROP_STUDY_CREATED_DATE, &vSet);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Clone
			CComPtr<IVsiStudy> pStudyOut;
			hr = pStudy->Clone(&pStudyOut);
			Assert::AreEqual(S_OK, hr, "Clone failed.");
			
			int iCmp;
			hr = pStudy->Compare(pStudyOut, VSI_PROP_STUDY_NAME, &iCmp);
			Assert::AreEqual(S_OK, hr, "Compare call failed.");
			Assert::AreEqual(iCmp, 0, "Study are different.");

			hr = pStudy->Compare(pStudyOut, VSI_PROP_STUDY_OWNER, &iCmp);
			Assert::AreEqual(S_OK, hr, "Compare call failed.");
			Assert::AreEqual(iCmp, 0, "Series are different.");

			hr = pStudy->Compare(pStudyOut, VSI_PROP_STUDY_CREATED_DATE, &iCmp);
			Assert::AreEqual(S_OK, hr, "Compare call failed.");
			Assert::AreEqual(iCmp, 0, "Series are different.");
		}

		/// <summary>
		/// <step>Set study identifier to "01234567890123456789"</step>
		/// <step>Retrieve study identifier<results>The study identifier is "01234567890123456789"</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetProperties()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			VSI_PROP_STUDY propsBstr[] = {
				VSI_PROP_STUDY_ID,
				VSI_PROP_STUDY_ID_CREATED,
				VSI_PROP_STUDY_NS,
				VSI_PROP_STUDY_NAME,
				VSI_PROP_STUDY_OWNER,
				VSI_PROP_STUDY_NOTES,
				VSI_PROP_STUDY_GRANTING_INSTITUTION,
				VSI_PROP_STUDY_VERSION_REQUIRED,
				VSI_PROP_STUDY_VERSION_CREATED
			};

			for (int i = 0; i < _countof(propsBstr); ++i)
			{
				CString strSet;
				strSet.Format(L"%d%d%d", i, i, i);

				CComVariant vSet(strSet.GetString());
				hr = pStudy->SetProperty(propsBstr[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				CComVariant vGet;
				hr = pStudy->GetProperty(propsBstr[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");
			}

			VSI_PROP_STUDY propsBool[] = {
				VSI_PROP_STUDY_LOCKED,
				VSI_PROP_STUDY_EXPORTED,
			};

			for (int i = 0; i < _countof(propsBool); ++i)
			{
				CComVariant vSet(true);
				hr = pStudy->SetProperty(propsBool[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				CComVariant vGet;
				hr = pStudy->GetProperty(propsBool[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");

				vSet = false;
				hr = pStudy->SetProperty(propsBool[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				vGet.Clear();
				hr = pStudy->GetProperty(propsBool[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");
			}

			VSI_PROP_STUDY propsDate[] = {
				VSI_PROP_STUDY_CREATED_DATE,
			};

			for (int i = 0; i < _countof(propsDate); ++i)
			{
				SYSTEMTIME st;
				GetSystemTime(&st);

				COleDateTime date(st);

				CComVariant vSet;
				V_VT(&vSet) = VT_DATE;
				V_DATE(&vSet) = DATE(date);

				hr = pStudy->SetProperty(propsDate[i], &vSet);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				CComVariant vGet;
				hr = pStudy->GetProperty(propsDate[i], &vGet);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");

				Assert::IsTrue(vGet == vSet, "Property value not match.");
			}
		}

		/// <summary>
		/// <step>Loop 10 times</steps>
		/// <step>  Create a Series and add it to the study</step>
		/// <step>  Retrieve Series count<results>The Series count should equal to loop count</results></step>
		/// </summary>
		[TestMethod]
		void TestAddSeries()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			for (int i = 0; i < 10; ++i)
			{
				CComPtr<IVsiSeries> pSeries;
				hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
				Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

				// ID
				CString strId;
				strId.Format(L"%d", i);
				CComVariant vSetId(strId.GetString());
				hr = pSeries->SetProperty(VSI_PROP_SERIES_ID, &vSetId);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				hr = pStudy->AddSeries(pSeries);
				Assert::AreEqual(S_OK, hr, "AddSeries failed.");

				// Test count again
				LONG lCount(99999);
				hr = pStudy->GetSeriesCount(&lCount, NULL);
				Assert::AreEqual(S_OK, hr, "GetSeriesCount failed.");

				Assert::AreEqual(i + 1, lCount, "Incorrect series count.");
			}
		}

		/// <summary>
		/// <step>Create a Series</step>
		/// <step>Set Series ID to "01234567890123456789"</step>
		/// <step>Set name to "Series 0123456789"</step>
		/// <step>Add Series to Study</step>
		/// <step>Retrieve Series from Study<results>Should have a valid Series</results></step>
		/// <step>Remove series</step>
		/// <step>Retrieve Series from Study<results>Should not have a valid Series</results></step>
		/// </summary>
		[TestMethod]
		void TestRemoveSeries()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiSeries> pSeries;
			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// ID
			CString strId(L"01234567890123456789");
			CComVariant vSetId((LPCWSTR)strId);
			hr = pSeries->SetProperty(VSI_PROP_SERIES_ID, &vSetId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Name
			CString strName(L"Series 0123456789");
			CComVariant vSetName((LPCWSTR)strName);
			hr = pSeries->SetProperty(VSI_PROP_SERIES_NAME, &vSetName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			hr = pStudy->AddSeries(pSeries);
			Assert::AreEqual(S_OK, hr, "AddSeries failed.");

			CComPtr<IVsiSeries> pGetSeries;

			hr = pStudy->GetSeries(strId, &pGetSeries);
			Assert::AreEqual(S_OK, hr, "GetSeries failed.");
			Assert::IsTrue(pGetSeries != NULL, "AddSeries failed.");

			// Remove series
			hr = pStudy->RemoveSeries(strId);
			Assert::AreEqual(S_OK, hr, "RemoveSeries failed.");

			pGetSeries.Release();
			hr = pStudy->GetSeries(strId, &pGetSeries);
			Assert::AreEqual(S_FALSE, hr, "GetSeries failed.");
			Assert::IsTrue(pGetSeries == NULL, "RemoveSeries failed.");
		}

		/// <summary>
		/// <step>Create a study folder which contains 2 series</steps>
		/// <step>LoadStudyData(“real study data with 2 series”)<results>succeed</results></step>
		/// <step>Retrieve Series count<results>The Series count should equal to 2</results></step>
		/// <step>Delete one of the series folder</step>
		/// <step>Retrieve Series count<results>The Series count should equal to 2 (since retrieved from cache)</results></step>
		/// <step>Remove all series (RemoveAllSeries will flush out cache)</step>
		/// <step>Retrieve Series count<results>The Series count should equal to 1 (since retrieved from actual study folder)</results></step>
		/// </summary>
		[TestMethod]
		void TestRemoveAllSeries()
		{
			// Create Study folder: strDir + "\Study"
			String^ strDir = testContextInstance->TestDeploymentDir;
			String^ strDirStudy = gcnew String(strDir);
			strDirStudy += "\\Study";
			Directory::CreateDirectory(strDirStudy);
			Assert::IsTrue(TRUE == Directory::Exists(strDirStudy), "Create Study folder failed.");

			// Create file Study.vxml under the Study folder
			String^ strDirStudyvxml = gcnew String(strDirStudy);
			strDirStudyvxml += "\\Study.vxml";
			FileStream^ fs = File::Create(strDirStudyvxml);
			Assert::IsTrue(File::Exists(strDirStudyvxml), "Create Study.vxml failed.");
			fs->Close();

			// Copy content from UnitTestStudy.vxml to Study.vxml
			String^ strDirStudyvxmlSrc = gcnew String(strDir);
			strDirStudyvxmlSrc += "\\UnitTestStudy.vxml";
			File::Copy(strDirStudyvxmlSrc, strDirStudyvxml, TRUE);

			// Create Series1 folder: strDir + "\Study\Series1"
			String^ strDirSeries1 = gcnew String(strDirStudy);
			strDirSeries1 += "\\Series1";
			Directory::CreateDirectory(strDirSeries1);
			Assert::IsTrue(TRUE == Directory::Exists(strDirSeries1), "Create Series1 folder failed.");

			// Create file Series.vxml under the Series1 folder
			String^ strDirSeries1vxml = gcnew String(strDirSeries1);
			strDirSeries1vxml += "\\Series.vxml";
			fs = File::Create(strDirSeries1vxml);
			Assert::IsTrue(File::Exists(strDirSeries1vxml), "Create Series.vxml failed.");
			fs->Close();

			// Copy content from UnitTestSeries.vxml to Series.vxml
			String^ strDirSeriesvxmlSrc = gcnew String(strDir);
			strDirSeriesvxmlSrc += "\\UnitTestSeries.vxml";
			File::Copy(strDirSeriesvxmlSrc, strDirSeries1vxml, TRUE);

			// Create Series2 folder: strDir + "\Study\Series2"
			String^ strDirSeries2 = gcnew String(strDirStudy);
			strDirSeries2 += "\\Series2";
			Directory::CreateDirectory(strDirSeries2);
			Assert::IsTrue(TRUE == Directory::Exists(strDirSeries2), "Create Series2 folder failed.");

			// Create file Series.vxml under the Series2 folder
			String^ strDirSeries2vxml = gcnew String(strDirSeries2);
			strDirSeries2vxml += "\\Series.vxml";
			fs = File::Create(strDirSeries2vxml);
			Assert::IsTrue(File::Exists(strDirSeries2vxml), "Create Series.vxml failed.");
			fs->Close();

			// Copy content from UnitTestSeries.vxml to Series.vxml
			File::Copy(strDirSeriesvxmlSrc, strDirSeries2vxml, TRUE);

			// Study->LoadStudyData(“real study data with 2 series”)
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
			Assert::IsTrue(NULL != pPersist, "Get IVsiPersistStudy failed.");

			CString strPathStudyvxml(strDirStudyvxml);
			hr = pPersist->LoadStudyData(strPathStudyvxml, VSI_DATA_TYPE_SERIES_LIST | VSI_DATA_TYPE_NO_SESSION_LINK, NULL);
			Assert::AreEqual(S_OK, hr, "LoadStudyData failed.");
			
			// Study->GetSeriesCount() – expects 2
			LONG lCount(99999);
			hr = pStudy->GetSeriesCount(&lCount, NULL);
			Assert::AreEqual(S_OK, hr, "GetSeriesCount failed.");
			Assert::AreEqual(2, lCount, "Incorrect series count.");

			// Delete one of the series folder
			File::Delete(strDirSeries1vxml);
			Directory::Delete(strDirSeries1);

			// Study->GetSeriesCount() – expects 2
			lCount = 0;
			hr = pStudy->GetSeriesCount(&lCount, NULL);
			Assert::AreEqual(S_OK, hr, "GetSeriesCount failed.");
			Assert::AreEqual(2, lCount, "Incorrect series count.");

			// Study->RemoveAllSeries()
			hr = pStudy->RemoveAllSeries();
			Assert::AreEqual(S_OK, hr, "RemoveAllSeries failed.");

			// Study->GetSeriesCount() – expects 1
			hr = pStudy->GetSeriesCount(&lCount, NULL);
			Assert::AreEqual(S_OK, hr, "GetSeriesCount failed.");
			Assert::AreEqual(1, lCount, "Incorrect series count.");

			// Clean Up
			File::Delete(strDirSeries2vxml);
			Directory::Delete(strDirSeries2);
			File::Delete(strDirStudyvxml);
			Directory::Delete(strDirStudy);
		}

		/// <summary>
		/// <step>Loop 100 times</steps>
		/// <step>  Create a Series and add it to the study</step>
		/// <step>  Retrieve Series count<results>The Series count should equal to loop count</results></step>
		/// </summary>
		[TestMethod]
		void TestGetSeriesCount()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			for (int i = 0; i < 100; ++i)
			{
				CComPtr<IVsiSeries> pSeries;
				hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
				Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

				// ID
				CString strId;
				strId.Format(L"%d", i);
				CComVariant vSetId(strId.GetString());
				hr = pSeries->SetProperty(VSI_PROP_SERIES_ID, &vSetId);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				hr = pStudy->AddSeries(pSeries);
				Assert::AreEqual(S_OK, hr, "AddSeries failed.");

				// Test count again
				LONG lCount(99999);
				hr = pStudy->GetSeriesCount(&lCount, NULL);
				Assert::AreEqual(S_OK, hr, "GetSeriesCount failed.");

				Assert::AreEqual(i + 1, lCount, "Incorrect series count.");
			}
		}

		/// <summary>
		/// <step>Create a Series</step>
		/// <step>Set Series ID to "01234567890123456789"</step>
		/// <step>Set name to "Series 0123456789"</step>
		/// <step>Add Series to Study</step>
		/// <step>Retrieve Series from Study<results>Should have a valid Series</results></step>
		/// </summary>
		[TestMethod]
		void TestGetSeries()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiSeries> pSeries;
			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// ID
			CString strId(L"01234567890123456789");
			CComVariant vSetId((LPCWSTR)strId);
			hr = pSeries->SetProperty(VSI_PROP_SERIES_ID, &vSetId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Name
			CString strName(L"Series 0123456789");
			CComVariant vSetName((LPCWSTR)strName);
			hr = pSeries->SetProperty(VSI_PROP_SERIES_NAME, &vSetName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			hr = pStudy->AddSeries(pSeries);
			Assert::AreEqual(S_OK, hr, "AddSeries failed.");

			CComPtr<IVsiSeries> pGetSeries;

			hr = pStudy->GetSeries(strId, &pGetSeries);
			Assert::AreEqual(S_OK, hr, "GetSeries failed.");
			Assert::IsTrue(pGetSeries != NULL, "AddSeries failed.");
		}

		/// <summary>
		/// <step>Create a Series</step>
		/// <step>Set Series ID to "01234567890123456789"</step>
		/// <step>Set name to "Series 0123456789"</step>
		/// <step>Add Series to Study</step>
		/// <step>Retrieve Series from Study using GetItem<results>Should have a valid Series</results></step>
		/// </summary>
		[TestMethod]
		void TestGetItem()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiSeries> pSeries;
			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// ID
			CString strId(L"01234567890123456789");
			CComVariant vSetId((LPCWSTR)strId);
			hr = pSeries->SetProperty(VSI_PROP_SERIES_ID, &vSetId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			// Name
			CString strName(L"Series 0123456789");
			CComVariant vSetName((LPCWSTR)strName);
			hr = pSeries->SetProperty(VSI_PROP_SERIES_NAME, &vSetName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			hr = pStudy->AddSeries(pSeries);
			Assert::AreEqual(S_OK, hr, "AddSeries failed.");

			CComPtr<IUnknown> pUnkSeries;

			hr = pStudy->GetItem(strId, &pUnkSeries);
			Assert::AreEqual(S_OK, hr, "GetSeries failed.");
			Assert::IsTrue(pUnkSeries != NULL, "GetItem failed.");

			CComQIPtr<IVsiSeries> pGetSeries(pUnkSeries);
			Assert::IsTrue(pGetSeries != NULL, "GetItem failed.");
		}

		/// <summary>
		/// <step>Loop 10 times to create 10 Series</step>
		/// <step>	Set Series ID</step>
		/// <step>	Set name</step>
		/// <step>	Add Series to Study</step>
		/// <step>Retrieve Series Enumerator</step>
		/// <step>Retrieve Series from Series Enumerator<results>Should have a valid Series</results></step>
		/// </summary>
		[TestMethod]
		void TestGetSeriesEnumerator()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComPtr<IVsiSeries> pSeries;
			hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CComVariant vSetId;
			CComVariant vSetName;

			for (int i = 0; i < 10; ++i)
			{
				// ID
				CString strId;
				strId.Format(L"%d", i);
				vSetId = strId.GetString();
				hr = pSeries->SetProperty(VSI_PROP_SERIES_ID, &vSetId);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				// Name
				CString strName;
				strName.Format(L"Series%d", i);
				vSetName = (LPCWSTR)strName;
				hr = pSeries->SetProperty(VSI_PROP_SERIES_NAME, &vSetName);
				Assert::AreEqual(S_OK, hr, "SetProperty failed.");

				hr = pStudy->AddSeries(pSeries);
				Assert::AreEqual(S_OK, hr, "AddSeries failed.");
			}

			// Enumerator
			CComPtr<IVsiEnumSeries> pEnum;
			hr = pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_MIN, FALSE, &pEnum, NULL);
			Assert::AreEqual(S_OK, hr, "GetSeriesEnumerator failed.");
			Assert::IsTrue(pEnum != NULL, "GetSeriesEnumerator failed.");

			int iCount(0);
			pSeries.Release();
			while (pEnum->Next(1, &pSeries, NULL) == S_OK)
			{
				++iCount;

				CComVariant vGetId;
				hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vGetId);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");
				Assert::IsTrue(vGetId == vSetId, "GetSeriesEnumerator failed.");

				CComVariant vGetName;
				hr = pSeries->GetProperty(VSI_PROP_SERIES_NAME, &vGetName);
				Assert::AreEqual(S_OK, hr, "GetProperty failed.");
				Assert::IsTrue(vGetName == vSetName, "GetSeriesEnumerator failed.");

				pSeries.Release();
			}

			Assert::AreEqual(10, iCount, "GetSeriesEnumerator failed.");
		}

		//test methods in IVsiPersistStudy
		/// <summary>
		/// <step>Create a Study object</step>
		/// <step>Load a study from a test file</step>
		/// <step>Check all properties against expected values<results>The call must succeed and the properties must match</results></step>
		/// </summary>
		[TestMethod]
		void TestLoadStydyData()
		{
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

			CComVariant vExpected(L"0123456789ABCDEFGHIJKLMON");
			CComVariant vGet;
			hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"0123456789ABCDEFGHIJKLMON");
			hr = pStudy->GetProperty(VSI_PROP_STUDY_ID_CREATED, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"Test1");
			hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			SYSTEMTIME st;
			st.wYear = 2005;
			st.wMonth = 1;
			st.wDay = 22;
			st.wHour = 19;
			st.wMinute = 0;
			st.wSecond = 2;
			st.wMilliseconds = 687;
			COleDateTime date(st);
			V_VT(&vExpected) = VT_DATE;
			V_DATE(&vExpected) = DATE(date);
			hr = pStudy->GetProperty(VSI_PROP_STUDY_CREATED_DATE, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"Sierra Developer");
			hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(false);
			hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");

			vExpected.Clear();
			vGet.Clear();

			vExpected = CComVariant(L"Testing 123 - Hello. VisualSonics Inc.");
			hr = pStudy->GetProperty(VSI_PROP_STUDY_NOTES, &vGet);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");

			Assert::IsTrue(vGet == vExpected, "Property value not match.");
		}

		/// <summary>
		/// <step>Create a Study object</step>
		/// <step>Set some properties to the object</step>
		/// <step>Save a study to a test file</step>
		/// <step>Load a study from a test file</step>
		/// <step>Check all properties against expected values<results>The call must succeed and the properties must match</results></step>
		/// </summary>
		[TestMethod]
		void TestSaveStydyData()
		{
			CComPtr<IVsiStudy> pStudy;

			HRESULT hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
			
			CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
			Assert::IsTrue(NULL != pPersist, "IVsiPersistStudy missing.");

			// Data path
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CComVariant vSetStuId(L"01E67435-7295-44D7-8116-483DC5A08102");
			hr = pStudy->SetProperty(VSI_PROP_STUDY_ID, &vSetStuId);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComVariant vSetStuIdCreated(L"A94E4564-14B5-46AE-8231-B1B1954A77D8");
			hr = pStudy->SetProperty(VSI_PROP_STUDY_ID_CREATED, &vSetStuIdCreated);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComVariant vSetName(L"StudyNew");
			hr = pStudy->SetProperty(VSI_PROP_STUDY_NAME, &vSetName);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			SYSTEMTIME st;
			GetSystemTime(&st);

			COleDateTime date(st);

			CComVariant vSetDate;
			V_VT(&vSetDate) = VT_DATE;
			V_DATE(&vSetDate) = DATE(date);

			hr = pStudy->SetProperty(VSI_PROP_STUDY_CREATED_DATE, &vSetDate);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComVariant vSetOwner(L"Sierra Developer");
			hr = pStudy->SetProperty(VSI_PROP_STUDY_OWNER, &vSetOwner);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComVariant vSetLocked(true);
			hr = pStudy->SetProperty(VSI_PROP_STUDY_LOCKED, &vSetLocked);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CComVariant vSetNotes(L"some notes go here");
			hr = pStudy->SetProperty(VSI_PROP_STUDY_NOTES, &vSetNotes);
			Assert::AreEqual(S_OK, hr, "SetProperty failed.");

			CString strPath;
			strPath.Format(L"%s\\TestSaveStudy.vxml", pszTestDir);
			hr = pPersist->SaveStudyData(strPath, VSI_DATA_TYPE_STUDY_LIST);
			Assert::AreEqual(S_OK, hr, "SaveStudyData failed.");

			CComPtr<IVsiStudy> pStudyLoaded;

			hr = pStudyLoaded.CoCreateInstance(__uuidof(VsiStudy));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
			
			CComQIPtr<IVsiPersistStudy> pPersistLoaded(pStudyLoaded);
			Assert::IsTrue(NULL != pPersistLoaded, "IVsiPersistSeries missing.");

			hr = pPersistLoaded->LoadStudyData(strPath, VSI_DATA_TYPE_STUDY_LIST, NULL);
			Assert::AreEqual(S_OK, hr, "LoadStudyData failed.");

			CComVariant vGetStuId;
			hr = pStudyLoaded->GetProperty(VSI_PROP_STUDY_ID, &vGetStuId);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetStuId == vGetStuId, "Clone failed.");

			CComVariant vGetStuIdCreated;
			hr = pStudyLoaded->GetProperty(VSI_PROP_STUDY_ID_CREATED, &vGetStuIdCreated);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetStuIdCreated == vGetStuIdCreated, "Clone failed.");

			CComVariant vGetName;
			hr = pStudyLoaded->GetProperty(VSI_PROP_STUDY_NAME, &vGetName);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetName == vGetName, "Clone failed.");

			CComVariant vGetDate;
			hr = pStudyLoaded->GetProperty(VSI_PROP_STUDY_CREATED_DATE, &vGetDate);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetDate == vGetDate, "Clone failed.");

			CComVariant vGetOwner;
			hr = pStudyLoaded->GetProperty(VSI_PROP_STUDY_OWNER, &vGetOwner);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetOwner == vGetOwner, "Clone failed.");

			CComVariant vGetLocked;
			hr = pStudyLoaded->GetProperty(VSI_PROP_STUDY_LOCKED, &vGetLocked);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetLocked == vGetLocked, "Clone failed.");

			CComVariant vGetNotes;
			hr = pStudyLoaded->GetProperty(VSI_PROP_STUDY_NOTES, &vGetNotes);
			Assert::AreEqual(S_OK, hr, "GetProperty failed.");
			Assert::IsTrue(vSetNotes == vGetNotes, "Clone failed.");

			// Clean Up
			String^ strDirTest = gcnew String(strDir);
			strDirTest += "\\TestSaveStudy.vxml";
			File::Delete(strDirTest);
		}

		//test methods in IVsiEnumStudy

	private:
		void CreateApp(IVsiApp** ppApp)
		{
			CComPtr<IVsiTestApp> pTestApp;

			HRESULT hr = pTestApp.CoCreateInstance(__uuidof(VsiTestApp));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance app failed.");

			hr = pTestApp->InitServices(VSI_TESTAPP_SERVICE_PDM | VSI_TESTAPP_SERVICE_SESSION_MGR | VSI_TESTAPP_SERVICE_OPERATOR_MGR);
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
