/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiTestOperator.cpp
**
**	Description:
**		VsiOperator Test
**
********************************************************************************/
#include "stdafx.h"
#include <VsiUnitTestHelper.h>
#include <VsiServiceKey.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterModule.h>
#include <VsiPdmModule.h>
#include <VsiOperatorModule.h>
#include <VsiStudyModule.h>

using namespace System;
using namespace System::Text;
using namespace System::Collections::Generic;
using namespace Microsoft::VisualStudio::TestTools::UnitTesting;
using namespace System::IO;

namespace VsiOperatorModuleTest
{
	[TestClass]
	public ref class VsiTestOperatorManager
	{
	private:
		TestContext^ testContextInstance;
		ULONG_PTR m_cookie;

	public: 
		/// <summary>
		/// Gets or sets the test context which provides
		/// information about and functionality for the current test run.
		/// </summary>
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
			LoadLibraryEx(L"VsiResVevo2100.dll", NULL, 0);

			m_cookie = 0;

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			ACTCTX actCtx;
			memset((void*)&actCtx, 0, sizeof(ACTCTX));
			actCtx.cbSize = sizeof(ACTCTX);
			actCtx.dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID | ACTCTX_FLAG_RESOURCE_NAME_VALID;
			actCtx.lpSource = L"VsiOperatorModuleTest.dll";
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

		[TestMethod]
		void TestCreate()
		{
			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
		};

		//Test IVsiOperatorManager

		/// <summary>
		/// <step>Initialize VsiOperatorManager<results>CVsiTestOperatorManager initializes successfully</results></step>
		/// <step>Uninitialize VsiOperatorManager<results>CVsiTestOperatorManager un-initializes successfully</results></step>
		/// </summary>
		[TestMethod]
		void TestInitializeUninitialize()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Clean up
			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>GetOperators<results>IVsiEnumOperator returns successfully & 0 operators</results></step>
		/// </summary>
		[TestMethod]
		void TestGetOperators()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComPtr<IVsiEnumOperator> pOperators;
			hr = pOperatorMgr->GetOperators(
				VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN,
				&pOperators);
			Assert::AreEqual(S_OK, hr, "GetOperators failed.");

			CComPtr<IVsiOperator> pOperator;
			hr = pOperators->Next(1, &pOperator, NULL);
			Assert::AreEqual(S_FALSE, hr, "Next succeeded - unexpected.");

			// Clean up
			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>GetOperator VsiOperatorManager with empty string<results>Failed...expected</results></step>
		/// <step>AddOperator for VsiOperatorManager<results>Success</results></step>
		/// <step>GetOperator VsiOperatorManager with valid string<results>Success</results></step>
		/// </summary>
		[TestMethod]
		void TestGetOperator()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComPtr<IVsiOperator> pOperator;
			WCHAR szIde[] = L"";
			hr = pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_ID, szIde, &pOperator);
			Assert::AreEqual(S_FALSE, hr, "GetOperator Succeeded - unexpected.");
			pOperator.Release();

			WCHAR szId[] = L"1234";

			hr = CreateOperator(
				L"GetOperator",
				L"GO",
				L"",
				VSI_OPERATOR_TYPE_ADMIN,
				VSI_OPERATOR_STATE_ENABLE,
				szId,
				&pOperator);
			Assert::AreEqual(S_OK, hr, "CreateOperator failed.");

			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			hr = pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_ID, szId, &pOperator);
			Assert::AreEqual(S_OK, hr, "GetOperator failed.");

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Invoke GetCurrentOperator()<results>No current operator</results></step>
		/// <step>Create an operator<results>Success</results></step>
		/// <step>Add the operator to Operator Manager<results>Success</results></step>
		/// <step>Set the operator as current operator<results>Success</results></step>
		/// <step>GetCurrentOperator for VsiOperatorManager<results>Success</results></step>
		/// </summary>
		[TestMethod]
		void TestSetGetCurrentOperator()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			WCHAR szId[] = L"1234";

			CComPtr<IVsiOperator> pOperator;
			hr = pOperatorMgr->GetCurrentOperator(&pOperator);
			Assert::AreEqual(S_FALSE, hr, "Current operator incorrectly assigned.");

			hr = CreateOperator(
				L"GetCurrentOperator",
				L"GCO",
				L"",
				VSI_OPERATOR_TYPE_ADMIN,
				VSI_OPERATOR_STATE_ENABLE,
				szId,
				&pOperator);
			Assert::AreEqual(S_OK, hr, "Create operator failed.");

			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");

			hr = pOperatorMgr->SetCurrentOperator(szId);
			Assert::AreEqual(S_OK, hr, "SetCurrentOperator failed.");

			CComPtr<IVsiOperator> pOperatorReturn;
			hr = pOperatorMgr->GetCurrentOperator(&pOperatorReturn);
			Assert::AreEqual(S_OK, hr, "GetCurrentOperator failed.");

			// Compare the two operators
			CompareOperators(pOperator, pOperatorReturn);

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Initialize Operator Manager with UnitTestOperatorManager.xml as its default data file<results>Success</results></step>
		/// <step>Load operators from UnitTestOperatorManager.xml<results>Success</results></step>
		/// <step>Choose one of the operators from UnitTestOperatorManager.xml and set it as current operator<results>Success</results></step>
		/// <step>Get current operator's name<results>Success</results></step>
		/// <step>Get current operator's data file path<results>Success</results></step>
		/// <step>Check current operator's data file path<results>Correct</results></step>
		/// </summary>
		[TestMethod]
		void TestGetCurrentOperatorDataPath()
		{
			//Initialize Operator Manager with UnitTestOperatorManager.xml as its default data file
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
			
			// Default data file path = pszTestDir + "\UnitTestOperatorManager.xml"
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);
					
			WCHAR szConfigFile[MAX_PATH];
			_snwprintf_s(
				szConfigFile,
				_countof(szConfigFile),
				L"%s\\UnitTestOperatorManager.xml",
				pszTestDir);

			hr = pOperatorMgr->Initialize(pApp, szConfigFile);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Load operators from UnitTestOperatorManager.xml
			CComQIPtr<IVsiPersistOperatorManager> pPersistOperatorMgr(pOperatorMgr);
			Assert::IsTrue(NULL != pPersistOperatorMgr, "IVsiPersistOperatorManager interface missing.");

			hr = pPersistOperatorMgr->LoadOperator(szConfigFile);
			Assert::AreEqual(S_OK, hr, "LoadOperator failed.");

			// Choose one of the operators from UnitTestOperatorManager.xml and set it as current operator
			hr = pOperatorMgr->SetCurrentOperator(L"1");
			Assert::AreEqual(S_OK, hr, "SetCurrentOperator failed.");

			// Get current operator's name
			CComPtr<IVsiOperator> pOperator;
			hr = pOperatorMgr->GetCurrentOperator(&pOperator);
			Assert::AreEqual(S_OK, hr, "GetCurrentOperator failed.");

			CComHeapPtr<OLECHAR> pszGetName;
			hr = pOperator->GetName(&pszGetName);
			Assert::AreEqual(S_OK, hr, "GetName failed.");

			// Get current operator's data file path
			CComHeapPtr<OLECHAR> pszPath;
			hr = pOperatorMgr->GetCurrentOperatorDataPath(&pszPath);
			Assert::AreEqual(S_OK, hr, "GetCurrentOperatorDataPath failed.");

			// Check current operator's data file path
			// Current operator data file path should be pszTestDir + "\" + current operator's name		
			_snwprintf_s(
				szConfigFile,
				_countof(szConfigFile),
				L"%s\\%s",
				pszTestDir,
				pszGetName);
			Assert::IsTrue(0 == wcscmp(szConfigFile, pszPath), "Data Path doesn't match.");

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create Operator 1<results>Success</results></step>
		/// <step>GetRelationshipWithCurrentOperator when no current operator<results>Fail as expected&VSI_OPERATOR_RELATIONSHIP_NONE</results></step>
		/// <step>Set operator 1 to be current operator<results>Success</results></step>
		/// <step>GetRelationshipWithCurrentOperator for operator 1<results>Success&VSI_OPERATOR_RELATIONSHIP_SAME</results></step>
		/// <step>Create Operator 2 and set it to be in the same group as operator 1<results>Success</results></step>
		/// <step>GetRelationshipWithCurrentOperator for operator 2<results>Success&VSI_OPERATOR_RELATIONSHIP_SAME_GROUP</results></step>
		/// <step>Create Operator 3 and set it to be in different groups than operator 1 <results>Success</results></step>
		/// <step>GetRelationshipWithCurrentOperator for operator 3<results>Success&VSI_OPERATOR_RELATIONSHIP_NONE</results></step>
		/// </summary>
		[TestMethod]
		void TestGetRelationshipWithCurrentOperator()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Create Operator 1
			CComPtr<IVsiOperator> pOperator;
			hr = CreateOperator(
				L"CurrentOperator",
				L"Password",
				L"Group A",
				VSI_OPERATOR_TYPE_ADMIN,
				VSI_OPERATOR_STATE_ENABLE,
				L"1",
				&pOperator);
			Assert::AreEqual(S_OK, hr, "Create operator failed.");

			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			// GetRelationshipWithCurrentOperator when no current operator
			VSI_OPERATOR_RELATIONSHIP Relationship;
			hr = pOperatorMgr->GetRelationshipWithCurrentOperator(VSI_OPERATOR_PROP_ID, L"1", &Relationship);
			Assert::AreEqual(S_FALSE, hr, "GetRelationshipWithCurrentOperator succeed which is not expected.");
			Assert::IsTrue(Relationship == VSI_OPERATOR_RELATIONSHIP_NONE, "Relationship not correct.");

			// Set operator 1 to be current operator
			hr = pOperatorMgr->SetCurrentOperator(L"1");
			Assert::AreEqual(S_OK, hr, "SetCurrentOperator failed.");

			// GetRelationshipWithCurrentOperator for operator 1
			hr = pOperatorMgr->GetRelationshipWithCurrentOperator(VSI_OPERATOR_PROP_NAME, L"CurrentOperator", &Relationship);
			Assert::AreEqual(S_OK, hr, "GetRelationshipWithCurrentOperator failed.");
			Assert::IsTrue(Relationship == VSI_OPERATOR_RELATIONSHIP_SAME, "Relationship not correct.");

			// Create Operator 2 and set it to be in the same group as operator 1 
			hr = CreateOperator(
				L"SameGoupOperator",
				L"Passwordsg",
				L"Group A",
				VSI_OPERATOR_TYPE_STANDARD,
				VSI_OPERATOR_STATE_DISABLE,
				L"2",
				&pOperator);
			Assert::AreEqual(S_OK, hr, "Create operator failed.");

			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			// GetRelationshipWithCurrentOperator for operator 2
			hr = pOperatorMgr->GetRelationshipWithCurrentOperator(VSI_OPERATOR_PROP_ID, L"2", &Relationship);
			Assert::AreEqual(S_OK, hr, "GetRelationshipWithCurrentOperator failed.");
			Assert::IsTrue(Relationship == VSI_OPERATOR_RELATIONSHIP_SAME_GROUP, "Relationship not correct.");

			// Create Operator 3 and set it to be in different groups than operator 1 
			hr = CreateOperator(
				L"DifferntGroupOperator",
				L"Passworddg",
				L"Group B",
				VSI_OPERATOR_TYPE_SERVICE_MODE,
				VSI_OPERATOR_STATE_DISABLE_SESSION,
				L"3",
				&pOperator);
			Assert::AreEqual(S_OK, hr, "Create operator failed.");

			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			// GetRelationshipWithCurrentOperator for operator 3
			hr = pOperatorMgr->GetRelationshipWithCurrentOperator(VSI_OPERATOR_PROP_NAME, L"DifferntGroupOperator", &Relationship);
			Assert::AreEqual(S_OK, hr, "GetRelationshipWithCurrentOperator failed.");
			Assert::IsTrue(Relationship == VSI_OPERATOR_RELATIONSHIP_NONE, "Relationship not correct.");

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Initialize Operator Manager with UnitTestOperatorManager.xml as its default data file<results>Success</results></step>
		/// <step>Load operators from UnitTestOperatorManager.xml<results>Success</results></step>
		/// <step>Create the data file directory for operator 1 from UnitTestOperatorManager.xml<results>Success</results></step>
		/// <step>Create VsiOperator.config and put it in operator 1's data file directory<results>Success</results></step>
		/// <step>Create operator A<results>Success</results></step>
		/// <step>Add operator A to Operator Manager and clone settings from operator 1<results>Success</results></step>
		/// <step>Check if operator A is added properly<results>Success</results></step>
		/// <step>Check if the data file directory for operator A (strPath + "\" + operator A's name) is created<results>Success</results></step>
		/// <step>Check if operator A's data file directory contains VsiOperator.config<results>Success & contains</results></step>
		/// </summary>
		[TestMethod]
		void TestAddOperator()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			// Default data file path = strPathXML
			String^ strDir = testContextInstance->TestDeploymentDir;
			CString strPath(strDir);
			CString strPathXML;
			strPathXML.Format(L"%s\\UnitTestOperatorManager.xml", strPath);

			hr = pOperatorMgr->Initialize(pApp, strPathXML);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			//Load operators from UnitTestOperatorManager.xml
			CComQIPtr<IVsiPersistOperatorManager> pPersistOperatorMgr(pOperatorMgr);
			Assert::IsTrue(NULL != pPersistOperatorMgr, "IVsiPersistOperatorManager interface missing.");

			hr = pPersistOperatorMgr->LoadOperator(strPathXML);
			Assert::AreEqual(S_OK, hr, "LoadOperator failed.");

			// Create the data file directory for operator 1: strPath + "\" + operator 1's name	
			CString strPath1;
			strPath1.Format(L"%s\\%s", strPath, L"1");
			String^ strDir1 = gcnew String(strPath1);
			Directory::CreateDirectory(strDir1);
			Assert::IsTrue(TRUE == Directory::Exists(strDir1), "Create Operator 1's data file directory failed.");

			// Create VsiOperator.config and put it in operator 1's data file directory
			CString strPath1in;
			strPath1in.Format(L"%s\\VsiOperator.config", strPath1);
			String^ strDir1in = gcnew String(strPath1in);

			// Create VsiOperator.config
			FileStream^ fs = File::Create(strDir1in);
			Assert::IsTrue(TRUE == File::Exists(strDir1in), "Create VsiOperator.config in operator 1's data file directory failed.");
			fs->Close();

			// Create operator A and add it to Operator Manager
			CComPtr<IVsiOperator> pOperator;
			WCHAR szIdA[] = L"6";
			WCHAR szNameA[] = L"Operator A";
			hr = CreateOperator(
				szNameA,
				L"Password A",
				L"Group A",
				VSI_OPERATOR_TYPE_ADMIN,
				VSI_OPERATOR_STATE_ENABLE,
				szIdA,
				&pOperator);
			Assert::AreEqual(S_OK, hr, "CreateOperator failed.");
			
			// Add operator A to Operator Manager and clone settings from operator 1
			hr = pOperatorMgr->AddOperator(pOperator, L"1");
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");

			// Check if operator A is added properly
			CComPtr<IVsiOperator> pOperatorReturn;
			hr = pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_ID, szIdA, &pOperatorReturn);
			Assert::AreEqual(S_OK, hr, "GetOperator failed.");

			// Compare the operators
			CompareOperators(pOperator, pOperatorReturn);

			pOperatorReturn.Release();
			hr = pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, szNameA, &pOperatorReturn);
			Assert::AreEqual(S_OK, hr, "GetOperator failed.");

			// Compare the operators
			CompareOperators(pOperator, pOperatorReturn);

			// Check if the data file directory for operator A (strPath + "\" + operator A's name) is created	
			CString strPathA;
			strPathA.Format(L"%s\\%s", strPath, szNameA);
			String^ strDirA = gcnew String(strPathA);
			Assert::IsTrue(TRUE == Directory::Exists(strDirA), "Create Operator A's data file directory failed.");

			// Check if operator A's data file directory contains VsiOperator.config 
			CString strPathAin;
			strPathAin.Format(L"%s\\VsiOperator.config", strPathA);
			String^ strDirAin = gcnew String(strPathAin);
			Assert::IsTrue(TRUE == File::Exists(strDirAin), "Clone VsiOperator.config to operator A's data file directory failed.");

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);

			// Clean Up
			File::Delete(strDir1in);
			Directory::Delete(strDir1);
			File::Delete(strDirAin);
			Directory::Delete(strDirA);
		}

		/// <summary>
		/// <step>Add 3 operators for VsiOperatorManager</step>
		/// <step>Get number of operators<results>3 operators</results></step>
		/// <step>Remove 3 operators for VsiOperatorManager<results>Successfully</results></step>
		/// <step>Get number of operators<results>0 operators</results></step>
		/// </summary>
		[TestMethod]
		void TestRemoveOperator()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			WCHAR szId1[] = L"1";
			WCHAR szId2[] = L"2";
			WCHAR szId3[] = L"3";

			CComPtr<IVsiOperator> pOperator;
			hr = CreateOperator(L"Operator1", L"O1", L"", VSI_OPERATOR_TYPE_ADMIN, VSI_OPERATOR_STATE_ENABLE, szId1, &pOperator);
			Assert::AreEqual(S_OK, hr, "CreateOperator failed.");
			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			hr = CreateOperator(L"Operator2", L"10", L"", VSI_OPERATOR_TYPE_ADMIN, VSI_OPERATOR_STATE_ENABLE, szId2, &pOperator);
			Assert::AreEqual(S_OK, hr, "CreateOperator failed.");
			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			hr = CreateOperator(L"Operator3", L"11", L"", VSI_OPERATOR_TYPE_ADMIN, VSI_OPERATOR_STATE_ENABLE, szId3, &pOperator);
			Assert::AreEqual(S_OK, hr, "CreateOperator failed.");
			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			// Get the list of operators from the VsiOperatorManager
			int iIndex = -1;
			std::list<CString> listCstrId;
			hr = GetNumberOperators(pOperatorMgr, &iIndex, &listCstrId);
			Assert::AreEqual(S_OK, hr, "GetNumberOperators failed.");
			Assert::AreEqual(3, iIndex, "Incorrect number of operators.");

			// Perform removal of operators
			std::list<CString>::iterator iterListId = listCstrId.begin();
			for ( ; iterListId != listCstrId.end(); ++iterListId)
			{
				const CString &csId = *iterListId;
				hr = pOperatorMgr->RemoveOperator((LPCWSTR)csId);
				Assert::AreEqual(S_OK, hr, "RemoveOperator failed.");
			}

			hr = GetNumberOperators(pOperatorMgr, &iIndex, NULL);
			Assert::AreEqual(S_OK, hr, "GetNumberOperators failed.");
			Assert::AreEqual(0, iIndex, "Incorrect number of operators.");

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create the operator manager<results>Success</results></step>
		/// <step>Invoke IsDirty()<results>Not dirty</results></step>
		/// <step>Invoke OperatorModified on operator manager</step>
		/// <step>Invoke IsDirty()<results>Is dirty</results></step>
		/// </summary>
		[TestMethod]
		void TestOperatorModified()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComQIPtr<IVsiPersistOperatorManager> pPersistOperatorMgr(pOperatorMgr);
			Assert::IsTrue(NULL != pPersistOperatorMgr, "IVsiPersistOperatorManager interface missing.");

			// Initially not dirty
			hr = pPersistOperatorMgr->IsDirty();
			Assert::AreEqual(S_FALSE, hr, "Incorrectly set to dirty.");

			// Set modified
			hr = pOperatorMgr->OperatorModified();
			Assert::AreEqual(S_OK, hr, "OperatorModified failed.");

			hr = pPersistOperatorMgr->IsDirty();
			Assert::AreEqual(S_OK, hr, "Incorrectly set to not dirty.");

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Get count standard|admin<results>0</results></step>
		/// <step>Add an admin operator without password</step>
		/// <step>Get count standard|admin<results>1</results></step>
		/// <step>Add an standard operator without password</step>
		/// <step>Get count standard|admin<results>2</results></step>
		/// <step>Get count standard|admin|password<results>0</results></step>
		/// <step>Add an standard operator with password</step>
		/// <step>Get count standard|admin|password<results>1</results></step>
		/// </summary>
		[TestMethod]
		void TestGetOperatorCount()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			DWORD dwCount(1);

			// Check count
			hr = pOperatorMgr->GetOperatorCount(
				VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN,
				&dwCount);
			Assert::AreEqual(S_OK, hr, "GetOperatorCount failed.");
			Assert::AreEqual(0, (int)dwCount, "Incorrect count.");

			// Add admin
			CComPtr<IVsiOperator> pOperator;
			hr = CreateOperator(
				L"Operator1",
				L"",
				L"",
				VSI_OPERATOR_TYPE_ADMIN,
				VSI_OPERATOR_STATE_ENABLE,
				L"1",
				&pOperator);
			Assert::AreEqual(S_OK, hr, "CreateOperator failed.");
			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			// Check count
			hr = pOperatorMgr->GetOperatorCount(
				VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN,
				&dwCount);
			Assert::AreEqual(S_OK, hr, "GetOperatorCount failed.");
			Assert::AreEqual(1, (int)dwCount, "Incorrect count.");

			// Add standard
			hr = CreateOperator(
				L"Operator2",
				L"",
				L"",
				VSI_OPERATOR_TYPE_STANDARD,
				VSI_OPERATOR_STATE_ENABLE,
				L"2",
				&pOperator);
			Assert::AreEqual(S_OK, hr, "CreateOperator failed.");
			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			// Check count
			hr = pOperatorMgr->GetOperatorCount(
				VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN,
				&dwCount);
			Assert::AreEqual(S_OK, hr, "GetOperatorCount failed.");
			Assert::AreEqual(2, (int)dwCount, "Incorrect count.");

			// Check count
			hr = pOperatorMgr->GetOperatorCount(
				VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_PASSWORD,
				&dwCount);
			Assert::AreEqual(S_OK, hr, "GetOperatorCount failed.");
			Assert::AreEqual(0, (int)dwCount, "Incorrect count.");

			// Add password
			hr = CreateOperator(
				L"Operator3",
				L"O3",
				L"",
				VSI_OPERATOR_TYPE_STANDARD,
				VSI_OPERATOR_STATE_ENABLE,
				L"3",
				&pOperator);
			Assert::AreEqual(S_OK, hr, "CreateOperator failed.");
			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			// Check count
			hr = pOperatorMgr->GetOperatorCount(
				VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_PASSWORD,
				&dwCount);
			Assert::AreEqual(S_OK, hr, "GetOperatorCount failed.");
			Assert::AreEqual(1, (int)dwCount, "Incorrect count.");

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Create an admin operator<results>Success</results></step>
		/// <step>Set the admin operator to be NOT authenticated<results>Success</results></step>
		/// <step>GetIsAuthenticated for the admin operator<results>Success and FALSE</results></step>
		/// <step>GetIsAdminAuthenticated<results>Success and FALSE</results></step>
		/// <step>Set the admin operator to be authenticated<results>Success</results></step>
		/// <step>GetIsAuthenticated for the admin operator<results>Success and TRUE</results></step>
		/// <step>GetIsAdminAuthenticated<results>Success and TRUE</results></step>

		/// <step>Create a standard operator with password<results>Success</results></step>
		/// <step>Set the standard operator to be NOT authenticated<results>Success</results></step>
		/// <step>GetIsAuthenticated for the standard operator<results>Success and TRUE</results></step>
		/// <step>GetIsAdminAuthenticated<results>Success and FALSE</results></step>
		/// <step>Set the standard operator to be authenticated<results>Success</results></step>
		/// <step>GetIsAuthenticated for the standard operator<results>Success and TRUE</results></step>

		/// <step>Create a standard operator without password<results>Success</results></step>
		/// <step>Set the standard operator to be authenticated<results>Success</results></step>
		/// <step>GetIsAuthenticated for the standard operator<results>Success and TRUE</results></step>
		/// <step>Set the standard operator to be NOT authenticated<results>Success</results></step>
		/// <step>GetIsAuthenticated for the standard operator<results>Success and FALSE</results></step>

		/// <step>Set a operator with invalid ID to be authenticated<results>Fail as expected</results></step>
		/// <step>GetIsAuthenticated for the operator with invalid ID<results>Fail as expected and FALSE</results></step>

		/// <step>ClearAuthenticatedList</step>
		/// <step>GetIsAdminAuthenticated<results>Success and FALSE</results></step>
		/// <step>GetIsAuthenticated for the admin operator<results>Success and FALSE</results></step>		
		/// <step>GetIsAuthenticated for the standard operator with password<results>Success and FALSE</results></step>
		/// <step>GetIsAuthenticated for the standard operator without password<results>Success and FALSE</results></step>
		/// </summary>
		[TestMethod]
		void TestIsAuthenticated()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			//Create an admin operator
			WCHAR szIda[] = L"1234";
			CComPtr<IVsiOperator> pOperator;
			hr = CreateOperator(
				L"AdminOperator",
				L"admin",
				L"",
				VSI_OPERATOR_TYPE_ADMIN,
				VSI_OPERATOR_STATE_ENABLE,
				szIda,
				&pOperator);
			Assert::AreEqual(S_OK, hr, "Create operator failed.");

			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			// Set the admin operator to be NOT authenticated
			hr = pOperatorMgr->SetIsAuthenticated(szIda, FALSE);
			Assert::AreEqual(S_OK, hr, "SetIsAuthenticated failed.");

			// GetIsAuthenticated for the admin operator
			BOOL bAuthenticated;
			hr = pOperatorMgr->GetIsAuthenticated(szIda, &bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAuthenticated failed.");
			Assert::IsTrue(FALSE == bAuthenticated, "Authentication status not match.");

			// GetIsAdminAuthenticated
			hr = pOperatorMgr->GetIsAdminAuthenticated(&bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAdminAuthenticated failed.");
			Assert::IsTrue(FALSE == bAuthenticated, "Authentication status not match.");

			// Set the admin operator to be authenticated
			hr = pOperatorMgr->SetIsAuthenticated(szIda, TRUE);
			Assert::AreEqual(S_OK, hr, "SetIsAuthenticated failed.");

			// GetIsAuthenticated for the admin operator
			hr = pOperatorMgr->GetIsAuthenticated(szIda, &bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAuthenticated failed.");
			Assert::IsTrue(TRUE == bAuthenticated, "Authentication status not match.");

			// GetIsAdminAuthenticated
			hr = pOperatorMgr->GetIsAdminAuthenticated(&bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAdminAuthenticated failed.");
			Assert::IsTrue(TRUE == bAuthenticated, "Authentication status not match.");

			// Create a standard operator with password
			WCHAR szIdsp[] = L"5678";
			hr = CreateOperator(
				L"StandardOperator",
				L"standard",
				L"",
				VSI_OPERATOR_TYPE_STANDARD,
				VSI_OPERATOR_STATE_ENABLE,
				szIdsp,
				&pOperator);
			Assert::AreEqual(S_OK, hr, "Create operator failed.");

			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			// Set the standard operator to be NOT authenticated
			hr = pOperatorMgr->SetIsAuthenticated(szIdsp, FALSE); //not authenticated
			Assert::AreEqual(S_OK, hr, "SetIsAuthenticated failed.");

			// GetIsAuthenticated for the standard operator
			hr = pOperatorMgr->GetIsAuthenticated(szIdsp, &bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAuthenticated failed.");
			Assert::IsTrue(FALSE == bAuthenticated, "Authentication status not match.");

			// GetIsAdminAuthenticated
			hr = pOperatorMgr->GetIsAdminAuthenticated(&bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAdminAuthenticated failed.");
			// The authenticated Admin operator is still there
			Assert::IsTrue(TRUE == bAuthenticated, "Authentication status not match.");

			// Set the standard operator to be authenticated
			hr = pOperatorMgr->SetIsAuthenticated(szIdsp, TRUE);
			Assert::AreEqual(S_OK, hr, "SetIsAuthenticated failed.");

			// GetIsAuthenticated for the standard operator
			hr = pOperatorMgr->GetIsAuthenticated(szIdsp, &bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAuthenticated failed.");
			Assert::IsTrue(TRUE == bAuthenticated, "Authentication status not match.");

			//Create a standard operator without password
			WCHAR szIds[] = L"90";
			hr = CreateOperator(
				L"StandardOperator",
				L"",
				L"",
				VSI_OPERATOR_TYPE_STANDARD,
				VSI_OPERATOR_STATE_ENABLE,
				szIds,
				&pOperator);
			Assert::AreEqual(S_OK, hr, "Create operator failed.");

			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			// Set the standard operator to be authenticated
			hr = pOperatorMgr->SetIsAuthenticated(szIds, TRUE);
			Assert::AreEqual(S_OK, hr, "SetIsAuthenticated failed.");

			// GetIsAuthenticated for the standard operator
			hr = pOperatorMgr->GetIsAuthenticated(szIds, &bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAuthenticated failed.");
			Assert::IsTrue(TRUE == bAuthenticated, "Authentication status not match.");

			// Set the standard operator to be NOT authenticated
			hr = pOperatorMgr->SetIsAuthenticated(szIds, FALSE);
			Assert::AreEqual(S_OK, hr, "SetIsAuthenticated failed.");

			// GetIsAuthenticated for the standard operator
			hr = pOperatorMgr->GetIsAuthenticated(szIds, &bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAuthenticated failed.");
			Assert::IsTrue(FALSE == bAuthenticated, "Authentication status not match.");
			
			// Set a operator with invalid ID to be authenticated
			hr = pOperatorMgr->SetIsAuthenticated(L"", TRUE);
			Assert::AreEqual(S_FALSE, hr, "SetIsAuthenticated succeed which is not expected.");

			// GetIsAuthenticated for the operator with invalid ID
			hr = pOperatorMgr->GetIsAuthenticated(L"", &bAuthenticated);
			Assert::AreEqual(S_FALSE, hr, "GetIsAuthenticated succeed which is not expected.");
			Assert::IsTrue(FALSE == bAuthenticated, "Authentication status not match.");

			// ClearAuthenticatedList
			hr = pOperatorMgr->ClearAuthenticatedList();

			// GetIsAdminAuthenticated
			hr = pOperatorMgr->GetIsAdminAuthenticated(&bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAdminAuthenticated failed.");
			Assert::IsTrue(FALSE == bAuthenticated, "Authentication status not match.");
						
			// GetIsAuthenticated for the admin operator
			hr = pOperatorMgr->GetIsAuthenticated(szIda, &bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAuthenticated failed.");
			Assert::IsTrue(FALSE == bAuthenticated, "Authentication status not match.");
						
			// GetIsAuthenticated for the standard operator with password
			hr = pOperatorMgr->GetIsAuthenticated(szIdsp, &bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAuthenticated failed.");
			Assert::IsTrue(FALSE == bAuthenticated, "Authentication status not match.");

			// GetIsAuthenticated for the standard operator without password
			hr = pOperatorMgr->GetIsAuthenticated(szIds, &bAuthenticated);
			Assert::AreEqual(S_OK, hr, "GetIsAuthenticated failed.");
			Assert::IsTrue(FALSE == bAuthenticated, "Authentication status not match.");

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		//TestIVsiPersistOperatorManager

		/// <summary>
		/// <step>Create the operator manager<results>Success</results></step>
		/// <step>Invoke IsDirty()<results>Not dirty</results></step>
		/// <step>Invoke OperatorModified on operator manager</step>
		/// <step>Invoke IsDirty()<results>Is dirty</results></step>
		/// </summary>
		[TestMethod]
		void TestPersistIsDirty()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComQIPtr<IVsiPersistOperatorManager> pPersistOperatorMgr(pOperatorMgr);
			Assert::IsTrue(NULL != pPersistOperatorMgr, "IVsiPersistOperatorManager interface missing.");

			// Initially not dirty
			hr = pPersistOperatorMgr->IsDirty();
			Assert::AreEqual(S_FALSE, hr, "Incorrectly set to dirty.");

			// Set modified
			hr = pOperatorMgr->OperatorModified();
			Assert::AreEqual(S_OK, hr, "OperatorModified() failed.");

			hr = pPersistOperatorMgr->IsDirty();
			Assert::AreEqual(S_OK, hr, "Incorrectly set to not dirty.");

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>LoadOperator with non-existing filename<results>Failed ... expected</results></step>
		/// <step>LoadOperator from UnitTestOperatorManager.xml for VsiOperatorManager<results>Success</results></step>
		/// <step>Check for number of administrator operators<results>3</results></step>
		/// <step>Check for number of standard operators<results>2</results></step>
		/// </summary>
		[TestMethod]
		void TestLoadOperator()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComQIPtr<IVsiPersistOperatorManager> pPersistOperatorMgr(pOperatorMgr);
			Assert::IsTrue(NULL != pPersistOperatorMgr, "IVsiPersistOperatorManager interface missing.");

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			// Non-existing file name
			{
				WCHAR szConfigFile[MAX_PATH];
				_snwprintf_s(
					szConfigFile,
					_countof(szConfigFile),
					L"%s\\TestOperatorManagerNonExistingFile.xml",
					pszTestDir);

				hr = pPersistOperatorMgr->LoadOperator(szConfigFile);
				Assert::AreEqual(S_FALSE, hr, "Incorrect return code from LoadOperator.");
			}

			// Load operator using UnitTestOperatorManager.xml file
			{
				WCHAR szConfigFile[MAX_PATH];
				_snwprintf_s(
					szConfigFile,
					_countof(szConfigFile),
					L"%s\\UnitTestOperatorManager.xml",
					pszTestDir);

				hr = pPersistOperatorMgr->LoadOperator(szConfigFile);
				Assert::AreEqual(S_OK, hr, "LoadOperator failed.");

				DWORD dwCount(0);

				hr = pOperatorMgr->GetOperatorCount(VSI_OPERATOR_FILTER_TYPE_ADMIN, &dwCount);
				Assert::AreEqual(S_OK, hr, "GetOperatorCount failed.");
				Assert::AreEqual(3, (int)dwCount, "Incorrect operator count.");

				hr = pOperatorMgr->GetOperatorCount(VSI_OPERATOR_FILTER_TYPE_STANDARD, &dwCount);
				Assert::AreEqual(S_OK, hr, "GetOperatorCount failed.");
				Assert::AreEqual(2, (int)dwCount, "Incorrect operator count.");
			}

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>LoadOperator from UnitTestOperatorManager.xml<results>Success</results></step>
		/// <step>SaveOperator to UnitTestOperatorManager1.xml<results>Success</results></step>
		/// <step>LoadOperator from UnitTestOperatorManager1.xml<results>Success</results></step>
		/// <step>Check for number of administrator operators<results>3</results></step>
		/// <step>Check for number of standard operators<results>2</results></step>
		/// </summary>
		[TestMethod]
		void TestSaveOperator()
		{
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			CComQIPtr<IVsiPersistOperatorManager> pPersistOperatorMgr(pOperatorMgr);
			Assert::IsTrue(NULL != pPersistOperatorMgr, "IVsiPersistOperatorManager interface missing.");

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			// Load operator using UnitTestOperatorManager.xml file
			{
				WCHAR szConfigFile[MAX_PATH];
				_snwprintf_s(
					szConfigFile,
					_countof(szConfigFile),
					L"%s\\UnitTestOperatorManager.xml",
					pszTestDir);

				hr = pPersistOperatorMgr->LoadOperator(szConfigFile);
				Assert::AreEqual(S_OK, hr, "LoadOperator failed.");
			}

			// Save operator using UnitTestOperatorManager1.xml file
			{
				WCHAR szConfigFile[MAX_PATH];
				_snwprintf_s(
					szConfigFile,
					_countof(szConfigFile),
					L"%s\\UnitTestOperatorManager1.xml",
					pszTestDir);

				hr = pPersistOperatorMgr->SaveOperator(szConfigFile);
				Assert::AreEqual(S_OK, hr, "SaveOperator failed.");
			}

			// Load operator using UnitTestOperatorManager1.xml file
			{
				WCHAR szConfigFile[MAX_PATH];
				_snwprintf_s(
					szConfigFile,
					_countof(szConfigFile),
					L"%s\\UnitTestOperatorManager1.xml",
					pszTestDir);

				hr = pPersistOperatorMgr->LoadOperator(szConfigFile);
				Assert::AreEqual(S_OK, hr, "LoadOperator failed.");

				DWORD dwCount(0);

				hr = pOperatorMgr->GetOperatorCount(VSI_OPERATOR_FILTER_TYPE_ADMIN, &dwCount);
				Assert::AreEqual(S_OK, hr, "GetOperatorCount failed.");
				Assert::AreEqual(3, (int)dwCount, "Incorrect operator count.");

				hr = pOperatorMgr->GetOperatorCount(VSI_OPERATOR_FILTER_TYPE_STANDARD, &dwCount);
				Assert::AreEqual(S_OK, hr, "GetOperatorCount failed.");
				Assert::AreEqual(2, (int)dwCount, "Incorrect operator count.");
			}

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		// TestIVsiEnumOperator

		/// <summary>
		/// <step>Get operators from UnitTestOperatorManager.xml to pOperators list<results>success & 5 operators</results></step>
		/// <step>Fetch 3 operators in the pOperators list<results>success & fetched 3 operators & operator 1&2&3</results></step>
		/// <step>Attempt to fetch another 3 operators in the pOperators list<results>fail & fetched 2 operators & operator 4&5</results></step>
		/// <step>Clone the pOperators list<results>success and 0 operators</results></step>
		/// <step>Attempt to fetch a operator in the pOperatorsClone list<results>fail & no operator fetched</results></step>
		/// <step>Reset the pOperators list<results>success & 5 operators</results></step>
		/// <step>Clone the pOperators list<results>success & 5 operators</results></step>
		/// <step>Fetch a operator in the pOperators list<results>success & fetched 1 operator & operator 1</results></step>
		/// <step>Skip 2 operators in the pOperators list<results>success</results></step>
		/// <step>Fetch a operator in the pOperators list<results>success & fetched 1 operator & operator 4</results></step>
		/// </summary>
		[TestMethod]
		void TestEnumOperator()
		{
			// Initialize Operator Manager with UnitTestOperatorManager.xml as its default data file
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			HRESULT hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
			
			// Default data file path = strPathXML
			String^ strDir = testContextInstance->TestDeploymentDir;
			CString strPath(strDir);
			CString strPathXML;
			strPathXML.Format(L"%s\\UnitTestOperatorManager.xml", strPath);

			hr = pOperatorMgr->Initialize(pApp, strPathXML);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			// Load operators from UnitTestOperatorManager.xml
			CComQIPtr<IVsiPersistOperatorManager> pPersistOperatorMgr(pOperatorMgr);
			Assert::IsTrue(NULL != pPersistOperatorMgr, "IVsiPersistOperatorManager interface missing.");

			hr = pPersistOperatorMgr->LoadOperator(strPathXML);
			Assert::AreEqual(S_OK, hr, "LoadOperator failed.");

			// Get operators from UnitTestOperatorManager.xml to pOperators list
			CComPtr<IVsiEnumOperator> pOperators;
			hr = pOperatorMgr->GetOperators(
				VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN,
				&pOperators);
			Assert::AreEqual(S_OK, hr, "GetOperators failed.");

			// Test Next, Skip, Reset, Clone
			// Array of pointers is double pointer, and can not use smart pointer CComPtr here.
			IVsiOperator * ppOperator[5];
			ZeroMemory(ppOperator, sizeof(ppOperator)); //initialize the array of pointers to NULL
			ULONG celtFetched(0);

			// Fetch 3 operators in the pOperators list
			hr = pOperators->Next(3, ppOperator, &celtFetched);
			Assert::AreEqual(S_OK, hr, "Next failed.");
			Assert::IsTrue(3 == celtFetched, "Fetched wrong number of operators.");
			
			// Return the 3 fetched operators (1&2&3)
			for (ULONG i = 0; i < celtFetched; ++i)
			{
				CComHeapPtr<OLECHAR> pszGetId;			
				hr = ppOperator[i]->GetId(&pszGetId);
				Assert::AreEqual(S_OK, hr, "GetId failed.");

				CString strFetchedId;
				strFetchedId.Format(L"%d", i+1);
				CString strGetId(pszGetId);
				Assert::IsTrue(strFetchedId == strGetId, "Fetched wrong operator.");

				// Not smart pointer, should be released manually
				ppOperator[i]->Release();
			}

			// Attempt to fetch another 3 operators in the pOperators list
			// Since UnitTestOperatorManager.xml contains only 5 operators, i.e, only 2 left
			// Return S_FALSE and actually fetched 2 operators
			hr = pOperators->Next(3, ppOperator, &celtFetched);
			Assert::AreEqual(S_FALSE, hr, "Next succeed - unexpected."); //if (nRem < celt) hr = S_FALSE;
			Assert::IsTrue(2 == celtFetched, "Fetched wrong number of operators."); 

			// Return the 2 fetched operators (4&5)
			for (ULONG i = 0; i < celtFetched; ++i)
			{
				CComHeapPtr<OLECHAR> pszGetId;			
				hr = ppOperator[i]->GetId(&pszGetId);
				Assert::AreEqual(S_OK, hr, "GetId failed.");

				CString strFetchedId;
				strFetchedId.Format(L"%d", i+4);
				CString strGetId(pszGetId);
				Assert::IsTrue(strFetchedId == strGetId, "Fetched wrong operator.");

				// Not smart pointer, should be released manually
				ppOperator[i]->Release();
			}

			// Clone the remaining of pOperators list (empty)
			CComPtr<IVsiEnumOperator> pOperatorsClone;
			hr = pOperators->Clone(&pOperatorsClone);
			Assert::AreEqual(S_OK, hr, "Clone operators failed.");

			// Attempt to fetch a operator in the pOperatorsClone list
			// Since pOperatorsColone is empty
			// Return S_FALSE and actually fetched 0 operator
			hr = pOperatorsClone->Next(1, ppOperator, &celtFetched);
			Assert::AreEqual(S_FALSE, hr, "Next failed.");
			Assert::IsTrue(0 == celtFetched, "Fetched wrong number of operators.");

			// Reset the pOperators list (1&2&3&4&5)
			hr = pOperators->Reset();
			Assert::AreEqual(S_OK, hr, "Reset failed.");
			
			// Clone the pOperators list (1&2&3&4&5)
			pOperatorsClone.Release();
			hr = pOperators->Clone(&pOperatorsClone);
			Assert::AreEqual(S_OK, hr, "Clone operators failed.");

			// Fetch a operator in the pOperators list
			hr = pOperatorsClone->Next(1, ppOperator, &celtFetched);
			Assert::AreEqual(S_OK, hr, "Next failed.");
			Assert::IsTrue(1 == celtFetched, "Fetched wrong number of operators.");

			// Return the fetched operator (1)
			for (ULONG i = 0; i < celtFetched; ++i)
			{
				CComHeapPtr<OLECHAR> pszGetId;			
				hr = ppOperator[i]->GetId(&pszGetId);
				Assert::AreEqual(S_OK, hr, "GetId failed.");

				CString strFetchedId;
				strFetchedId.Format(L"%d", i+1);
				CString strGetId(pszGetId);
				Assert::IsTrue(strFetchedId == strGetId, "Fetched wrong operator.");

				// Not smart pointer, should be released manually
				ppOperator[i]->Release();
			}

			// Skip 2 operators in the pOperators list
			hr = pOperatorsClone->Skip(2);
			Assert::AreEqual(S_OK, hr, "Skip failed.");

			// Fetch a operator in the pOperators list
			hr = pOperatorsClone->Next(1, ppOperator, &celtFetched);
			Assert::AreEqual(S_OK, hr, "Next failed.");
			Assert::IsTrue(1 == celtFetched, "Fetched wrong number of operators.");

			// Return the fetched operator (1)
			for (ULONG i = 0; i < celtFetched; ++i)
			{
				CComHeapPtr<OLECHAR> pszGetId;			
				hr = ppOperator[i]->GetId(&pszGetId);
				Assert::AreEqual(S_OK, hr, "GetId failed.");

				CString strFetchedId;
				strFetchedId.Format(L"%d", i+4);
				CString strGetId(pszGetId);
				Assert::IsTrue(strFetchedId == strGetId, "Fetched wrong operator.");

				// Not smart pointer, should be released manually
				ppOperator[i]->Release();
			}

			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

	private:
		void CreateApp(IVsiApp** ppApp)
		{
			CComPtr<IVsiTestApp> pTestApp;

			HRESULT hr = pTestApp.CoCreateInstance(__uuidof(VsiTestApp));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance app failed.");

			hr = pTestApp->InitServices(VSI_TESTAPP_SERVICE_PDM);
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
			
			hr = pPdm->AddParameter(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_OPERATOR,
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

		HRESULT CreateOperator(
			LPCWSTR pszName, LPCWSTR pszPassword, LPCWSTR pszGroup, VSI_OPERATOR_TYPE dwType,
			VSI_OPERATOR_STATE dwState, WCHAR* pszId, IVsiOperator **ppOperator)
		{
			HRESULT hr;

			CComPtr<IVsiOperator> pOperator;
			hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperator->SetName(pszName);
			Assert::AreEqual(S_OK, hr, "SetName failed.");

			hr = pOperator->SetPassword(pszPassword, TRUE);
			Assert::AreEqual(S_OK, hr, "SetPassword failed.");

			hr = pOperator->SetGroup(pszGroup);
			Assert::AreEqual(S_OK, hr, "SetGroup failed.");

			hr = pOperator->SetType(dwType);
			Assert::AreEqual(S_OK, hr, "SetType failed.");

			hr = pOperator->SetState(dwState);
			Assert::AreEqual(S_OK, hr, "SetState failed.");

			hr = pOperator->SetId(pszId);
			Assert::AreEqual(S_OK, hr, "SetId failed.");

			*ppOperator = pOperator.Detach();

			return hr;
		}

		// Compare two operators for equality using name, password, type, state and pzid
		void CompareOperators(IVsiOperator *pLeftOperator, IVsiOperator *pRightOperator)
		{
			HRESULT hr;

			// Compare names
			CComHeapPtr<OLECHAR> pszName;
			hr = pLeftOperator->GetName(&pszName);
			Assert::AreEqual(S_OK, hr, "GetName failed.");
			CComHeapPtr<OLECHAR> pszNameReturn;
			hr = pRightOperator->GetName(&pszNameReturn);
			Assert::AreEqual(S_OK, hr, "GetName failed.");
			Assert::IsTrue(0 == lstrcmp(pszName, pszNameReturn), "Name not match.");

			// Compare password
			CComHeapPtr<OLECHAR> pszPassword;
			hr = pLeftOperator->GetPassword(&pszPassword);
			Assert::AreEqual(S_OK, hr, "GetPassword failed.");
			CComHeapPtr<OLECHAR> pszPasswordReturn;
			hr = pRightOperator->GetPassword(&pszPasswordReturn);
			Assert::AreEqual(S_OK, hr, "GetPassword failed.");
			Assert::IsTrue(0 == lstrcmp(pszPassword, pszPasswordReturn), "Password not match.");

			// Compare type
			VSI_OPERATOR_TYPE dwType(VSI_OPERATOR_TYPE_NONE);
			hr = pLeftOperator->GetType(&dwType);
			Assert::AreEqual(S_OK, hr, "GetType failed.");
			VSI_OPERATOR_TYPE dwTypeReturn(VSI_OPERATOR_TYPE_NONE);
			hr = pRightOperator->GetType(&dwTypeReturn);
			Assert::AreEqual(S_OK, hr, "GetType failed.");
			Assert::IsTrue(dwType == dwTypeReturn, "Type not match.");

			// Compare state
			VSI_OPERATOR_STATE dwState(VSI_OPERATOR_STATE_DISABLE);
			hr = pLeftOperator->GetState(&dwState);
			Assert::AreEqual(S_OK, hr, "GetState failed.");
			VSI_OPERATOR_STATE dwStateReturn(VSI_OPERATOR_STATE_DISABLE);
			hr = pRightOperator->GetState(&dwStateReturn);
			Assert::AreEqual(S_OK, hr, "GetState failed.");
			Assert::IsTrue(dwState == dwStateReturn, "Type not match.");

			// Compare Ids
			CComHeapPtr<OLECHAR> pszId;
			hr = pLeftOperator->GetId(&pszId);
			Assert::AreEqual(S_OK, hr, "GetId failed.");
			CComHeapPtr<OLECHAR> pszIdReturn;
			hr = pRightOperator->GetId(&pszIdReturn);
			Assert::AreEqual(S_OK, hr, "GetId failed.");
			Assert::IsTrue(0 == lstrcmp(pszId, pszIdReturn), "ID not match.");
		}

		// Traverse the operator manager for the number of operator it contains and store the id of
		// each operator in a list.
		HRESULT GetNumberOperators(
			IVsiOperatorManager *pOperatorMgr,
			int *piCount,
			std::list<CString> *plistCstr)
		{
			// Need hrLoop since the checking of the while loop should be separated from
			// the hr checking of other functionalities (while loop will exit when it is not S_OK)
			// return this value as hr is confusing.
			HRESULT hr;
			CComPtr<IVsiOperator> pOperator;
			CComPtr<IVsiEnumOperator> pOperators;
			hr = pOperatorMgr->GetOperators(
				VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN,
				&pOperators);
			Assert::AreEqual(S_OK, hr, "GetOperators failed.");

			int iCount = 0;
			while (S_OK == pOperators->Next(1, &pOperator, NULL))
			{
				iCount++;

				CComHeapPtr<OLECHAR> pszId;
				hr = pOperator->GetId(&pszId);
				Assert::AreEqual(S_OK, hr, "GetId failed.");

				if (plistCstr != NULL)
					plistCstr->push_back(CString(pszId));

				pOperator.Release();
			}

			pOperators.Release();

			if (NULL != piCount)
				*piCount = iCount;

			return hr;
		}

	};
}
