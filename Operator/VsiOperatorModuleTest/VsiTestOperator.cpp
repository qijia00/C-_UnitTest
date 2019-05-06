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
#include <VsiPdmModule.h>
#include <VsiOperatorModule.h>

using namespace System;
using namespace System::Text;
using namespace System::Collections::Generic;
using namespace Microsoft::VisualStudio::TestTools::UnitTesting;

namespace VsiOperatorModuleTest
{
	[TestClass]
	public ref class VsiTestOperator
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

		//Test IVsiOperator

		/// <summary>
		/// <step>Create Operator</step>
		/// <results>Operator created</results>
		/// </summary>
		[TestMethod]
		void TestCreate()
		{
			CComPtr<IVsiOperator> pOperator;

			HRESULT hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
		};

		/// <summary>
		/// <step>Set Operator identifier to "01234567890123456789"</step>
		/// <step>Retrieve Operator identifier</step>
		/// <results>The Operator identifier is "01234567890123456789"</results>
		/// </summary>
		[TestMethod]
		void TestOperatorId()
		{
			CComPtr<IVsiOperator> pOperator;

			HRESULT hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CString strSetId(L"01234567890123456789");

			hr = pOperator->SetId(strSetId);
			Assert::AreEqual(S_OK, hr, "SetId failed.");

			CComHeapPtr<OLECHAR> pszGetId;
			hr = pOperator->GetId(&pszGetId);
			Assert::AreEqual(S_OK, hr, "GetId failed.");

			CString strGetId(pszGetId);
			Assert::IsTrue(strGetId == strSetId, "Identifier not match.");
		};

		/// <summary>
		/// <step>Set Operator name to "Operator 01234567890123456789"</step>
		/// <step>Retrieve Operator name</step>
		/// <results>The Operator name is "Operator 01234567890123456789"</results>
		/// </summary>
		[TestMethod]
		void TestOperatorName()
		{
			CComPtr<IVsiOperator> pOperator;

			HRESULT hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CString strSetName(L"Operator 01234567890123456789");

			hr = pOperator->SetName(strSetName);
			Assert::AreEqual(S_OK, hr, "SetName failed.");

			CComHeapPtr<OLECHAR> pszGetName;
			hr = pOperator->GetName(&pszGetName);
			Assert::AreEqual(S_OK, hr, "GetName failed.");

			CString strGetName(pszGetName);
			Assert::IsTrue(strGetName == strSetName, "Name not match.");
		}

		/// <summary>
		/// <step>Call HasPassword()<results>Failed since no operator name is not allowed</results></step>
		/// <step>Set Operator password to "Password 01234567890123456789"</step>
		/// <step>Retrieve Operator password<results>The Operator password is "Password 01234567890123456789"</results></step>
		/// <step>Call HasPassword()<results>Has password</results></step>
		/// </summary>
		[TestMethod]
		void TestOperatorPassword()
		{
			CComPtr<IVsiOperator> pOperator;

			HRESULT hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperator->HasPassword();
			Assert::AreEqual(E_FAIL, hr, "HasPassword failed."); //no operator name is not allowed

			hr = pOperator->SetName(L"no operator name is not allowed");
			Assert::AreEqual(S_OK, hr, "SetPassword failed.");

			CString strSetPassword(L"Password 01234567890123456789");

			hr = pOperator->SetPassword(strSetPassword, TRUE);
			Assert::AreEqual(S_OK, hr, "SetPassword failed.");

			CComHeapPtr<OLECHAR> pszGetPassword;
			hr = pOperator->GetPassword(&pszGetPassword);
			Assert::AreEqual(S_OK, hr, "GetPassword failed.");

			CString strGetPassword(pszGetPassword);
			Assert::IsTrue(strGetPassword == strSetPassword, "Password not match.");

			hr = pOperator->HasPassword();
			Assert::AreEqual(S_OK, hr, "HasPassword failed.");
		}

		/// <summary>
		/// <step>CreateOperator with hashed password<results>Success</results></step>
		/// <step>SaveOperator to UnitTestOperatorManager2.xml<results>Success</results></step>
		/// <step>LoadOperator from UnitTestOperatorManager2.xml<results>Success</results></step>
		/// <step>Call TestPassword with same password<results>Success</results></step>
		/// <step>Call TestPassword with different password<results>Fail - expected</results></step>
		/// </summary>
		[TestMethod]
		void TestOperatorTestPassword()
		{
			//Create an instance of Operator
			CComPtr<IVsiOperator> pOperator;

			HRESULT hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			//Set a name to the Operator
			CString strName(L"OperatorNameCanNotBeEmpty");
			hr = pOperator->SetName(strName);
			Assert::AreEqual(S_OK, hr, "SetName failed.");

			//Set a password to the Operator
			CString strPassword(L"Password 01234567890123456789");
			hr = pOperator->SetPassword(strPassword, FALSE); //hash the password
			Assert::AreEqual(S_OK, hr, "SetPassword failed.");

			//Create an instance of OperatorManager
			CComPtr<IVsiApp> pApp;
			CreateApp(&pApp);
			Assert::IsTrue(pApp != NULL, "Create App failed.");

			CComPtr<IVsiOperatorManager> pOperatorMgr;

			hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			//Add the Operator to the OperatorManager
			hr = pOperatorMgr->AddOperator(pOperator, NULL);
			Assert::AreEqual(S_OK, hr, "AddOperator failed.");
			pOperator.Release();

			//switch to PersistOperatorManager
			CComQIPtr<IVsiPersistOperatorManager> pPersistOperatorMgr(pOperatorMgr);
			Assert::IsTrue(NULL != pPersistOperatorMgr, "IVsiPersistOperatorManager interface missing.");

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			// Save operator using UnitTestOperatorManager2.xml file
			{
				WCHAR szConfigFile[MAX_PATH];
				_snwprintf_s(
					szConfigFile,
					_countof(szConfigFile),
					L"%s\\UnitTestOperatorManager2.xml",
					pszTestDir);

				hr = pPersistOperatorMgr->SaveOperator(szConfigFile);
				Assert::AreEqual(S_OK, hr, "SaveOperator failed.");
			}

			//In order to make sure LoadOperator does read the correct value back from the xml
			//but not from  the operator already contained by the PersistOperatorManager
			//Remove PersistOperatorManager, i.e., Remove OperatorManager
			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
			pOperatorMgr.Release();

			//Recreate an instance of OperatorManager then switch to PersistOperatorManager
			hr = pOperatorMgr.CoCreateInstance(__uuidof(VsiOperatorManager));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperatorMgr->Initialize(pApp, L"");
			Assert::AreEqual(S_OK, hr, "Initialize failed.");

			pPersistOperatorMgr = pOperatorMgr;
			Assert::IsTrue(NULL != pPersistOperatorMgr, "IVsiPersistOperatorManager interface missing.");

			// Load operator using UnitTestOperatorManager2.xml file
			{
				WCHAR szConfigFile[MAX_PATH];
				_snwprintf_s(
					szConfigFile,
					_countof(szConfigFile),
					L"%s\\UnitTestOperatorManager2.xml",
					pszTestDir);
				hr = pPersistOperatorMgr->LoadOperator(szConfigFile);
				Assert::AreEqual(S_OK, hr, "LoadOperator failed.");

				hr = pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, strName, &pOperator);
				Assert::AreEqual(S_OK, hr, "GetOperator failed.");				

				hr = pOperator->TestPassword(L"Password 01234567890123456789"); //TestPassword will always hash its input password
				Assert::AreEqual(S_OK, hr, "TestPassword failed.");

				hr = pOperator->TestPassword(L"Password01234567890123456789");
				Assert::AreEqual(S_FALSE, hr, "TestPassword succeeded - unexpected.");
			}
			
			// Clean up
			hr = pOperatorMgr->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			CloseApp(pApp);
		}

		/// <summary>
		/// <step>Set Operator group to "Group A1"</step>
		/// <step>Retrieve Operator group</step>
		/// <results>The Operator group is "Group A1"</results>
		/// </summary>
		[TestMethod]
		void TestOperatorGroup()
		{
			CComPtr<IVsiOperator> pOperator;

			HRESULT hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			CString strSetGroup(L"Group A1");

			hr = pOperator->SetGroup(strSetGroup);
			Assert::AreEqual(S_OK, hr, "SetGroup failed.");

			CComHeapPtr<OLECHAR> pszGetGroup;
			hr = pOperator->GetGroup(&pszGetGroup);
			Assert::AreEqual(S_OK, hr, "GetGroup failed.");

			CString strGetGroup(pszGetGroup);
			Assert::IsTrue(strGetGroup == strSetGroup, "Group not match.");
		}

		/// <summary>
		/// <step>Set Operator type to VSI_OPERATOR_TYPE_STANDARD</step>
		/// <step>Retrieve Operator type<results>The Operator type is VSI_OPERATOR_TYPE_STANDARD</results></step>
		/// <step>Set Operator type to VSI_OPERATOR_TYPE_ADMIN</step>
		/// <step>Retrieve Operator type<results>The Operator type is VSI_OPERATOR_TYPE_ADMIN</results></step>
		/// </summary>
		[TestMethod]
		void TestOperatorType()
		{
			CComPtr<IVsiOperator> pOperator;

			HRESULT hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperator->SetType(VSI_OPERATOR_TYPE_STANDARD);
			Assert::AreEqual(S_OK, hr, "SetType failed.");

			VSI_OPERATOR_TYPE dwType(VSI_OPERATOR_TYPE_NONE);
			hr = pOperator->GetType(&dwType);
			Assert::AreEqual(S_OK, hr, "GetType failed.");

			Assert::IsTrue(dwType == VSI_OPERATOR_TYPE_STANDARD, "Type not standard.");

			hr = pOperator->SetType(VSI_OPERATOR_TYPE_ADMIN);
			Assert::AreEqual(S_OK, hr, "SetType failed.");

			hr = pOperator->GetType(&dwType);
			Assert::AreEqual(S_OK, hr, "GetType failed.");

			Assert::IsTrue(dwType == VSI_OPERATOR_TYPE_ADMIN, "Type not standard.");
		}

		/// <summary>
		/// <step>Set Operator state to VSI_OPERATOR_STATE_ENABLE</step>
		/// <step>Retrieve Operator state<results>The Operator state is VSI_OPERATOR_STATE_ENABLE</results></step>
		/// <step>Set Operator state to VSI_OPERATOR_STATE_DISABLE</step>
		/// <step>Retrieve Operator state<results>The Operator type is VSI_OPERATOR_STATE_DISABLE</results></step>
		/// </summary>
		[TestMethod]
		void TestOperatorState()
		{
			CComPtr<IVsiOperator> pOperator;

			HRESULT hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperator->SetState(VSI_OPERATOR_STATE_ENABLE);
			Assert::AreEqual(S_OK, hr, "SetType failed.");

			VSI_OPERATOR_STATE dwState(VSI_OPERATOR_STATE_DISABLE);
			hr = pOperator->GetState(&dwState);
			Assert::AreEqual(S_OK, hr, "GetState failed.");

			Assert::IsTrue(dwState == VSI_OPERATOR_STATE_ENABLE, "State not enabled.");

			hr = pOperator->SetState(VSI_OPERATOR_STATE_DISABLE);
			Assert::AreEqual(S_OK, hr, "SetType failed.");

			hr = pOperator->GetState(&dwState);
			Assert::AreEqual(S_OK, hr, "GetState failed.");

			Assert::IsTrue(dwState == VSI_OPERATOR_STATE_DISABLE, "State not disabled.");
		}

		/// <summary>
		/// <step>Set Default Study Access to VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE</step>
		/// <step>Retrieve Default Study Access<results>The Default Study Access is VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE</results></step>
		/// <step>Set Default Study Access to VSI_OPERATOR_STATE_DISABLE</step>
		/// <step>Retrieve Default Study Access<results>The Default Study Access is VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PUBLIC</results></step>
		/// </summary>
		[TestMethod]
		void TestStudyAccess()
		{
			CComPtr<IVsiOperator> pOperator;

			HRESULT hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pOperator->SetDefaultStudyAccess(VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE);
			Assert::AreEqual(S_OK, hr, "SetDefaultStudyAccess failed.");

			VSI_OPERATOR_DEFAULT_STUDY_ACCESS dwState(VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PUBLIC);
			hr = pOperator->GetDefaultStudyAccess(&dwState);
			Assert::AreEqual(S_OK, hr, "GetDefaultStudyAccess failed.");

			Assert::IsTrue(dwState == VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE, "Default study access not private.");

			hr = pOperator->SetDefaultStudyAccess(VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PUBLIC);
			Assert::AreEqual(S_OK, hr, "SetDefaultStudyAccess failed.");

			hr = pOperator->GetDefaultStudyAccess(&dwState);
			Assert::AreEqual(S_OK, hr, "GetDefaultStudyAccess failed.");

			Assert::IsTrue(dwState == VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PUBLIC, "Default study access not public.");
		}
		
		/// <summary>
		/// <step>Set Session Info to a struct data</step>
		/// <step>Retrieve Session Info<results>The Session Info is the same struct data</results></step>
		/// </summary>
		[TestMethod]
		void TestSessionInfo()
		{
			CComPtr<IVsiOperator> pOperator;

			HRESULT hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");
						
			//Set Session Info
			VSI_OPERATOR_SESSION_INFO info;
			SYSTEMTIME stSetIn, stSetOut, stGetIn, stGetOut;
			GetSystemTime(&stSetIn);
			info.stLastLogin = stSetIn;
			Sleep(1000);
			GetSystemTime(&stSetOut);
			info.stLastLogout = stSetOut;
			DWORD dwSetTotalSeconds = GetTimeDiffInSeconds(&stSetOut, &stSetIn);
			info.dwTotalSeconds = dwSetTotalSeconds;
			hr = pOperator->SetSessionInfo(&info);
			Assert::AreEqual(S_OK, hr, "SetSessionInfo failed.");

			//Get Session Info
			GetSystemTime(&stGetIn);
			info.stLastLogin = stGetIn;
			Sleep(3000);
			GetSystemTime(&stGetOut);
			info.stLastLogout = stGetOut;
			DWORD dwGetTotalSeconds = GetTimeDiffInSeconds(&stGetOut, &stGetIn);
			info.dwTotalSeconds = dwGetTotalSeconds;
			hr = pOperator->GetSessionInfo(&info);
			Assert::AreEqual(S_OK, hr, "GetSessionInfo failed."); 

			Assert::IsTrue(0 == memcmp(&(info.stLastLogin), &stSetIn, sizeof(SYSTEMTIME)), "last login time in session info not same.");
			Assert::IsTrue(0 == memcmp(&(info.stLastLogout), &stSetOut, sizeof(SYSTEMTIME)), "last logout time in session info not same.");
			Assert::IsTrue(info.dwTotalSeconds == dwSetTotalSeconds, "total seconds in session info not same.");
		}

		//Test IVsiOperatorManager
		//LoadXML is used by LoadOperator, which is tested by TestLoadOperator() in VsiTestOperatorManager.cpp
		//LoadXML is used by SaveOperator, which is tested by TestSaveOperator() in VsiTestOperatorManager.cpp

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

		DWORD GetTimeDiffInSeconds(LPSYSTEMTIME lpstNow, LPSYSTEMTIME lpstLast)
		{
			DWORD dwTimeDiff = 0;
			//FILETIME is 64-bit representing the number of 100-nanosecond intervals since January 1, 1601.
			FILETIME ftNow, ftLast;
			
			//Convert the SYSTEMTIME structure to a FILETIME structure
			SystemTimeToFileTime(lpstNow, &ftNow);
			SystemTimeToFileTime(lpstLast, &ftLast);

			//Copy the resulting FILETIME structure to a ULARGE_INTEGER structure
			ULARGE_INTEGER uliNow, uliLast; //ULARGE_INTEGER LowPart and HighPart are 32-bit DWORD
			uliNow.HighPart = ftNow.dwHighDateTime;
			uliNow.LowPart = ftNow.dwLowDateTime;
			uliLast.HighPart = ftLast.dwHighDateTime;
			uliLast.LowPart = ftLast.dwLowDateTime;

			// get time difference in seconds
			ULARGE_INTEGER uliTimeDiff;
			if (uliNow.QuadPart > uliLast.QuadPart) //time can be big, so use QuadPart (64-bit ULONGLONG) 
			{
				uliTimeDiff.QuadPart = uliNow.QuadPart - uliLast.QuadPart; //in 100 nanoseconds
				uliTimeDiff.QuadPart = uliTimeDiff.QuadPart / (ULONGLONG)(1E7); //in seconds (1 second is 1E7 100 nanoseconds)
				//uliTimeDiff.QuadPart = uliTimeDiff.QuadPart / 60; //in minutes (1 minute is 60 seconds)

				dwTimeDiff = uliTimeDiff.LowPart; //assume the time diff won't be that big, so only LowPart (32-bit DWORD) is needed.
			}

			return dwTimeDiff;
		}
	};
}
