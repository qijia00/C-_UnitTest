/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiMessageLogTest.cpp
**
**	Description:
**		VsiMessageLogTest Test
**
********************************************************************************/
#include "stdafx.h"
#include <VsiUnitTest.h>
#include <Common\VsiMessageLogModule\resource.h>
#include <VsiMessageLogModule.h>
#include <VsiException.h>
#include <VsiGdiPlus.h>

using namespace System;
using namespace System::Text;
using namespace System::Collections::Generic;
using namespace Microsoft::VisualStudio::TestTools::UnitTesting;


namespace VsiTestMessageLogModule
{
	[TestClass]
	public ref class VsiTestMessageLogViewer
	{
	private:
		TestContext^ testContextInstance;
		ULONG_PTR m_cookie;
		CVsiGdiPlus *m_pGdiPlus;
		IVsiMessageLog *m_pMsgLog;
		IVsiMessageLogViewer *m_pMsgLogViewer;
		HANDLE m_hLogProcess;

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
			m_pMsgLog = NULL;
			m_pMsgLogViewer = NULL;
			m_hLogProcess = NULL;

			m_pGdiPlus = new CVsiGdiPlus();
			m_pGdiPlus->Initialize();

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			ACTCTX actCtx;
			memset((void*)&actCtx, 0, sizeof(ACTCTX));
			actCtx.cbSize = sizeof(ACTCTX);
			actCtx.dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID | ACTCTX_FLAG_RESOURCE_NAME_VALID;
			actCtx.lpSource = L"VsiMessageLogModule.dll";
			actCtx.lpAssemblyDirectory = pszTestDir;
			actCtx.lpResourceName = MAKEINTRESOURCE(2);

			HANDLE hCtx = CreateActCtx(&actCtx);
			Assert::IsTrue(INVALID_HANDLE_VALUE != hCtx, "CreateActCtx failed.");

			ULONG_PTR cookie;
			BOOL bRet = ActivateActCtx(hCtx, &cookie);
			Assert::IsTrue(FALSE != bRet, "ActivateActCtx failed.");

			m_cookie = cookie;

			HRESULT hr;
			CComPtr<IVsiMessageLog> pMsgLog;
			CComPtr<IVsiMessageLogViewer> pMsgLogViewer;

			hr = pMsgLog.CoCreateInstance(__uuidof(VsiMessageLog));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			hr = pMsgLogViewer.CoCreateInstance(__uuidof(VsiMessageLogViewer));
			Assert::AreEqual(S_OK, hr, "CoCreateInstance failed.");

			m_pMsgLog = pMsgLog.Detach();
			m_pMsgLogViewer = pMsgLogViewer.Detach();
		};
		//
		//Use TestCleanup to run code after each test has run
		[TestCleanup()]
		void MyTestCleanup()
		{
			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, NULL);
			if (wndViewer != NULL)
			{
				wndViewer.PostMessage(WM_CLOSE);
				HRESULT hr = VsiUnitTest::WaitForNoWindow(
					2000, VS_WND_CLASS_LOG_VIEWER, NULL);
				if (S_OK != hr)
				{
					if (NULL != m_hLogProcess)
						TerminateProcess(m_hLogProcess, 0);
				}
			}

			if (m_pMsgLogViewer != NULL)
			{
				m_pMsgLogViewer->Deactivate();
				m_pMsgLogViewer->Uninitialize();
				m_pMsgLogViewer->Release();
				m_pMsgLogViewer = NULL;
			}

			if (m_pMsgLog != NULL)
			{
				m_pMsgLog->Uninitialize();
				m_pMsgLog->Release();
				m_pMsgLog = NULL;
			}

			if (NULL != m_hLogProcess)
			{
				CloseHandle(m_hLogProcess);
				m_hLogProcess = NULL;
			}

			if (0 != m_cookie)
			{
				DeactivateActCtx(0, m_cookie);
				m_cookie = 0;
			}
			m_pGdiPlus->Uninitialize();
			delete m_pGdiPlus;
			m_pGdiPlus = NULL;
		};
		//
		#pragma endregion 

		/// <summary>
		/// <step>Creates Message Log Viewer</step>
		/// </summary>
		/// <results>Message Log Viewer created</results>
		[TestMethod]
		void TestCreate()
		{
			Assert::IsTrue(m_pMsgLog != NULL, "Failed to create Message Log.");
			Assert::IsTrue(m_pMsgLogViewer != NULL, "Failed to create Message Log Viewer.");
		};

		/// <summary>
		/// <step>Initializes Message Log Viewer in Review Mode</step>
		/// </summary>
		/// <results>Message Log Viewer initialized</results>
		[TestMethod]
		void TestInitializeReview()
		{
			InitMsgLogViewer(VSI_LOG_FLAG_REVIEW);
			UninitMsgLogViewer();
		}

		/// <summary>
		/// <step>Initializes Message Log Viewer in Review Mode</step>
		/// <step>Un-initializes Message Log Viewer</step>
		/// </summary>
		/// <results>Message Log Viewer un-initialized</results>
		[TestMethod]
		void TestUninitialize()
		{
			InitMsgLogViewer(VSI_LOG_FLAG_REVIEW);
			UninitMsgLogViewer();
		}

		/// <summary>
		/// <step>Activates Message Log Viewer</step>
		/// <step>Searches Message Log Viewer window class<results>Available</results></step>
		/// <step>Deactivates Message Log Viewer</step>
		/// <step>Searches Message Log Viewer window class<results>Not available</results></step>
		/// </summary>
		[TestMethod]
		void TestCreateUI()
		{
			HRESULT hr = S_OK;

			InitMsgLogViewer(VSI_LOG_FLAG_REVIEW);

			hr = m_pMsgLogViewer->Activate(L"UnitTest");
			Assert::AreEqual(S_OK, hr, "Activate failed.");

			HWND hwnd = FindWindow(VS_WND_CLASS_LOG_VIEWER, L"UnitTest");
			Assert::IsTrue(NULL != hwnd, "Activate failed.");

			SetForegroundWindow(hwnd);

			hr = m_pMsgLogViewer->Deactivate();
			Assert::AreEqual(S_OK, hr, "Deactivate failed.");

			hwnd = FindWindow(VS_WND_CLASS_LOG_VIEWER, L"UnitTest");
			Assert::IsTrue(NULL == hwnd, "Deactivate failed.");

			UninitMsgLogViewer();
		}

		/// <summary>
		/// <step>Initializes Message Log Viewer in Review Mode</step>
		/// <step>Shows Message Log Viewer</step>
		/// <step>Checks menu layout (see SDD)</step>
		/// </summary>
		/// <results>No error</results>
		[TestMethod]
		void TestReviewModeMenu()
		{
			HRESULT hr;
			int i = 0;

			InitMsgLogViewer(VSI_LOG_FLAG_REVIEW);

			hr = m_pMsgLogViewer->Activate(L"UnitTestReviewModeMenu");
			Assert::AreEqual(S_OK, hr, "Activate failed.");

			HWND hwnd = FindWindow(VS_WND_CLASS_LOG_VIEWER, L"UnitTestReviewModeMenu");
			Assert::IsTrue(NULL != hwnd, "Activate failed.");

			SetForegroundWindow(hwnd);

			HMENU hmenu = GetMenu(hwnd);
			Assert::IsTrue(NULL != hmenu, "GetMenu failed.");

			// Log menu
			{
				HMENU hmenuLog = GetSubMenu(hmenu, 0);
				Assert::IsTrue(NULL != hmenuLog, "GetSubMenu failed.");

				Assert::AreEqual((UINT)ID_LOG_OPEN, GetMenuItemID(hmenuLog, i = 0));
				Assert::AreEqual((UINT)ID_LOG_SAVEAS, GetMenuItemID(hmenuLog, ++i));
				Assert::AreEqual((UINT)0, GetMenuItemID(hmenuLog, ++i));
				Assert::AreEqual((UINT)ID_LOG_EXIT, GetMenuItemID(hmenuLog, ++i));
				Assert::AreEqual((UINT)-1, GetMenuItemID(hmenuLog, ++i));
			}

			// Edit menu
			{
				HMENU hmenuView = GetSubMenu(hmenu, 1);
				Assert::IsTrue(NULL != hmenuView, "GetSubMenu failed.");

				Assert::AreEqual((UINT)ID_EDIT_COPY, GetMenuItemID(hmenuView, i = 0));
				Assert::AreEqual((UINT)ID_EDIT_COPY_SHORT, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)0, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)ID_EDIT_GOTO, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)-1, GetMenuItemID(hmenuView, ++i));
			}

			// View menu
			{
				HMENU hmenuView = GetSubMenu(hmenu, 2);
				Assert::IsTrue(NULL != hmenuView, "GetSubMenu failed.");

				Assert::AreEqual((UINT)ID_VIEW_ALWAYSONTOP, GetMenuItemID(hmenuView, i = 0));
				Assert::AreEqual((UINT)0, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)ID_VIEW_MESSAGELOG, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)ID_VIEW_TIMETRACE, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)ID_VIEW_EVENTTRACE, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)0, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)ID_VIEW_SHOWEVENTTIMES, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)ID_VIEW_REVERSETIMEORDER, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)-1, GetMenuItemID(hmenuView, ++i));
			}

			// Filter menu
			{
				HMENU hmenuFilter = GetSubMenu(hmenu, 3);
				Assert::IsTrue(NULL != hmenuFilter, "GetSubMenu failed.");

				Assert::AreEqual((UINT)ID_FILTER_ALLMESSAGES, GetMenuItemID(hmenuFilter, i = 0));
				Assert::AreEqual((UINT)ID_FILTER_ERROR_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_WARNING_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_INFO_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_ACTIVITY_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_THREAD_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_EVENT_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_FIRMWARE_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)-1, GetMenuItemID(hmenuFilter, ++i));
			}

			//Help menu
			{
				HMENU hmenuHelp = GetSubMenu(hmenu, 4);
				Assert::IsTrue(NULL != hmenuHelp, "GetSubMenu failed.");

				Assert::AreEqual((UINT)ID_ABOUT_MESSAGELOGVIEWER, GetMenuItemID(hmenuHelp, i = 0));
				Assert::AreEqual((UINT)-1, GetMenuItemID(hmenuHelp, ++i));
			}

			// Close
			hr = m_pMsgLogViewer->Deactivate();
			Assert::AreEqual(S_OK, hr, "Deactivate failed.");

			UninitMsgLogViewer();
		}

		/// <summary>
		/// <step>Initializes Message Log Viewer in Log Mode</step>
		/// <step>Shows Message Log Viewer</step>
		/// <step>Checks menu layout (see SDD)</step>
		/// </summary>
		/// <results>No error</results>
		[TestMethod]
		void TestLogModeMenu()
		{
			HRESULT hr;
			int i = 0;

			InitMsgLogViewer(VSI_LOG_FLAG_NONE);

			hr = m_pMsgLogViewer->Activate(L"UnitTestLogModeMenu");
			Assert::AreEqual(S_OK, hr, "Activate failed.");

			HWND hwnd = FindWindow(VS_WND_CLASS_LOG_VIEWER, L"UnitTestLogModeMenu");
			Assert::IsTrue(NULL != hwnd, "Activate failed.");

			SetForegroundWindow(hwnd);

			HMENU hmenu = GetMenu(hwnd);
			Assert::IsTrue(NULL != hmenu, "GetMenu failed.");

			// Log menu
			{
				HMENU hmenuLog = GetSubMenu(hmenu, 0);
				Assert::IsTrue(NULL != hmenuLog, "GetSubMenu failed.");

				Assert::AreEqual((UINT)ID_LOG_CLEAR, GetMenuItemID(hmenuLog, i = 0));
				Assert::AreEqual((UINT)0, GetMenuItemID(hmenuLog, ++i));
				Assert::AreEqual((UINT)ID_LOG_MANAGEARCHIVES, GetMenuItemID(hmenuLog, ++i));
				Assert::AreEqual((UINT)ID_LOG_SAVEAS, GetMenuItemID(hmenuLog, ++i));
				Assert::AreEqual((UINT)ID_LOG_NOERRORPOPUPS, GetMenuItemID(hmenuLog, ++i));
				Assert::AreEqual((UINT)ID_LOG_LOCKLOG, GetMenuItemID(hmenuLog, ++i));
				Assert::AreEqual((UINT)0, GetMenuItemID(hmenuLog, ++i));
				Assert::AreEqual((UINT)ID_LOG_EXIT, GetMenuItemID(hmenuLog, ++i));
				Assert::AreEqual((UINT)-1, GetMenuItemID(hmenuLog, ++i));
			}

			// Edit menu
			{
				HMENU hmenuView = GetSubMenu(hmenu, 1);
				Assert::IsTrue(NULL != hmenuView, "GetSubMenu failed.");

				Assert::AreEqual((UINT)ID_EDIT_COPY, GetMenuItemID(hmenuView, i = 0));
				Assert::AreEqual((UINT)ID_EDIT_COPY_SHORT, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)0, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)ID_EDIT_GOTO, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)-1, GetMenuItemID(hmenuView, ++i));
			}

			// View menu
			{
				HMENU hmenuView = GetSubMenu(hmenu, 2);
				Assert::IsTrue(NULL != hmenuView, "GetSubMenu failed.");

				Assert::AreEqual((UINT)ID_VIEW_ALWAYSONTOP, GetMenuItemID(hmenuView, i = 0));
				Assert::AreEqual((UINT)0, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)ID_VIEW_THREADACTIVITY, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)ID_VIEW_MESSAGELOG, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)0, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)ID_VIEW_REVERSETIMEORDER, GetMenuItemID(hmenuView, ++i));
				Assert::AreEqual((UINT)-1, GetMenuItemID(hmenuView, ++i));
			}

			// Filter menu
			{
				HMENU hmenuFilter = GetSubMenu(hmenu, 3);
				Assert::IsTrue(NULL != hmenuFilter, "GetSubMenu failed.");

				Assert::AreEqual((UINT)ID_FILTER_ALLMESSAGES, GetMenuItemID(hmenuFilter, i = 0));
				Assert::AreEqual((UINT)ID_FILTER_ERROR_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_WARNING_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_INFO_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_ACTIVITY_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_THREAD_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_EVENT_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_FIRMWARE_ONLY, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)0, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_ALLOW_DEBUG, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)ID_FILTER_ALLOW_EVENT, GetMenuItemID(hmenuFilter, ++i));
				Assert::AreEqual((UINT)-1, GetMenuItemID(hmenuFilter, ++i));
			}

			//Help menu
			{
				HMENU hmenuHelp = GetSubMenu(hmenu, 4);
				Assert::IsTrue(NULL != hmenuHelp, "GetSubMenu failed.");

				Assert::AreEqual((UINT)ID_ABOUT_MESSAGELOGVIEWER, GetMenuItemID(hmenuHelp, i = 0));
				Assert::AreEqual((UINT)-1, GetMenuItemID(hmenuHelp, ++i));
			}

			// Close
			hr = m_pMsgLogViewer->Deactivate();
			Assert::AreEqual(S_OK, hr, "Deactivate failed.");

			UninitMsgLogViewer();
		}

		/// <summary>
		/// <step>Shows Message Log Viewer</step>
		/// <step>Retrieves "always on top" style<results>Off</results></step>
		/// <step>Invokes "Always On Top" command</step>
		/// <step>Retrieves "always on top" style<results>On</results></step>
		/// <step>Invokes "Always On Top" command</step>
		/// <step>Retrieves "always on top" style<results>Off</results></step>
		/// </summary>
		[TestMethod]
		void TestAlwaysOnTopStyle()
		{
			HRESULT hr;
			DWORD dwStyle;

			InitMsgLogViewer(VSI_LOG_FLAG_NONE);

			// Shows Message Log Viewer
			hr = m_pMsgLogViewer->Activate(L"UnitTestThreadActivityView");
			Assert::AreEqual(S_OK, hr, "Activate failed.");

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, L"UnitTestThreadActivityView");
			Assert::IsTrue(NULL != wndViewer, "Activate failed.");

			SetForegroundWindow(wndViewer);

			dwStyle = wndViewer.GetExStyle();
			Assert::AreEqual((ULONG)0, (WS_EX_TOPMOST & dwStyle), "Incorrect Top Most style.");

			// Invokes "Always On Top" command
			wndViewer.SendMessage(WM_COMMAND, MAKEWPARAM(ID_VIEW_ALWAYSONTOP, 0));

			dwStyle = wndViewer.GetExStyle();
			Assert::IsTrue(0 != (WS_EX_TOPMOST & dwStyle), "Missing Top Most style.");

			// Invokes "Always On Top" command
			wndViewer.SendMessage(WM_COMMAND, MAKEWPARAM(ID_VIEW_ALWAYSONTOP, 0));

			dwStyle = wndViewer.GetExStyle();
			Assert::AreEqual((ULONG)0, (WS_EX_TOPMOST & dwStyle), "Incorrect Top Most style.");

			// Close
			hr = m_pMsgLogViewer->Deactivate();
			Assert::AreEqual(S_OK, hr, "Deactivate failed.");

			UninitMsgLogViewer();
		}

		/// <summary>
		/// <step>Shows Message Log Viewer</step>
		/// <step>Switches to Thread Activity View mode</step>
		/// </summary>
		/// <results>The Thread Activity View window is visible</results>
		[TestMethod]
		void TestThreadActivityView()
		{
			HRESULT hr;

			InitMsgLogViewer(VSI_LOG_FLAG_NONE);

			// Shows Message Log Viewer
			hr = m_pMsgLogViewer->Activate(L"UnitTestThreadActivityView");
			Assert::AreEqual(S_OK, hr, "Activate failed.");

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, L"UnitTestThreadActivityView");
			Assert::IsTrue(NULL != wndViewer, "Activate failed.");

			SetForegroundWindow(wndViewer);

			// Switches to Thread Activity View mode
			wndViewer.SendMessage(WM_COMMAND, MAKEWPARAM(ID_VIEW_THREADACTIVITY, 0));

			CWindow wndActList = wndViewer.GetDlgItem(IDC_THREAD_ACTIVITY_LIST);
			Assert::IsTrue(wndActList != NULL, "Activity List window missing.");

			// Close
			hr = m_pMsgLogViewer->Deactivate();
			Assert::AreEqual(S_OK, hr, "Deactivate failed.");

			UninitMsgLogViewer();
		}

		/// <summary>
		/// <step>Shows Message Log Viewer</step>
		/// <step>Switches to Message Log View mode</step>
		/// </summary>
		/// <results>The Message Log View window is visible</results>
		[TestMethod]
		void TestMessageLogView()
		{
			HRESULT hr;

			InitMsgLogViewer(VSI_LOG_FLAG_NONE);

			// Shows Message Log Viewer
			hr = m_pMsgLogViewer->Activate(L"UnitTestMessageLogView");
			Assert::AreEqual(S_OK, hr, "Activate failed.");

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, L"UnitTestMessageLogView");
			Assert::IsTrue(NULL != wndViewer, "Activate failed.");

			SetForegroundWindow(wndViewer);

			// Switches to Message Log View mode
			wndViewer.SendMessage(WM_COMMAND, MAKEWPARAM(ID_VIEW_MESSAGELOG, 0));

			CWindow wndMsgList = wndViewer.GetDlgItem(IDC_MESSAGE_LIST);
			Assert::IsTrue(wndMsgList != NULL, "Message List window missing.");

			// Close
			hr = m_pMsgLogViewer->Deactivate();
			Assert::AreEqual(S_OK, hr, "Deactivate failed.");

			UninitMsgLogViewer();
		}

		/// <summary>
		/// <step>Shows Message Log Viewer</step>
		/// <step>Switches to Time Trace View mode</step>
		/// </summary>
		/// <results>The Message Log View window is visible</results>
		/// <results>The Time Trace View window is visible</results>
		[TestMethod]
		void TestTimeTraceView()
		{
			HRESULT hr = S_OK;

			InitMsgLogViewer(VSI_LOG_FLAG_NONE);

			// Shows Message Log Viewer
			hr = m_pMsgLogViewer->Activate(L"UnitTestTimeTraceView");
			Assert::AreEqual(S_OK, hr, "Activate failed.");

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, L"UnitTestTimeTraceView");
			Assert::IsTrue(NULL != wndViewer, "Activate failed.");

			SetForegroundWindow(wndViewer);

			// Switches to Time Trace View mode
			wndViewer.SendMessage(WM_COMMAND, MAKEWPARAM(ID_VIEW_TIMETRACE, 0));

			CWindow wndMsgList = wndViewer.GetDlgItem(IDC_MESSAGE_LIST);
			Assert::IsTrue(wndMsgList != NULL, "Message List window missing.");

			CWindow wndTimeTraceList = wndViewer.GetDlgItem(IDC_TIMETRACE_LIST);
			Assert::IsTrue(wndTimeTraceList != NULL, "Time Trace window missing.");

			// Close
			hr = m_pMsgLogViewer->Deactivate();
			Assert::AreEqual(S_OK, hr, "Deactivate failed.");

			UninitMsgLogViewer();
		}

		/// <summary>
		/// <step>Open test file TestMessageLog001.txt</step>
		/// <step>Retrieves message count in Message Log View window</step>
		/// </summary>
		/// <results>16 messages in Message Log View window</results>
		[TestMethod]
		void TestOpenLogFile()
		{
			HRESULT hr = S_OK;

			LaunchLogViewer();

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, NULL);
			Assert::IsTrue(wndViewer != NULL, "Viewer not found.");

			SetForegroundWindow(wndViewer);

			// Open file
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strDataPath;
			strDataPath.Format(L"%s\\TestMessageLog001.txt", pszTestDir);

			wndViewer.PostMessage(WM_COMMAND, MAKEWPARAM(ID_LOG_OPEN, 0));

			VsiUnitTest::WaitForDisabled(2000, wndViewer);

			Assert::IsTrue(!wndViewer.IsWindowEnabled());

			Sleep(500);

			// Send path
			hr = VsiUnitTest::SendKeys(strDataPath, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			WORD wKey = VK_EXECUTE;
			hr = VsiUnitTest::SendKeys(&wKey, 1, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			VsiUnitTest::WaitForEnabled(2000, wndViewer);

			Sleep(2000);

			// Retrieves count
			CWindow wndMsgList = wndViewer.GetDlgItem(IDC_MESSAGE_LIST);
			Assert::IsTrue(wndMsgList != NULL, "Message List window not found.");

			int iCount = ListView_GetItemCount(wndMsgList);
			Assert::AreEqual(16, iCount, "Count not matched.");

			// Close viewer
			wndViewer.PostMessage(WM_CLOSE);

			VsiUnitTest::WaitForNoWindow(2000, VS_WND_CLASS_LOG_VIEWER, NULL);
		}

		/// <summary>
		/// <step>Open test file MessageLog001.txt</step>
		/// <step>Invokes "Error Messages Only" command</step>
		/// <step>Retrieves message count in Message Log View window</step>
		/// </summary>
		/// <results>5 message in Message Log View window</results>
		[TestMethod]
		void TestFilterErrorMessageOnly()
		{
			// Shows Message Log Viewer
			LaunchLogViewer();

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, NULL);
			Assert::IsTrue(wndViewer != NULL, "Viewer not found.");

			SetForegroundWindow(wndViewer);

			// Open file
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strDataPath;
			strDataPath.Format(L"%s\\TestMessageLog001.txt", pszTestDir);

			wndViewer.PostMessage(WM_COMMAND, MAKEWPARAM(ID_LOG_OPEN, 0));

			VsiUnitTest::WaitForDisabled(2000, wndViewer);

			Assert::IsTrue(!wndViewer.IsWindowEnabled());

			Sleep(500);

			// Send path
			HRESULT hr = VsiUnitTest::SendKeys(strDataPath, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			WORD wKey = VK_EXECUTE;
			hr = VsiUnitTest::SendKeys(&wKey, 1, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			VsiUnitTest::WaitForEnabled(2000, wndViewer);

			wndViewer.PostMessage(WM_COMMAND, MAKEWPARAM(ID_FILTER_ERROR_ONLY, 0));

			Sleep(2000);

			// Retrieves count
			CWindow wndMsgList = wndViewer.GetDlgItem(IDC_MESSAGE_LIST);
			Assert::IsTrue(wndMsgList != NULL, "Message List window not found.");

			int iCount = ListView_GetItemCount(wndMsgList);
			Assert::AreEqual(5, iCount, "Count not matched.");

			// Close viewer
			wndViewer.PostMessage(WM_CLOSE);

			VsiUnitTest::WaitForNoWindow(2000, VS_WND_CLASS_LOG_VIEWER, NULL);
		}

		/// <summary>
		/// <step>Open test file MessageLog001.txt</step>
		/// <step>Invokes "Warning Messages Only" command</step>
		/// <step>Retrieves message count in Message Log View window</step>
		/// </summary>
		/// <results>4 message in Message Log View window</results>
		[TestMethod]
		void TestFilterWarningMessageOnly()
		{
			// Shows Message Log Viewer
			LaunchLogViewer();

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, NULL);
			Assert::IsTrue(wndViewer != NULL, "Viewer not found.");

			SetForegroundWindow(wndViewer);

			// Open file
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strDataPath;
			strDataPath.Format(L"%s\\TestMessageLog001.txt", pszTestDir);

			wndViewer.PostMessage(WM_COMMAND, MAKEWPARAM(ID_LOG_OPEN, 0));

			VsiUnitTest::WaitForDisabled(2000, wndViewer);

			Assert::IsTrue(!wndViewer.IsWindowEnabled());

			Sleep(500);

			// Send path
			HRESULT hr = VsiUnitTest::SendKeys(strDataPath, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			WORD wKey = VK_EXECUTE;
			hr = VsiUnitTest::SendKeys(&wKey, 1, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			VsiUnitTest::WaitForEnabled(2000, wndViewer);

			wndViewer.PostMessage(WM_COMMAND, MAKEWPARAM(ID_FILTER_WARNING_ONLY, 0));

			Sleep(2000);

			// Retrieves count
			CWindow wndMsgList = wndViewer.GetDlgItem(IDC_MESSAGE_LIST);
			Assert::IsTrue(wndMsgList != NULL, "Message List window not found.");

			int iCount = ListView_GetItemCount(wndMsgList);
			Assert::AreEqual(4, iCount, "Count not matched.");

			// Close viewer
			wndViewer.PostMessage(WM_CLOSE);

			VsiUnitTest::WaitForNoWindow(2000, VS_WND_CLASS_LOG_VIEWER, NULL);
		}

		/// <summary>
		/// <step>Open test file MessageLog001.txt</step>
		/// <step>Invokes "Info Messages Only" command</step>
		/// <step>Retrieves message count in Message Log View window</step>
		/// </summary>
		/// <results>3 message in Message Log View window</results>
		[TestMethod]
		void TestFilterInfoMessageOnly()
		{
			// Shows Message Log Viewer
			LaunchLogViewer();

			VsiUnitTest::WaitForWindow(2000, VS_WND_CLASS_LOG_VIEWER, NULL);

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, NULL);
			Assert::IsTrue(wndViewer != NULL, "Viewer not found.");

			SetForegroundWindow(wndViewer);

			// Open file
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strDataPath;
			strDataPath.Format(L"%s\\TestMessageLog001.txt", pszTestDir);

			wndViewer.PostMessage(WM_COMMAND, MAKEWPARAM(ID_LOG_OPEN, 0));

			VsiUnitTest::WaitForDisabled(2000, wndViewer);

			Assert::IsTrue(!wndViewer.IsWindowEnabled());

			Sleep(500);

			// Send path
			HRESULT hr = VsiUnitTest::SendKeys(strDataPath, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			WORD wKey = VK_EXECUTE;
			hr = VsiUnitTest::SendKeys(&wKey, 1, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			VsiUnitTest::WaitForEnabled(2000, wndViewer);

			wndViewer.PostMessage(WM_COMMAND, MAKEWPARAM(ID_FILTER_INFO_ONLY, 0));

			Sleep(2000);

			// Retrieves count
			CWindow wndMsgList = wndViewer.GetDlgItem(IDC_MESSAGE_LIST);
			Assert::IsTrue(wndMsgList != NULL, "Message List missing.");
			int iCount = ListView_GetItemCount(wndMsgList);
			Assert::AreEqual(3, iCount, "Count not matched.");

			// Close viewer
			wndViewer.PostMessage(WM_CLOSE);

			VsiUnitTest::WaitForNoWindow(2000, VS_WND_CLASS_LOG_VIEWER, NULL);
		}


		/// <summary>
		/// <step>Open test file MessageLog001.txt</step>
		/// <step>Sets Find Edit box to "Test1"</step>
		/// <step>Retrieves message count in Message Log View window</step>
		/// </summary>
		/// <results>7 message in Message Log View window</results>
		[TestMethod]
		void TestFilterEdit()
		{
			// Shows Message Log Viewer
			LaunchLogViewer();

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, NULL);
			Assert::IsTrue(wndViewer != NULL, "Viewer not found.");

			SetForegroundWindow(wndViewer);

			// Open file
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strDataPath;
			strDataPath.Format(L"%s\\TestMessageLog001.txt", pszTestDir);

			wndViewer.PostMessage(WM_COMMAND, MAKEWPARAM(ID_LOG_OPEN, 0));

			VsiUnitTest::WaitForDisabled(2000, wndViewer);

			Assert::IsTrue(!wndViewer.IsWindowEnabled());

			Sleep(500);

			// Send path
			HRESULT hr = VsiUnitTest::SendKeys(strDataPath, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			WORD wKey = VK_EXECUTE;
			hr = VsiUnitTest::SendKeys(&wKey, 1, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			VsiUnitTest::WaitForEnabled(2000, wndViewer);

			Sleep(2000);

			// Set filter string
			wndViewer.SendMessage(WM_COMMAND, IDC_FILTER);

			Sleep(50);

			hr = VsiUnitTest::SendKeys(L"Test1", 500);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			Sleep(500);

			// Retrieves count
			CWindow wndMsgList = wndViewer.GetDlgItem(IDC_MESSAGE_LIST);
			Assert::IsTrue(wndMsgList != NULL, "Message List missing.");
			int iCount = ListView_GetItemCount(wndMsgList);
			Assert::AreEqual(6, iCount, "Count not matched.");

			// Close viewer
			wndViewer.PostMessage(WM_CLOSE);

			VsiUnitTest::WaitForNoWindow(2000, VS_WND_CLASS_LOG_VIEWER, NULL);
		}

		/// <summary>
		/// <step>Open test file MessageLog001.txt</step>
		/// <step>Sets focus to 3rd row</step>
		/// <step>Send F3 and Shift-F3</step>
		/// </summary>
		/// <results>message focus in Message Log View window go to next row and previous row</results>
		[TestMethod]
		void TestFindNextPrev()
		{
			// Shows Message Log Viewer
			LaunchLogViewer();

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, NULL);
			Assert::IsTrue(wndViewer != NULL, "Viewer not found.");

			SetForegroundWindow(wndViewer);

			// Open file
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strDataPath;
			strDataPath.Format(L"%s\\TestMessageLog001.txt", pszTestDir);

			wndViewer.PostMessage(WM_COMMAND, MAKEWPARAM(ID_LOG_OPEN, 0));

			VsiUnitTest::WaitForDisabled(2000, wndViewer);

			Assert::IsTrue(!wndViewer.IsWindowEnabled());

			Sleep(500);

			// Send path
			HRESULT hr = VsiUnitTest::SendKeys(strDataPath, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			WORD wKey = VK_EXECUTE;
			hr = VsiUnitTest::SendKeys(&wKey, 1, 2000);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			VsiUnitTest::WaitForEnabled(2000, wndViewer);

			// Send find string
			wndViewer.SendMessage(WM_COMMAND, IDC_FIND);

			Sleep(50);

			hr = VsiUnitTest::SendKeys(L"8", 500);
			Assert::IsTrue(SUCCEEDED(hr), "SendKeys failed.");

			Sleep(500);

			// Go to previous message
			wndViewer.SendMessage(WM_COMMAND, MAKEWPARAM(ID_EDIT_FIND_PREV, 0));
			Sleep(50);

			wndViewer.SendMessage(WM_COMMAND, MAKEWPARAM(ID_EDIT_FIND_PREV, 0));
			Sleep(500);

			// Retrieves selection
			CWindow wndMsgList = wndViewer.GetDlgItem(IDC_MESSAGE_LIST);
			Assert::IsTrue(wndMsgList != NULL, "Message List missing.");

			int iSel = ListView_GetNextItem(wndMsgList, -1, LVIS_SELECTED | LVIS_FOCUSED);
			Assert::AreEqual(8, iSel, "Selection not matched.");

			Sleep(5000);
			// Go to next message
			wndViewer.SendMessage(WM_COMMAND,  MAKEWPARAM(ID_EDIT_FIND_NEXT, 0));

			Sleep(1000);

			// Retrieves selection
			iSel = ListView_GetNextItem(wndMsgList, -1, LVIS_SELECTED | LVIS_FOCUSED);
			Assert::AreEqual(7, iSel, "Selection not matched.");

			// Close viewer
			wndViewer.PostMessage(WM_CLOSE);

			VsiUnitTest::WaitForNoWindow(2000, VS_WND_CLASS_LOG_VIEWER, NULL);
		}



	private:
		void InitMsgLogViewer(VSI_LOG_FLAG dwMode)
		{
			HRESULT hr;

			if (dwMode & VSI_LOG_FLAG_REVIEW)
			{
				hr = m_pMsgLog->Initialize(dwMode, NULL, NULL);
				Assert::AreEqual(S_OK, hr, "Initialize failed.");
			}
			else
			{
				String^ strDir = testContextInstance->TestDeploymentDir;
				pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

				hr = m_pMsgLog->Initialize(
					VSI_LOG_FLAG_NO_INIT_MSG | VSI_LOG_FLAG_NO_ERROR_BOX, pszTestDir, L"VsiLogTest");
				Assert::AreEqual(S_OK, hr, "Initialize failed.");
			}

			hr = m_pMsgLogViewer->Initialize(
				m_pMsgLog, NULL,
				(dwMode & VSI_LOG_FLAG_REVIEW) ? VSI_LOGV_FLAG_REVIEW : VSI_LOGV_FLAG_NONE);
			Assert::AreEqual(S_OK, hr, "Initialize failed.");
		}

		void UninitMsgLogViewer()
		{
			HRESULT hr;

			hr = m_pMsgLogViewer->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");

			hr = m_pMsgLog->Uninitialize();
			Assert::AreEqual(S_OK, hr, "Uninitialize failed.");
		}

		void LaunchLogViewer()
		{
			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			CString strViewer;
			strViewer.Format(L"%s\\VsiMessageLogViewer.exe", pszTestDir);

			// Shows Message Log Viewer
			STARTUPINFO si = { sizeof(STARTUPINFO) };
			PROCESS_INFORMATION pi = { 0 };
			BOOL bRet = CreateProcess(
				strViewer,
				NULL,
				NULL, NULL,
				FALSE,
				NORMAL_PRIORITY_CLASS,
				NULL,
				NULL,
				&si, &pi);
			Assert::IsTrue(FALSE != bRet, "CreateProcess failed.");

			CloseHandle(pi.hThread);
			m_hLogProcess = pi.hProcess;

			HRESULT hr = VsiUnitTest::WaitForWindow(2000, VS_WND_CLASS_LOG_VIEWER, NULL);
			Assert::AreEqual(S_OK, hr, "WaitForWindow failed.");

			WaitForInputIdle(m_hLogProcess, 1000);

			CWindow wndViewer = FindWindow(VS_WND_CLASS_LOG_VIEWER, NULL);
			Assert::IsTrue(wndViewer != NULL, "Viewer not found.");
	
			SetForegroundWindow(wndViewer);
		}
	};
}
