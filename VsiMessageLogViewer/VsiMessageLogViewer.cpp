/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiMessageLogViewer.cpp
**
**	Description:
**		Implementation of WinMain
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiGdiPlus.h>
#include <VsiGlobalDef.h>
#include <VsiNewWinApi.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include "resource.h"


class CVsiMessageLogViewer : public CAtlExeModuleT<CVsiMessageLogViewer>
{
public:
	LPCWSTR m_pszCmdLine;

private:
	// Message Log viewer
	CComPtr<IVsiMessageLogViewer> m_pMessageLogViewer;

	CVsiGdiPlus m_GdiPlus;

public:
	CVsiMessageLogViewer();
	~CVsiMessageLogViewer();

	int _WinMain(LPCWSTR pszCmdLine, int iShowCmd);

// CAtlExeModuleT
	HRESULT PreMessageLoop(int iShowCmd);
	HRESULT PostMessageLoop();
	void RunMessageLoop() throw();
};

CVsiMessageLogViewer _AtlModule;


CVsiMessageLogViewer::CVsiMessageLogViewer() :
	m_pszCmdLine(NULL)
{
}

CVsiMessageLogViewer::~CVsiMessageLogViewer()
{
}

int CVsiMessageLogViewer::_WinMain(LPCWSTR pszCmdLine, int iShowCmd)
{
	m_pszCmdLine = pszCmdLine;

	return WinMain(iShowCmd);
}

/// <summary>
/// Overrides CAtlExeModuleT::PreMessageLoop()
/// </summary>
HRESULT CVsiMessageLogViewer::PreMessageLoop(int iShowCmd)
{
	UNREFERENCED_PARAMETER(iShowCmd);

	HRESULT hr = S_OK;
	CString strErrorMsg;

	try
	{
		CString strOpenFile;

		// Setups the debug heap manager
		_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);

		// Do not display the critical-error-handler message box
		SetErrorMode(
			SEM_FAILCRITICALERRORS |
			SEM_NOGPFAULTERRORBOX |
			SEM_NOOPENFILEERRORBOX);

		// SetProcessUserModeExceptionPolicy (new in Window 7 SP1)
		// Otherwise exceptions that are thrown from an application that runs in a 64-bit version of Windows are ignored
		VSI_SET_PROCESS_USERMODE_EXCEPTION_ON();

		hr = OleInitialize(NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating Message Log.\n");

		// GDI+
		hr = m_GdiPlus.Initialize();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure initializing GDI+");

		// Create Notify Log
		hr = VsiInitializeLog();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating Message Log.\n");

		// Initialize commutl
		hr = VsiCommUtlInitialize();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"VsiCommUtlInitialize failed");

		// Initialize commctl
		hr = VsiCommCtlInitialize(0);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"VsiCommCtlInitialize failed");

		VsiThemeAttachAll(_AtlBaseModule.GetModuleInstance(), GetCurrentThreadId());

		hr = g_pMessageLog->Initialize(VSI_LOG_FLAG_REVIEW, NULL, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating Message Log.\n");

		// Open file
		if (NULL != m_pszCmdLine && lstrlen(m_pszCmdLine) > 0)
		{
			CComVariant vtError;
			CComQIPtr<IVsiPersistMessageLog> pPersist(g_pMessageLog);
			HRESULT hrLoad = pPersist->LoadArchive(m_pszCmdLine, &vtError);
			if (FAILED(hrLoad))
			{
				if (VT_BSTR == V_VT(&vtError))
				{
					VsiMessageBox(NULL, V_BSTR(&vtError), CString(MAKEINTRESOURCE(IDS_APP_NAME)), MB_OK | MB_ICONERROR);
				}
				else
				{
					CString strMsg;
					strMsg.Format(L"Failed to open - %s", m_pszCmdLine);
					VsiMessageBox(NULL, strMsg, CString(MAKEINTRESOURCE(IDS_APP_NAME)), MB_OK | MB_ICONERROR);
				}

				// Quit application
				return S_FALSE;  // Jump to PostMessageLoop
			}
			else
			{
				strOpenFile = wcsrchr(m_pszCmdLine, L'\\') + 1;
			}
		}

		// Initialize viewer
		hr = m_pMessageLogViewer.CoCreateInstance(__uuidof(VsiMessageLogViewer));
		if (FAILED(hr))
			VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure initializing Message Log viewer");

		hr = m_pMessageLogViewer->Initialize(
			g_pMessageLog,
			NULL,
			VSI_LOGV_FLAG_REVIEW | VSI_LOGV_FLAG_POST_QUIT_MSG);
		if (FAILED(hr))
			VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure initializing Message Log viewer");

		CString strTitle;
		if (strOpenFile.IsEmpty())
			strTitle.LoadString(IDS_APP_NAME);
		else
			strTitle.Format(L"%s - %s", (LPCWSTR)strOpenFile, CString(MAKEINTRESOURCE(IDS_APP_NAME)));

		hr = m_pMessageLogViewer->Activate((LPCWSTR)strTitle);
		if (FAILED(hr))
			VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure initializing Message Log viewer");
	}
	VSI_CATCH_(err)
	{
		hr = S_FALSE;  // Jump to PostMessageLoop

		// Message
		WCHAR szSupport[500];
		VsiGetTechSupportInfo(szSupport, _countof(szSupport));

		CString strMsg(err.m_strMsg);
		strMsg += szSupport;

		VsiMessageBox(
			NULL,
			strMsg,
			CString(MAKEINTRESOURCE(IDS_APP_NAME)),
			MB_OK | MB_SYSTEMMODAL | MB_ICONERROR);
	}

	return hr;
}

/// <summary>
/// Overrides CAtlExeModuleT::PostMessageLoop()
/// </summary>
HRESULT CVsiMessageLogViewer::PostMessageLoop()
{
	// Detach viewer
	if (m_pMessageLogViewer != NULL)
	{
		m_pMessageLogViewer->Deactivate();
		m_pMessageLogViewer->Uninitialize();
		m_pMessageLogViewer.Release();
	}

	VsiThemeDetachAll(GetCurrentThreadId());

	// Uninitialize commctl
	VsiCommCtlUninitialize();

	// Uninitialize commutl
	VsiCommUtlUninitialize();

	if (g_pMessageLog != NULL)
	{
		g_pMessageLog->Uninitialize();

		VsiUninitializeLog();
	}

	// GDI+
	m_GdiPlus.Uninitialize();

	OleUninitialize();

	return S_OK;
}

/// <summary>
/// Overrides CAtlExeModuleT::RunMessageLoop()
/// </summary>
void CVsiMessageLogViewer::RunMessageLoop() throw()
{
	MSG msg;
	while (GetMessage(&msg, 0, 0, 0) > 0)
	{
		if (S_OK != m_pMessageLogViewer->TranslateAccelerator(&msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
}

extern "C" int WINAPI _tWinMain(
	HINSTANCE hInstance, HINSTANCE hPrevInstance,
	LPTSTR lpCmdLine, int iShowCmd)
{
	UNREFERENCED_PARAMETER(hInstance);
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(iShowCmd);

	return _AtlModule._WinMain(lpCmdLine, iShowCmd);
}

