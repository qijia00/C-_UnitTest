/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorExport.cpp
**
**	Description:
**		Implementation of CVsiOperatorExport
**
*******************************************************************************/

#include "stdafx.h"
#include <Shellapi.h>
#include <ATLComTime.h>
#include <VsiWTL.h>
#include <VsiGlobalDef.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiServiceProvider.h>
#include <VsiServiceKey.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiImportExportXml.h>
#include "VsiOperatorExport.h"


CVsiOperatorExport::CVsiOperatorExport()
{
}

CVsiOperatorExport::~CVsiOperatorExport()
{
	_ASSERT(m_pApp == NULL);
	_ASSERT(m_pCmdMgr == NULL);
	_ASSERT(m_pPdm == NULL);
}

STDMETHODIMP CVsiOperatorExport::Activate(IVsiApp *pApp, HWND hwndParent)
{
	HRESULT hr = S_OK;

	try
	{
		m_pApp = pApp;
		VSI_CHECK_INTERFACE(m_pApp, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Get Command Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CMD_MGR,
			__uuidof(IVsiCommandManager),
			(IUnknown**)&m_pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get PDM
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HWND hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorExport::Deactivate()
{
	HRESULT hr = S_FALSE;

	if (IsWindow())
	{
		BOOL bRet = DestroyWindow();
		_ASSERT(bRet);

		hr = S_OK;
	}

	m_pCmdMgr.Release();
	m_pPdm.Release();
	m_pApp.Release();

	return hr;
}

STDMETHODIMP CVsiOperatorExport::GetWindow(HWND *phWnd)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(phWnd, VSI_LOG_ERROR, L"GetWindow failed");

		*phWnd = m_hWnd;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorExport::GetAccelerator(HACCEL *phAccel)
{
	UNREFERENCED_PARAMETER(phAccel);
	return E_NOTIMPL;
}

STDMETHODIMP CVsiOperatorExport::GetIsBusy(
	DWORD dwStateCurrent,
	DWORD dwState,
	BOOL bTryRelax,
	BOOL *pbBusy)
{
	UNREFERENCED_PARAMETER(dwStateCurrent);
	UNREFERENCED_PARAMETER(dwState);
	UNREFERENCED_PARAMETER(bTryRelax);

	*pbBusy = !GetParent().IsWindowEnabled();

	return S_OK;
}

STDMETHODIMP CVsiOperatorExport::PreTranslateMessage(
	MSG *pMsg, BOOL *pbHandled)
{
	*pbHandled = IsDialogMessage(pMsg);

	return S_OK;
}

STDMETHODIMP CVsiOperatorExport::Initialize(IUnknown *pPropBag)
{
	HRESULT hr = S_OK;

	try
	{
		CComQIPtr<IVsiPropertyBag> pProperties(pPropBag);
		VSI_CHECK_INTERFACE(pProperties, VSI_LOG_ERROR, NULL);

		CComVariant vSize;
		hr = pProperties->Read(L"size", &vSize);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		m_iRequiredSize = V_I8(&vSize);

		CComVariant vPath;
		hr = pProperties->Read(L"cabPath", &vPath);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);		

		for (int i = 0; i < 9999; ++i)
		{
			CComVariant vFileName;
			hr = pProperties->ReadId(i, &vFileName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr)
			{
				CString strFilePath(V_BSTR(&vPath));
				strFilePath += V_BSTR(&vFileName);

				m_vecExportFiles.push_back(strFilePath);
			}
			else
			{
				// No more
				break;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

LRESULT CVsiOperatorExport::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		OSVERSIONINFO osv = {0};
		osv.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
		GetVersionEx((LPOSVERSIONINFO) &osv);
		if (VER_PLATFORM_WIN32_NT == osv.dwPlatformId && osv.dwMajorVersion > 5)  // Vista or later
		{
			// Composited window (double-buffering)
			ModifyStyleEx(0, WS_EX_COMPOSITED);
		}

		// Set font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
		VsiThemeRecurSetFont(m_hWnd, hFont);

		// Layout
		CComPtr<IVsiWindowLayout> pLayout;
		HRESULT hr = pLayout.CoCreateInstance(__uuidof(VsiWindowLayout));
		if (SUCCEEDED(hr))
		{
			hr = pLayout->Initialize(m_hWnd, VSI_WL_FLAG_AUTO_RELEASE);
			if (SUCCEEDED(hr))
			{
				pLayout->AddControl(0, IDC_IE_NEW_FOLDER, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDOK, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDCANCEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_FOLDER_TREE, VSI_WL_SIZE_XY);
				pLayout->AddControl(0, IDC_IE_SELECTED, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_REQUIRED_LABEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_REQUIRED, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_AVAILABLE_LABEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_AVAILABLE, VSI_WL_MOVE_X);

				pLayout.Detach();
			}
		}
		
		SetWindowText(CString(MAKEINTRESOURCE(IDS_OPER_EXPORT_OPERATOR)));

		hr = m_pFExplorer.CoCreateInstance(__uuidof(VsiFolderExplorer));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		UpdateSizeUI();
		UpdateUI();

		// Refresh tree
		SetDlgItemText(IDC_IE_PATH, CString(MAKEINTRESOURCE(IDS_LOADING)));
		PostMessage(WM_VSI_REFRESH_TREE);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorExport::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	if (m_pFExplorer != NULL)
	{
		m_pFExplorer->Uninitialize();
		m_pFExplorer.Release();
	}

	return 0;
}

LRESULT CVsiOperatorExport::OnRefreshTree(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_FE_FLAG dwFlags(VSI_FE_FLAG_NONE);
		CComPtr<IFolderFilter> pFolderFilter;

		DWORD dwState = VsiGetBitfieldValue<DWORD>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_STATE, m_pPdm);
		if (0 < (VSI_SYS_STATE_SYSTEM & dwState) &&
			0 == (VSI_SYS_STATE_ENGINEER_MODE & dwState))
		{
			// For Sierra system, shows removable and network drive
			dwFlags = VSI_FE_FLAG_SHOW_NETWORK | VSI_FE_FLAG_SHOW_REMOVEABLE;

			// Install custom filter		
			hr = pFolderFilter.CoCreateInstance(__uuidof(VsiFolderFilter));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else
		{
			// For workstation and eng mode, shows Desktop
			dwFlags = VSI_FE_FLAG_SHOW_DESKTOP | VSI_FE_FLAG_EXPLORER;
		}

		hr = m_pFExplorer->Initialize(
			GetDlgItem(IDC_IE_FOLDER_TREE),
			dwFlags | VSI_FE_FLAG_IGNORE_LAST_FOLDER,
			pFolderFilter);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Set selection
		PostMessage(WM_VSI_SET_SELECTION);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorExport::OnSetSelection(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	CWaitCursor wait;

	try
	{
		// Set default selected folder
		CString strPath(VsiGetRangeValue<CString>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_PATH_EXPORT_BACKUP,
			m_pPdm));

		hr = m_pFExplorer->SetSelectedNamespacePath(strPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK != hr)
		{
			UpdateUI();
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorExport::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Export Operator - OK");

	HRESULT hr = S_OK;

	try
	{
		if (m_iAvailableSize < m_iRequiredSize)
		{
			ShowNotEnoughSpaceMessage();
			return 0;
		}

		// Overwrite test
		if (CheckForOverwrite() > 0)
		{
			int iRet = ShowOverwriteMessage();
			if (IDYES != iRet)
			{
				return 0;
			}
		}

		CComHeapPtr<OLECHAR> pszPathName;
		hr = m_pFExplorer->GetSelectedNamespacePath(&pszPathName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		VsiSetRangeValue<LPCWSTR>(
			pszPathName,
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_PATH_EXPORT_BACKUP,
			m_pPdm);

		WCHAR szPathTo[MAX_PATH];
		if (GetSelectedPath(szPathTo, _countof(szPathTo)))
		{
			int iMaxSize = m_vecExportFiles.size() * MAX_PATH;
			std::unique_ptr<WCHAR[]> pszFrom(new WCHAR[iMaxSize]);

			LPWSTR pszWrite = pszFrom.get();
			CString strTest;
			for (auto iter = m_vecExportFiles.cbegin(); iter != m_vecExportFiles.cend(); ++iter)
			{
				const CString& strFilePath(*iter);

				wcscpy_s(pszWrite, iMaxSize, strFilePath);

				int iLength = strFilePath.GetLength() + 1;

				pszWrite += iLength;  // After \0
				iMaxSize -= iLength;
			}

			*pszWrite = 0;  // Double NULL terminated
			*(szPathTo + lstrlen(szPathTo) + 1) = 0;  // Double NULL terminated

			SHFILEOPSTRUCT sfos = {
				NULL,
				FO_MOVE,
				pszFrom.get(),
				szPathTo,
				FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT,
				FALSE,
				NULL,
				NULL
			};

			int iRet = SHFileOperation(&sfos);
			if (0 != iRet)
			{
				VsiMessageBox(
					*this,
					L"An error has occurred while exporting the user settings. "
					L"Make sure you have sufficient privileges to write to the destination folder.",
					L"Export User Failed",
					MB_OK | MB_ICONERROR);
			}
			else
			{
				// Close
				m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_OPERATOR_EXPORT, NULL);
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorExport::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Export Operator - Cancel");

	HRESULT hr = S_OK;

	try
	{
		// Delete temp files - not critical
		for (auto iter = m_vecExportFiles.cbegin(); iter != m_vecExportFiles.cend(); ++iter)
		{
			const CString& strFilePath(*iter);

			BOOL bRet = DeleteFile(strFilePath);
			if (!bRet)
			{
				VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"Failed to delete temp file - %s", strFilePath.GetString()));
			}
		}

		m_vecExportFiles.clear();
	}
	VSI_CATCH(hr);

	m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_OPERATOR_EXPORT, NULL);

	return 0;
}

BOOL CVsiOperatorExport::GetAvailablePath(LPWSTR pszBuffer, int iBufferLength)
{
	return GetSelectedPath(pszBuffer, iBufferLength);
}

LRESULT CVsiOperatorExport::OnTreeSelChanged(int, LPNMHDR, BOOL&)
{
	if (m_pFExplorer != NULL)
	{
		UpdateAvailableSpace();
		UpdateSizeUI();
		UpdateUI();
	}

	return 0;
}

void CVsiOperatorExport::UpdateUI()
{
	CVsiImportExportUIHelper<CVsiOperatorExport>::UpdateUI();

	WCHAR szPath[VSI_MAX_PATH];
	GetDlgItem(IDOK).EnableWindow(GetSelectedPath(szPath, _countof(szPath)));
}

int CVsiOperatorExport::CheckForOverwrite()
{
	HRESULT hr = S_OK;
	BOOL bOverwrite = FALSE;

	try
	{
		WCHAR szPath[VSI_MAX_PATH];
		if (GetSelectedPath(szPath, _countof(szPath)))
		{
			CString strTest;
			for (auto iter = m_vecExportFiles.cbegin(); iter != m_vecExportFiles.cend(); ++iter)
			{
				const CString& strFilePath(*iter);

				LPCWSTR pszName = PathFindFileName(strFilePath);
				strTest.Format(L"%s\\%s", szPath, pszName);

				if (VsiGetFileExists(strTest))
				{
					bOverwrite = TRUE;
					break;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return bOverwrite ? 1 : 0;
}

