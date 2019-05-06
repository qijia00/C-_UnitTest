/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorAdmin.cpp
**
**	Description:
**		Implementation of CVsiOperatorAdmin
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiGlobalDef.h>
#include <VsiSaxUtils.h>
#include <VsiRes.h>
#include <VsiServiceProvider.h>
#include <VsiServiceKey.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiAppControllerModule.h>
#include <VsiPdmModule.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiMeasurementModule.h>
#include <VsiAnalRsltXmlTags.h>
#include "VsiOperatorXml.h"
#include "VsiOperatorProp.h"
#include "VsiOperatorAdmin.h"


typedef struct tagVSI_COLUMN_INFO
{
	DWORD dwIndex;
	DWORD dwWidth;
	LPCWSTR pszLabel;
} VSI_COLUMN_INFO;


// Operators
enum
{
	VSI_OPERATOR_COL_NAME = 0,
	VSI_OPERATOR_COL_TYPE,
	VSI_OPERATOR_COL_GROUP,
};

static const VSI_COLUMN_INFO g_ColumnInfoOperatorSecured[] =
{
	{ VSI_OPERATOR_COL_NAME, 180, L"Name" },
	{ VSI_OPERATOR_COL_TYPE, 120, L"Type" },
	{ VSI_OPERATOR_COL_GROUP, 120, L"Group" },
};

static const VSI_COLUMN_INFO g_ColumnInfoOperator[] =
{
	{ VSI_OPERATOR_COL_NAME, 280, L"Name" },
	{ VSI_OPERATOR_COL_TYPE, 120, L"Type" },
};

// Usage Log
enum
{
	VSI_USAGE_COL_START = 0,
	VSI_USAGE_COL_END,
	VSI_USAGE_COL_TOTAL,
	VSI_USAGE_COL_OPERATOR,
};

VSI_COLUMN_INFO g_ColumnInfoUsage[] =
{
	{ VSI_USAGE_COL_START, 200, L"Start Date" },
	{ VSI_USAGE_COL_END, 200, L"End Date" },
	{ VSI_USAGE_COL_TOTAL, 100, L"Total Time" },
	{ VSI_USAGE_COL_OPERATOR, 200, L"User" },
};


#define VSI_LABEL_ADMINISTRATOR L"Administrator"
#define VSI_LABEL_STANDARD L"Standard"
#define VSI_LABEL_ENABLE L"Active"
#define VSI_LABEL_DISABLE L"Inactive"
#define VSI_MAX_LINES_DISPLAY_OPERATOR_NAME 9

CVsiOperatorAdmin::CVsiOperatorAdmin() :
	m_pOperatorMgr(NULL),
	m_iSortColumn(VSI_OPERATOR_COL_NAME),
	m_bSortDescending(FALSE)
{
	m_dwTitleID = IDS_PREF_OPERATOR;
}

CVsiOperatorAdmin::~CVsiOperatorAdmin()
{
}

STDMETHODIMP CVsiOperatorAdmin::SetObjects(ULONG nObjects, IUnknown **ppUnk)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_VERIFY(0 != nObjects, VSI_LOG_ERROR, NULL);

		hr = IPropertyPageImpl<CVsiOperatorAdmin>::SetObjects(nObjects, ppUnk);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		m_pApp = m_ppUnk[0];
		VSI_CHECK_INTERFACE(m_pApp, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);

		// Get PDM
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Gets Operator Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_OPERATOR_MGR,
			__uuidof(IVsiOperatorManager),
			(IUnknown**)&m_pOperatorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"CVsiOperatorAdmin::Activate failed");

		// Gets Command Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CMD_MGR,
			__uuidof(IVsiCommandManager),
			(IUnknown**)&m_pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Gets session
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_SESSION_MGR,
			__uuidof(IVsiSession),
			(IUnknown**)&m_pSession);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// User Management Mode 
		m_bSecureMode = VsiGetBooleanValue<bool>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
			VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
		m_bSecureModeToChange = m_bSecureMode;
		VsiSetBooleanValue<bool>(
			m_bSecureModeToChange,
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_SYS_SECURE_MODE_TO_CHANGE, m_pPdm);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorAdmin::Apply()
{
	HRESULT hr = S_OK;
	BOOL bErrorBox = TRUE;
	BOOL bContinue = TRUE;

	try
	{
		// Save User Management Mode state
		VsiSetBooleanValue<bool>(
			m_bSecureModeToChange,
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_SYS_SECURE_MODE_TO_CHANGE, m_pPdm);

		// Save usage log state
		VsiSetBooleanValue<bool>(
			BST_CHECKED == IsDlgButtonChecked(IDC_OPERATOR_TURNON_USAGE_CHECK),
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_SYS_USAGE_LOG, m_pPdm);

		CComQIPtr<IVsiPersistOperatorManager> pOpmgrPersist(m_pOperatorMgr);
		VSI_CHECK_INTERFACE(pOpmgrPersist, VSI_LOG_ERROR, NULL);

		if (S_OK == pOpmgrPersist->IsDirty())
		{
			// Start saving
			VSI_LOG_MSG(VSI_LOG_INFO | VSI_LOG_DISABLE_BOX, L"Error box disabled");
			bErrorBox = FALSE;
			
			// Saves operator list
			hr = pOpmgrPersist->SaveOperator(NULL);
			
			if (FAILED(hr))
			{
				bContinue = FALSE;

				WCHAR szSupport[500];
				VsiGetTechSupportInfo(szSupport, _countof(szSupport));

				CString strMsg(L"An error occurred while saving user settings.\n");
				strMsg += szSupport;

				VsiMessageBox(
					GetTopLevelParent(),
					strMsg,
					L"User Administration Error",
					MB_OK | MB_ICONERROR);
			}
			else
			{
				SetDirty(FALSE);
			}

		}

		if (bContinue)
		{
			// Create backup when turning off user management mode
			if (m_bSecureMode && !m_bSecureModeToChange)
			{
				CComPtr<IVsiAppBackup> pAppBackup;
				hr = pAppBackup.CoCreateInstance(__uuidof(VsiAppBackup));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CString strPathBackup(VsiGetRangeValue<CString>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_SYS_PATH_BACKUP, m_pPdm));

				CString strDataPath;
				BOOL bResult = VsiGetApplicationDataDirectory(
					AtlFindStringResourceInstance(IDS_VSI_PRODUCT_NAME),
					strDataPath.GetBufferSetLength(MAX_PATH));
				strDataPath.ReleaseBuffer();
				VSI_VERIFY(bResult, VSI_LOG_ERROR, L"Failure getting application data directory");

				hr = pAppBackup->Initialize(strPathBackup, strDataPath);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				VSI_SYSTEM_BACKUP_INFO info;
				ZeroMemory(&info, sizeof(info));

				VSI_APP_MODE_FLAGS dwAppMode(VSI_APP_MODE_NONE);
				hr = m_pApp->GetAppMode(&dwAppMode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	
				if (VSI_APP_MODE_WORKSTATION & dwAppMode)
				{
					info.type = VSI_BACKUP_TYPE_WORKSTATION;
				}
				else
				{
					info.type = VSI_BACKUP_TYPE_SYSTEM;
				}

				CComHeapPtr<OLECHAR> pOperatorName;

				BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
					VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
				if (bSecureMode)
				{
					CComPtr<IVsiOperator> pOperator;
					hr = m_pOperatorMgr->GetCurrentOperator(&pOperator);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pOperator->GetName(&pOperatorName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					info.pszBy = pOperatorName;
				}

				GetSystemTime(&info.stDate);

				CString strVersion(VsiGetRangeValue<CString>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL, VSI_PARAMETER_SYS_VERSION_SOFTWARE, m_pPdm));

				info.pszVersionSoftware = strVersion;

				hr = pAppBackup->SystemBackup(&info);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			// Close all open series when turning off user management mode
			if (m_bSecureMode && !m_bSecureModeToChange)  
			{
				DWORD dwCount(0);
				hr = m_pOperatorMgr->GetOperatorCount(
					VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_STANDARD,
					&dwCount);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (0 < (int)dwCount)
				{
					CComPtr<IVsiEnumOperator> pOperators;
					hr = m_pOperatorMgr->GetOperators(
						VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_STANDARD,
						&pOperators);

					if (SUCCEEDED(hr))
					{
						CComPtr<IVsiOperator> pOperator;
						while (pOperators->Next(1, &pOperator, NULL) == S_OK)
						{
							CComHeapPtr<OLECHAR> pszId;
							hr = pOperator->GetId(&pszId);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							CComHeapPtr<OLECHAR> pszSeriesNs;
							hr = m_pSession->CheckActiveSeries(FALSE, pszId, &pszSeriesNs);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
							if (NULL != pszSeriesNs)
							{
								hr = m_pSession->CheckActiveSeries(TRUE, pszId, NULL);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							}

							pOperator.Release();
						}
					}
				}
			}

			// Process custom packages and presets when enable User Management Mode
			if (!m_bSecureMode && m_bSecureModeToChange)  
			{
				VSI_LOG_MSG(VSI_LOG_INFO | VSI_LOG_DISABLE_BOX, L"Error box disabled");
				bErrorBox = FALSE;

				hr = ProcessCustomPackagesPresets();
				if (FAILED(hr))
				{
					// Message box and Rollback changes
					bContinue = FALSE;

					VsiSetBooleanValue<bool>(
						m_bSecureMode,
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
						VSI_PARAMETER_SYS_SECURE_MODE_TO_CHANGE, m_pPdm);

					// Save usage log state
					VsiSetBooleanValue<bool>(
						false,
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
						VSI_PARAMETER_SYS_USAGE_LOG, m_pPdm);

					WCHAR szSupport[500];
					VsiGetTechSupportInfo(szSupport, _countof(szSupport));

					CString strMsg(L"An error occurred while processing Measurement Packages and Presets.\n");
					strMsg += szSupport;

					VsiMessageBox(
						GetTopLevelParent(),
						strMsg,
						L"Enable User Management Error",
						MB_OK | MB_ICONERROR);
				}
				else
				{
					SetDirty(FALSE);
				}

				hr = SaveOperatorsSettings();
				if (FAILED(hr))
				{
					// Message box and Rollback changes
					bContinue = FALSE;

					VsiSetBooleanValue<bool>(
						m_bSecureMode,
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
						VSI_PARAMETER_SYS_SECURE_MODE_TO_CHANGE, m_pPdm);

					// Save usage log state
					VsiSetBooleanValue<bool>(
						false,
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
						VSI_PARAMETER_SYS_USAGE_LOG, m_pPdm);

					WCHAR szSupport[500];
					VsiGetTechSupportInfo(szSupport, _countof(szSupport));

					CString strMsg(L"An error occurred while saving operators' settings.\n");
					strMsg += szSupport;

					VsiMessageBox(
						GetTopLevelParent(),
						strMsg,
						L"Enable User Management Error",
						MB_OK | MB_ICONERROR);
				}
				else
				{
					SetDirty(FALSE);
				}			
			}
		}
	}
	VSI_CATCH(hr);

	if (!bErrorBox)
	{
		VSI_LOG_MSG(VSI_LOG_INFO | VSI_LOG_ENABLE_BOX, L"Error box enabled");
	}

	return (bContinue? S_OK : S_FALSE);
}

LRESULT CVsiOperatorAdmin::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		// Set font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
		VsiThemeRecurSetFont(m_hWnd, hFont);

		// Layout
		CComPtr<IVsiWindowLayout> pLayout;
		HRESULT hr = pLayout.CoCreateInstance(__uuidof(VsiWindowLayout));
		if (SUCCEEDED(hr))
		{
			hr = pLayout->Initialize(m_hWnd, VSI_WL_FLAG_MINIMUM_SIZE | VSI_WL_FLAG_AUTO_RELEASE);
			if (SUCCEEDED(hr))
			{
				pLayout->AddControl(0, IDOK, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDCANCEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_OPERATOR_LIST, VSI_WL_SIZE_Y);
				pLayout->AddControl(0, IDC_OPERATOR_USAGE_GROUP, VSI_WL_SIZE_X);
				pLayout->AddControl(0, IDC_OPERATOR_USAGE, VSI_WL_SIZE_XY);
				pLayout->AddControl(0, IDC_OPERATOR_USAGE_EXPORT, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_OPERATOR_USAGE_CLEAR, VSI_WL_MOVE_X);

				pLayout.Detach();
			}
		}
		
		SetWindowText(L"User Administration");

		m_wndOperatorList = GetDlgItem(IDC_OPERATOR_LIST);

		ListView_SetExtendedListViewStyleEx(
			m_wndOperatorList,
			LVS_EX_LABELTIP | LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER,
			LVS_EX_LABELTIP | LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);

		UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);
		m_wndOperatorList.SendMessage(uiThemeCmd, VSI_LV_CMD_NEW_THEME);

		// Remove image list
		HIMAGELIST hilOld = ListView_SetImageList(m_wndOperatorList, NULL, LVSIL_SMALL);
		if (NULL != hilOld)
		{
			ImageList_Destroy(hilOld);
		}
		hilOld = ListView_SetImageList(m_wndOperatorList, NULL, LVSIL_STATE);
		if (NULL != hilOld)
		{
			ImageList_Destroy(hilOld);
		}

		InitOperatorListColumns();

		// Usage Log
		//if (VsiLicenseUtils::CheckFeatureAvailable(VSI_LICENSE_FEATURE_USAGE_LOG))
		//{
			//GetDlgItem(IDC_OPERATOR_USAGE_GROUP).ShowWindow(SW_SHOWNA);

			//BOOL bUsageLog = VsiGetBooleanValue<BOOL>(
			//	VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
			//	VSI_PARAMETER_SYS_USAGE_LOG, m_pPdm);

			//if (bUsageLog && m_bSecureMode)
			//{
			//	CheckDlgButton(IDC_OPERATOR_TURNON_USAGE_CHECK, BST_CHECKED);
			//	GetDlgItem(IDC_OPERATOR_USAGE).EnableWindow(TRUE);
			//}

			//GetDlgItem(IDC_OPERATOR_TURNON_USAGE_CHECK).ShowWindow(SW_SHOWNA);

			//GetDlgItem(IDC_OPERATOR_USAGE).ShowWindow(SW_SHOWNA);

			//GetDlgItem(IDC_OPERATOR_USAGE_EXPORT).ShowWindow(SW_SHOWNA);

			//GetDlgItem(IDC_OPERATOR_USAGE_CLEAR).ShowWindow(SW_SHOWNA);
		//}

		if (m_bSecureMode)
		{
			CheckDlgButton(IDC_OPERATOR_TURNON_SECURE_CHECK, BST_CHECKED);

			CComPtr<IVsiOperator> pOperatorCurrent;
			hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorCurrent);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (NULL != pOperatorCurrent)
			{
				VSI_OPERATOR_TYPE dwType;
				hr = pOperatorCurrent->GetType(&dwType);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				BOOL bAdmin = 0 != (VSI_OPERATOR_TYPE_ADMIN & dwType);			
				if (bAdmin)
				{
					GetDlgItem(IDC_OPERATOR_TURNON_SECURE_CHECK).EnableWindow(TRUE);
					GetDlgItem(IDC_OPERATOR_ADD).EnableWindow(TRUE);
					GetDlgItem(IDC_OPERATOR_TURNON_USAGE_CHECK).EnableWindow(TRUE);

					GetDlgItem(IDC_OPERATOR_TURNON_USAGE_CHECK).EnableWindow(TRUE);
				}
				else
				{
					GetDlgItem(IDC_OPERATOR_TURNON_SECURE_CHECK).EnableWindow(FALSE);

					CString strAdminOnly(MAKEINTRESOURCE(IDS_PREF_OPERATOR_ADMIN_ONLY_GROUP));

					CString strLabel;
					GetDlgItem(IDC_OPERATOR_SECURE_GROUP).GetWindowText(strLabel);
					strLabel += strAdminOnly;
					GetDlgItem(IDC_OPERATOR_SECURE_GROUP).SetWindowText(strLabel);

					GetDlgItem(IDC_OPERATOR_USAGE_GROUP).GetWindowText(strLabel);
					strLabel += strAdminOnly;
					GetDlgItem(IDC_OPERATOR_USAGE_GROUP).SetWindowText(strLabel);
				}
			}

			GetDlgItem(IDC_OPERATOR_EXPORT).ShowWindow(SW_SHOWNA);
			GetDlgItem(IDC_OPERATOR_IMPORT).EnableWindow(TRUE);
			GetDlgItem(IDC_OPERATOR_IMPORT).ShowWindow(SW_SHOWNA);
		}
		else
		{
			GetDlgItem(IDC_OPERATOR_TURNON_SECURE_CHECK).EnableWindow(TRUE);
			GetDlgItem(IDC_OPERATOR_ADD).EnableWindow(TRUE);
		}

		m_wndUsageList = GetDlgItem(IDC_OPERATOR_USAGE);

		ListView_SetExtendedListViewStyleEx(
			m_wndUsageList,
			LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER,
			LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);

		m_wndUsageList.SendMessage(uiThemeCmd, VSI_LV_CMD_NEW_THEME);

		// Remove image list
		hilOld = ListView_SetImageList(m_wndOperatorList, NULL, LVSIL_SMALL);
		if (NULL != hilOld)
		{
			ImageList_Destroy(hilOld);
		}
		hilOld = ListView_SetImageList(m_wndOperatorList, NULL, LVSIL_STATE);
		if (NULL != hilOld)
		{
			ImageList_Destroy(hilOld);
		}

		// Usage log
		LV_COLUMN lvc = { LVCF_FMT | LVCF_WIDTH | LVCF_TEXT };
		lvc.fmt = LVCFMT_LEFT;

		for (int i = 0; i < _countof(g_ColumnInfoUsage); ++i)
		{
			lvc.cx = g_ColumnInfoUsage[i].dwWidth;
			lvc.pszText = const_cast<LPWSTR>(g_ColumnInfoUsage[i].pszLabel);
			ListView_InsertColumn(m_wndUsageList, g_ColumnInfoUsage[i].dwIndex, &lvc);
		}

		// Add operators
		RefreshOperators();

		// Add usage log
		RefreshUsageLog();

		// Scrolling
		RECT rc;
		GetClientRect(&rc);
		CVsiScrollImpl::SetScrollSize(rc.right, rc.bottom);
		CVsiScrollImpl::GetSystemSettings();

		// Sets focus
		PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)m_wndOperatorList, TRUE);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorAdmin::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	return 0;
}

LRESULT CVsiOperatorAdmin::OnRefreshOperatorList(UINT, WPARAM wp, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	LPCWSTR pszOperatorId(NULL);
	CComHeapPtr<OLECHAR> pszId;

	try
	{
		if (NULL != wp)
		{
			LPCWSTR pszOperator = reinterpret_cast<LPCWSTR>(wp);

			CComPtr<IVsiOperator> pOperator;
			hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, pszOperator, &pOperator);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (NULL != pOperator)
			{
				hr = pOperator->GetId(&pszId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				pszOperatorId = pszId;
			}
		}
	}
	VSI_CATCH(hr);

	RefreshOperators(pszOperatorId);

	return 0;
}

LRESULT CVsiOperatorAdmin::OnBnClickedSecure(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		//if (VsiLicenseUtils::CheckFeatureAvailable(VSI_LICENSE_FEATURE_SECURE_MODE))
		//{
			if (!m_bSecureModeToChange)  // Enable secured mode
			{
				// Need admin account exist
				DWORD dwAdmins(0);
				hr = m_pOperatorMgr->GetOperatorCount(
					VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_PASSWORD, &dwAdmins);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (0 == dwAdmins)
				{
					VsiMessageBox(
						*this,
						L"Administrator account with password is mandatory to enable this feature.",
						L"Enable User Management Mode",
						MB_OK | MB_ICONWARNING);

					return 0;
				}

				// Check open series, must close open series first
				CComHeapPtr<OLECHAR> pszSeriesNs;
				hr = m_pSession->CheckActiveSeries(FALSE, NULL, &pszSeriesNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
				if (pszSeriesNs != NULL)
				{
					VsiMessageBox(
						*this,
						L"Open series must be closed first to enable this feature.",
						L"Enable User Management Mode",
						MB_OK | MB_ICONWARNING);

					return 0;
				}

				// No matter Admin or not, re-type (Admin) password
				CComPtr<IVsiOperatorLogon> pLogon;
				hr = pLogon.CoCreateInstance(__uuidof(VsiOperatorLogon));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pLogon->Activate(
					m_pApp,
					GetParent(),
					L"Verification",
					L"Log in as administrator to activate User Management Mode.",
					VSI_OPERATOR_TYPE_ADMIN | VSI_OPERATOR_TYPE_PASSWORD_MANDATRORY,
					TRUE,
					NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"User Management Mode - On");

					m_bSecureModeToChange = true;

					SetDirty(TRUE);
					CheckDlgButton(IDC_OPERATOR_TURNON_SECURE_CHECK, BST_CHECKED);

					// Pop up warning with list of operators have no password
					DWORD dwCount(0);
					hr = m_pOperatorMgr->GetOperatorCount(
						VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_STANDARD |
							VSI_OPERATOR_FILTER_TYPE_NO_PASSWORD,
						&dwCount);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (0 < (int)dwCount)
					{
						CComPtr<IVsiEnumOperator> pOperators;
						hr = m_pOperatorMgr->GetOperators(
							VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_STANDARD |
								VSI_OPERATOR_FILTER_TYPE_NO_PASSWORD,
							&pOperators);

						if (SUCCEEDED(hr))
						{
							CString strMsg(L"The following users need password to log in:\n");
							bool bFirst(true);
							CComPtr<IVsiOperator> pOperator;
							while (pOperators->Next(1, &pOperator, NULL) == S_OK)
							{
								CComHeapPtr<OLECHAR> pszName;
								hr = pOperator->GetName(&pszName);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								if (bFirst)
								{
									bFirst = false;
								}
								else
								{
									if (VSI_MAX_LINES_DISPLAY_OPERATOR_NAME > (int)dwCount)
									{
										strMsg += L"\n";
									}
									else
									{
										strMsg += L";  ";
									}
								}

								strMsg += pszName;

								pOperator.Release();
							}

							VsiMessageBox(*this, strMsg, L"Enable User Management Mode", MB_OK | MB_ICONWARNING);
						}
					}

					// Allow Usage Log
					GetDlgItem(IDC_OPERATOR_TURNON_USAGE_CHECK).EnableWindow(TRUE);

					// Show new columns
					InitOperatorListColumns();
					RefreshOperators();
				}
			}
			else  // Disable secured mode, buttons already grayed out when standard operator logs in, so must be Admin
			{
				VSI_APP_MODE_FLAGS dwAppMode(VSI_APP_MODE_NONE);
				hr = m_pApp->GetAppMode(&dwAppMode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	
				int iRet = -1;

				if (VSI_APP_MODE_WORKSTATION & dwAppMode)
				{
					// Warn message for workstation
					iRet = VsiMessageBox(
						*this,
						L"After User Management Mode is disabled:\r\n\r\n"
						L"-All study and series data will be available to all users\r\n"
						L"-A backup of the current system settings will be created\r\n"
						L"-Custom Measurement Packages created in User Management Mode will be saved to the User Profiles, but will not be available when User Management Mode is off\r\n\r\n"
						L"Are you sure you want to proceed?\r\n",
						L"Disable User Management Mode",
						MB_YESNO | MB_ICONEXCLAMATION);
				}
				else
				{
					// Warn message for system
					iRet = VsiMessageBox(
						*this,
						L"After User Management Mode is disabled:\r\n\r\n"
						L"-Open series will be closed\r\n"
						L"-All study and series data will be available to all users\r\n"
						L"-A backup of the current system settings will be created\r\n"
						L"-Custom Measurement Packages and Presets created in User Management Mode will be saved to the User Profiles, but will not be available when User Management Mode is off\r\n\r\n"
						L"Are you sure you want to proceed?\r\n",
						L"Disable User Management Mode",
						MB_YESNO | MB_ICONEXCLAMATION);
				}

				if (IDYES == iRet)
				{
					// Admin password is required to disable User Management Mode
					CComPtr<IVsiOperatorLogon> pLogon;
					hr = pLogon.CoCreateInstance(__uuidof(VsiOperatorLogon));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pLogon->Activate(
						m_pApp,
						GetParent(),
						L"Verification",
						L"Only administrator can disable User Management Mode.",
						VSI_OPERATOR_TYPE_ADMIN | VSI_OPERATOR_TYPE_PASSWORD_MANDATRORY,
						FALSE,
						NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"User Management Mode - Off");

						m_bSecureModeToChange = false;

						SetDirty(TRUE);

						// Uncheck and Disable Usage Log
						CheckDlgButton(IDC_OPERATOR_TURNON_USAGE_CHECK, BST_UNCHECKED);
						GetDlgItem(IDC_OPERATOR_TURNON_USAGE_CHECK).EnableWindow(FALSE);
						GetDlgItem(IDC_OPERATOR_USAGE).EnableWindow(FALSE);

						CheckDlgButton(IDC_OPERATOR_TURNON_SECURE_CHECK, BST_UNCHECKED);

						// Show new columns
						InitOperatorListColumns();
						RefreshOperators();
					}
				}
			}
		//}
		//else
		//{
		//	m_pCmdMgr->InvokeCommand(ID_CMD_FEATURE_NOT_AVAILABLE, NULL);
		//}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorAdmin::OnBnClickedOperatorAdd(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Operator - Add");

	try
	{
		CVsiOperatorProp prop(CVsiOperatorProp::VSI_OPERATION_ADD, m_pApp, NULL, m_bSecureModeToChange, false);
		INT_PTR iRet = prop.DoModal();

		if (IDOK == iRet)
		{
			CComPtr<IVsiOperator> pOperator;
			hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"LoadOperator - failure creating operator");

			pOperator->SetName(prop.m_strName);
			pOperator->SetPassword(prop.m_strPassword, FALSE);
			pOperator->SetType(prop.m_dwType);
			if (m_bSecureModeToChange)
			{
				pOperator->SetDefaultStudyAccess(prop.m_dwDefaultStudyAccess);
				pOperator->SetGroup(prop.m_strGroup);
			}

			WCHAR szId[128];
			VsiGetGuid(szId, _countof(szId));
			pOperator->SetId(szId);

			m_pOperatorMgr->AddOperator(pOperator, prop.m_strIdCopyFrom);

			// Update UI
			int i = AddOperator(pOperator);
			if (i >= 0)
			{
				ListView_SetItemState(
					m_wndOperatorList, i,
					LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);

				ListView_EnsureVisible(m_wndOperatorList, i, FALSE);
			}

			SortOperatorList();

			SetDirty(TRUE);
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorAdmin::OnBnClickedOperatorDelete(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Operator - Delete");

	try
	{
		int i = ListView_GetNextItem(m_wndOperatorList, -1, LVIS_SELECTED);
		if (i >= 0)
		{
			CString strId;

			if (GetSelectedOperator(strId))
			{
				CComPtr<IVsiOperator> pOperator;

				hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_ID, strId, &pOperator);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComHeapPtr<OLECHAR> pszName;
				hr = pOperator->GetName(&pszName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				BOOL bDelete = FALSE;

				DWORD dwAdminWithPasswordCount(0);
				hr = m_pOperatorMgr->GetOperatorCount(
					VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_PASSWORD,
					&dwAdminWithPasswordCount);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				BOOL bLoginRequired = 0 < dwAdminWithPasswordCount;

				CString strTitle(L"Delete Confirmation");
				if (bLoginRequired)
				{
					CString strMsg(L"Only users with administrator rights can delete user account.\r\n"
						L"Enter the password to confirm deletion.");

					CComPtr<IVsiOperatorLogon> pLogon;
					hr = pLogon.CoCreateInstance(__uuidof(VsiOperatorLogon));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pLogon->Activate(
						m_pApp,
						GetParent(),
						strTitle,
						strMsg,
						VSI_OPERATOR_TYPE_ADMIN | VSI_OPERATOR_TYPE_PASSWORD_MANDATRORY,
						FALSE,
						NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					
					bDelete = (S_OK == hr);
				}
				else
				{
					CString strMsg;
					strMsg.Format(L"Are you sure you want to delete user \'%s\'?", pszName);
					
					int iRet = VsiMessageBox(*this, strMsg, strTitle, MB_YESNO | MB_ICONEXCLAMATION);
					bDelete = (IDYES == iRet);
				}

				if (bDelete)
				{
					m_pOperatorMgr->RemoveOperator(strId);

					ListView_DeleteItem(m_wndOperatorList, i);

					WORD pwId[] =
					{
						IDC_OPERATOR_DELETE,
						IDC_OPERATOR_MODIFY
					};

					for (int j = 0; j < _countof(pwId); ++j)
					{
						GetDlgItem(pwId[j]).EnableWindow(FALSE);
					}

					int iCount = ListView_GetItemCount(m_wndOperatorList);
					if (0 < iCount)
					{
						ListView_SetItemState(
							m_wndOperatorList, (i >= iCount) ? (i - 1) : i,
							LVIS_SELECTED | LVIS_FOCUSED,
							LVIS_SELECTED | LVIS_FOCUSED);
					}

					m_pOperatorMgr->OperatorModified();

					SetDirty(TRUE);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorAdmin::OnBnClickedOperatorModify(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Operator - Modify");

	HRESULT hr = S_OK;

	try
	{
		CString strId;

		if (GetSelectedOperator(strId))
		{
			CComPtr<IVsiOperator> pOperator;
			hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_ID, strId, &pOperator);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CVsiOperatorProp prop(
				CVsiOperatorProp::VSI_OPERATION_MODIFY,
				m_pApp,
				pOperator,
				m_bSecureModeToChange,
				IsSelectedOperatorLogin());

			INT_PTR iRet = prop.DoModal();
			if (IDOK == iRet)
			{
				pOperator->SetName(prop.m_strName);
				if (prop.m_bUpdatePassword)
				{
					pOperator->SetPassword(prop.m_strPassword, prop.m_strPassword.IsEmpty());
				}
				pOperator->SetType(prop.m_dwType);
				if (m_bSecureModeToChange)
				{
					pOperator->SetDefaultStudyAccess(prop.m_dwDefaultStudyAccess);
					pOperator->SetGroup(prop.m_strGroup);
				}

				m_pOperatorMgr->OperatorModified();

				// Update UI
				int i = ListView_GetNextItem(m_wndOperatorList, -1, LVIS_SELECTED);
				if (i >= 0)
				{
					UpdateOperator(i, pOperator);
				}

				SortOperatorList();

				SetDirty(TRUE);
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorAdmin::OnBnClickedOperatorExport(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	CVsiCabFile cab;

	try
	{
		VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Operator - Export");

		CString strId;

		if (GetSelectedOperator(strId))
		{
			CComPtr<IVsiOperator> pOperator;

			hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_ID, strId, &pOperator);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszName;
			hr = pOperator->GetName(&pszName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszDataPath;
			hr = m_pOperatorMgr->GetOperatorDataPath(
				VSI_OPERATOR_PROP_NAME, pszName, &pszDataPath);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Get a temp path
			WCHAR szTempPath[MAX_PATH];
			int iRet = GetTempPath(_countof(szTempPath), szTempPath);
			VSI_WIN32_VERIFY(0 != iRet, VSI_LOG_ERROR, NULL);

			CString strCabFilename;
			strCabFilename.Format(L"%s.operator.vbak", pszName.operator wchar_t *());

			// Temp file name does not have to be unique
			CString pszFilePath(szTempPath);			
			pszFilePath += strCabFilename;

			// Create cab file
			{
				CWaitCursor wait;
											
				hr = cab.Create(pszFilePath);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating cab file");

				hr = cab.AddFolder(pszDataPath, L"Data");
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating cab file");

				// Get a backup info path
				WCHAR szBackupInfoPath[MAX_PATH];
				int iRet = GetTempPath(_countof(szBackupInfoPath), szBackupInfoPath);
				VSI_WIN32_VERIFY(0 != iRet, VSI_LOG_ERROR, NULL);

				PathAppend(szBackupInfoPath, L"OperatorBackup.xml");

				CComPtr<IMXWriter> pWriter;
				hr = VsiCreateXmlSaxWriter(&pWriter);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pWriter->put_indent(VARIANT_TRUE);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pWriter->put_encoding(CComBSTR(L"utf-8"));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CVsiMXFileWriterStream out;
				hr = out.Open(CComBSTR(szBackupInfoPath));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pWriter->put_output(CComVariant(static_cast<IStream*>(&out)));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComQIPtr<ISAXContentHandler> pch(pWriter);
				VSI_CHECK_INTERFACE(pch, VSI_LOG_ERROR, NULL);

				CComPtr<IMXAttributes> pmxa;
				hr = VsiCreateXmlSaxAttributes(&pmxa);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComQIPtr<ISAXAttributes> pa(pmxa);

				hr = pch->startDocument();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				WCHAR szNull[] = L"";
				CComBSTR bstrNull(L"");

				CComBSTR bstrAtt(VSI_OPERATOR_BACKUP_ATT_VERSION);
				CComBSTR bstrValue(VSI_OPERATOR_BACKUP_VALUE_VERSION);

				hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, bstrValue);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				{
					bstrAtt = VSI_OPERATOR_BACKUP_ATT_TYPE;
					bstrValue = VSI_OPERATOR_BACKUP_VALUE_TYPE_OPERATOR;

					hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, bstrValue);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				{
					bstrAtt = VSI_OPERATOR_BACKUP_ATT_SOFTWARE_VERSION;
					bstrValue = VsiGetRangeValue<CString>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL, VSI_PARAMETER_SYS_VERSION_SOFTWARE, m_pPdm);

					hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, bstrValue);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				{
					bstrAtt = VSI_OPERATOR_BACKUP_ATT_DATE;

					SYSTEMTIME st;
					GetSystemTime(&st);

					CComVariant vDate;
					hr = VsiXmlSysTimeToVariantXml(&st, &vDate);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					bstrValue = V_BSTR(&vDate);

					hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, bstrValue);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				{
					CComHeapPtr<OLECHAR> pOperatorName;

					BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
						VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
					if (bSecureMode)
					{
						CComPtr<IVsiOperator> pOperator;
						hr = m_pOperatorMgr->GetCurrentOperator(&pOperator);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						hr = pOperator->GetName(&pOperatorName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						bstrAtt = VSI_OPERATOR_BACKUP_ATT_BY;
						bstrValue = pOperatorName;

						hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, bstrValue);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}

				hr = pch->startElement(szNull, 0, szNull, 0,
					VSI_OPERATOR_BACKUP_ELM_BACKUP, VSI_STATIC_STRING_COUNT(VSI_OPERATOR_BACKUP_ELM_BACKUP), pa.p);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Save operator settings
				CComQIPtr<IVsiPersistOperator> pPersistOperator(pOperator);
				hr = pPersistOperator->SaveSax(pch);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pch->endElement(szNull, 0, szNull, 0,
					VSI_OPERATOR_BACKUP_ELM_BACKUP, VSI_STATIC_STRING_COUNT(VSI_OPERATOR_BACKUP_ELM_BACKUP));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pch->endDocument();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				out.Close();

				hr = cab.AddFile(szBackupInfoPath, L"OperatorBackup.xml");
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure adding backup info file");

				cab.Close();
			}

			LARGE_INTEGER size = { 0 };
			hr = VsiGetFileSize(pszFilePath, &size);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IVsiPropertyBag> pProperties;
			hr = pProperties.CoCreateInstance(__uuidof(VsiPropertyBag));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vSize(size.QuadPart);
			hr = pProperties->Write(L"size", &vSize);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vPath(szTempPath);				
			hr = pProperties->Write(L"cabPath", &vPath);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vFile(strCabFilename);	
			hr = pProperties->WriteId(0, &vFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant v(pProperties.p);
			m_pCmdMgr->InvokeCommand(ID_CMD_OPEN_OPERATOR_EXPORT, &v);
		}
	}
	VSI_CATCH(hr);

	cab.Close();

	return 0;
}

LRESULT CVsiOperatorAdmin::OnBnClickedOperatorImport(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Operator - Import");

		CComPtr<IVsiPropertyBag> pProperties;
		hr = pProperties.CoCreateInstance(__uuidof(VsiPropertyBag));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vHwnd((ULONGLONG)m_hWnd);
		hr = pProperties->Write(L"refreshHWND", &vHwnd);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vMsg(WM_VSI_REFRESH_OPERATOR_LIST);
		hr = pProperties->Write(L"refreshMsg", &vMsg);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant v(pProperties.p);
		m_pCmdMgr->InvokeCommand(ID_CMD_OPEN_OPERATOR_IMPORT, &v);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorAdmin::OnBnClickedUsage(WORD, WORD, HWND, BOOL&)
{
	//if (VsiLicenseUtils::CheckFeatureAvailable(VSI_LICENSE_FEATURE_USAGE_LOG))
	//{
		BOOL bUsageLogOn = BST_CHECKED != IsDlgButtonChecked(IDC_OPERATOR_TURNON_USAGE_CHECK);
		CheckDlgButton(IDC_OPERATOR_TURNON_USAGE_CHECK, bUsageLogOn ? BST_CHECKED : BST_UNCHECKED);

		GetDlgItem(IDC_OPERATOR_USAGE).EnableWindow(bUsageLogOn);

		SetDirty(TRUE);
	//}
	//else
	//{
	//	m_pCmdMgr->InvokeCommand(ID_CMD_FEATURE_NOT_AVAILABLE, NULL);
	//}

	return 0;
}

LRESULT CVsiOperatorAdmin::OnBnClickedUsageExport(WORD, WORD, HWND, BOOL&)
{
	return 0;
}

LRESULT CVsiOperatorAdmin::OnBnClickedUsageClear(WORD, WORD, HWND, BOOL&)
{
	return 0;
}

void CVsiOperatorAdmin::InitOperatorListColumns()
{
	const VSI_COLUMN_INFO *pColumnInfo(NULL);
	int iColumns(0);

	if (m_bSecureModeToChange)
	{
		pColumnInfo = g_ColumnInfoOperatorSecured;
		iColumns = _countof(g_ColumnInfoOperatorSecured);
	}
	else
	{
		pColumnInfo = g_ColumnInfoOperator;
		iColumns = _countof(g_ColumnInfoOperator);
	}

	// Delete old
	int iCount = Header_GetItemCount(ListView_GetHeader(m_wndOperatorList));

	for (int i = 0; i < iCount; ++i)
	{
		ListView_DeleteColumn(m_wndOperatorList, 0);
	}

	// Create columns
	LV_COLUMN lvc = { LVCF_FMT | LVCF_WIDTH | LVCF_TEXT };
	lvc.fmt = LVCFMT_LEFT;
	for (int i = 0; i < iColumns; ++i)
	{
		lvc.cx = pColumnInfo[i].dwWidth;
		lvc.pszText = const_cast<LPWSTR>(pColumnInfo[i].pszLabel);
		ListView_InsertColumn(m_wndOperatorList, pColumnInfo[i].dwIndex, &lvc);
	}
}

void CVsiOperatorAdmin::RefreshOperators(LPCWSTR pszSelectId)
{
	HRESULT hr = S_OK;

	try
	{
		// UI
		ListView_DeleteAllItems(m_wndOperatorList);

		CComPtr<IVsiEnumOperator> pOperators;
		hr = m_pOperatorMgr->GetOperators(
			VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN,
			&pOperators);
		if (SUCCEEDED(hr))
		{
			CComPtr<IVsiOperator> pOperator;
			while (pOperators->Next(1, &pOperator, NULL) == S_OK)
			{
				AddOperator(pOperator);

				pOperator.Release();
			}
		}

		// Sort operators
		SortOperatorList();

		// Select operator
		if (NULL != pszSelectId)
		{
			SelectOperator(pszSelectId);
		}
		else
		{
			if (m_bSecureMode)
			{
				CComPtr<IVsiOperator> pOperatorCurrent;
				hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorCurrent);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (NULL != pOperatorCurrent)
				{
					// Selects operator
					CComHeapPtr<OLECHAR> pszId;
					hr = pOperatorCurrent->GetId(&pszId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					SelectOperator(pszId);
				}
				else
				{
					// Selects 1st operator
					ListView_SetItemState(
						m_wndOperatorList, 0,
						LVIS_SELECTED | LVIS_FOCUSED,
						LVIS_SELECTED | LVIS_FOCUSED);
				}
			}
			else
			{
				// Selects 1st operator
				ListView_SetItemState(
					m_wndOperatorList, 0,
					LVIS_SELECTED | LVIS_FOCUSED,
					LVIS_SELECTED | LVIS_FOCUSED);
			}
		}
	}
	VSI_CATCH(hr);
}

void CVsiOperatorAdmin::SelectOperator(LPCWSTR pszId)
{
	LVITEM lvi = { LVIF_PARAM };

	int iCount = ListView_GetItemCount(m_wndOperatorList);
	for (lvi.iItem = 0; lvi.iItem < iCount; ++lvi.iItem)
	{
		if (ListView_GetItem(m_wndOperatorList, &lvi))
		{
			if (0 == wcscmp(reinterpret_cast<OLECHAR*>(lvi.lParam), pszId))
			{
				lvi.mask = LVIF_STATE;
				lvi.state = LVIS_SELECTED | LVIS_FOCUSED;
				lvi.stateMask = LVIS_SELECTED | LVIS_FOCUSED;

				ListView_SetItem(m_wndOperatorList, &lvi);

				break;
			}
		}
	}
}

VSI_OPERATOR_TYPE CVsiOperatorAdmin::GetLoginOperatorType()
{
	VSI_OPERATOR_TYPE dwType = VSI_OPERATOR_TYPE_NONE;

	HRESULT hr(S_OK);

	try
	{
		CComPtr<IVsiOperator> pOperatorCurrent;
		hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorCurrent);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (NULL != pOperatorCurrent)
		{
			hr = pOperatorCurrent->GetType(&dwType);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return dwType;
}

bool CVsiOperatorAdmin::IsSelectedOperatorLogin()
{
	bool bRet = false;
	HRESULT hr(S_OK);

	try
	{
		CString strId;
		if (GetSelectedOperator(strId))
		{
			CComPtr<IVsiOperator> pOperatorCurrent;
			hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorCurrent);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (NULL != pOperatorCurrent)
			{
				CComHeapPtr<OLECHAR> pszIdCurrent;
				hr = pOperatorCurrent->GetId(&pszIdCurrent);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (strId == pszIdCurrent)
				{
					bRet = true;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return bRet;
}

BOOL CVsiOperatorAdmin::GetSelectedOperator(CString &strId)
{
	BOOL bRet(FALSE);

	int iIndex = ListView_GetNextItem(m_wndOperatorList, -1, LVIS_SELECTED);
	if (iIndex != -1)
	{
		LVITEM lvi = { LVIF_PARAM };
		lvi.iItem = iIndex;
		ListView_GetItem(m_wndOperatorList, &lvi);

		strId = reinterpret_cast<OLECHAR*>(lvi.lParam);

		bRet = TRUE;
	}

	return bRet;
}

LRESULT CVsiOperatorAdmin::OnListCustomDraw(int, LPNMHDR pnmh, BOOL &bHandled)
{
	if (m_bSecureModeToChange)
	{
		LPNMLVCUSTOMDRAW plvcd = (LPNMLVCUSTOMDRAW)pnmh;

		switch (plvcd->nmcd.dwDrawStage)
		{
		default:
			{
				bHandled = FALSE;
			}
			break;

		case (CDDS_SUBITEM | CDDS_ITEMPREPAINT):
			{
				// Change font color
				CRect rc;
				ListView_GetSubItemRect(
					plvcd->nmcd.hdr.hwndFrom,
					(int)plvcd->nmcd.dwItemSpec,
					plvcd->iSubItem,
					LVIR_BOUNDS,
					&rc);
				rc.right = rc.left + plvcd->nmcd.rc.right - plvcd->nmcd.rc.left;

				WCHAR szText[500];
				LVITEM lvi = { LVIF_IMAGE | LVIF_STATE | LVIF_TEXT | LVIF_INDENT };
				lvi.iItem = (int)plvcd->nmcd.dwItemSpec;
				lvi.iSubItem = plvcd->iSubItem;
				lvi.stateMask = LVIS_STATEIMAGEMASK;
				lvi.pszText = szText;
				lvi.cchTextMax = _countof(szText);

				ListView_GetItem(plvcd->nmcd.hdr.hwndFrom, &lvi);

				if (0 == plvcd->iSubItem)
				{
					DWORD dwExStyle = ListView_GetExtendedListViewStyle(plvcd->nmcd.hdr.hwndFrom);
					if (LVS_EX_CHECKBOXES & dwExStyle)
					{
						rc.left += 4;
						int iTop = rc.top + (rc.Height() - 16) / 2;

						int iState = ((lvi.state & LVIS_STATEIMAGEMASK) >> 12) - 1;

						ImageList_Draw(
							ListView_GetImageList(plvcd->nmcd.hdr.hwndFrom, LVSIL_STATE),
							iState,
							plvcd->nmcd.hdc,
							rc.left,
							iTop,
							ILD_TRANSPARENT);

						rc.left += 12;
						rc.right += 16;
					}

					rc.right += 6;
				}

				if (0 != *szText)
				{
					rc.left += 6;
					rc.right -= 2;

					HRESULT hr = S_OK;

					try
					{
						COLORREF rgbText;

						if (m_bSecureModeToChange)
						{
							rgbText = ListView_GetTextColor(plvcd->nmcd.hdr.hwndFrom);

							CComPtr<IVsiOperator> pOperator;
							hr = m_pOperatorMgr->GetOperator(
								VSI_OPERATOR_PROP_ID,
								reinterpret_cast<LPCOLESTR>(plvcd->nmcd.lItemlParam), 
								&pOperator);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

							hr = pOperator->HasPassword();
							if (S_FALSE == hr)
							{
								COLORREF rgbBk = ListView_GetTextBkColor(plvcd->nmcd.hdr.hwndFrom);

								SetTextColor(plvcd->nmcd.hdc, VsiThemeGetDisabledColor(rgbText, rgbBk));
							}
						}

						LVCOLUMN lvc = { LVCF_FMT };
						ListView_GetColumn(
							plvcd->nmcd.hdr.hwndFrom,
							plvcd->iSubItem,
							&lvc);

						UINT uFormat = DT_END_ELLIPSIS | DT_NOPREFIX | DT_SINGLELINE | DT_VCENTER;
						if (lvc.fmt & LVCFMT_CENTER)
						{
							uFormat |= DT_CENTER;
						}
						else if (lvc.fmt & LVCFMT_RIGHT)
						{
							uFormat |= DT_RIGHT;
						}

						SetBkMode(plvcd->nmcd.hdc, TRANSPARENT);
						DrawText(plvcd->nmcd.hdc, szText, -1, &rc, uFormat);

						if (m_bSecureModeToChange)
						{
							SetTextColor(plvcd->nmcd.hdc, rgbText);
						}
					}
					VSI_CATCH(hr);
				}
			}
			return CDRF_SKIPDEFAULT;
		}
	}
	else
	{
		bHandled = FALSE;
	}

	return CDRF_DODEFAULT;
}

LRESULT CVsiOperatorAdmin::OnListItemActivate(int, LPNMHDR, BOOL&)
{
	if (GetDlgItem(IDC_OPERATOR_MODIFY).IsWindowEnabled())
	{
		BOOL b;
		OnBnClickedOperatorModify(BN_CLICKED, IDC_OPERATOR_MODIFY, 0, b);
	}

	return 0;
}

LRESULT CVsiOperatorAdmin::OnListDeleteItem(int, LPNMHDR pnmh, BOOL&)
{
	NMLISTVIEW *pNMListView = (NMLISTVIEW*)pnmh;

	LVITEM lvi = { LVIF_PARAM };
	lvi.iItem = pNMListView->iItem;
	ListView_GetItem(pnmh->hwndFrom, &lvi);

	if (NULL != lvi.lParam)
	{
		CComHeapPtr<OLECHAR> pszId;
		pszId.Attach(reinterpret_cast<OLECHAR*>(lvi.lParam));
		pszId.Free();
	}

	return 0;
}

LRESULT CVsiOperatorAdmin::OnListItemChanged(int, LPNMHDR pnmh, BOOL&)
{
	LPNMLISTVIEW plv = (LPNMLISTVIEW)pnmh;

	if (LVIF_STATE & plv->uChanged)
	{
		if (plv->uNewState != plv->uOldState)
		{
			BOOL bRemove(FALSE);
			BOOL bModify(FALSE);
			BOOL bExport(FALSE);

			if (plv->iItem >= 0)
			{
				if (LVIS_SELECTED & plv->uNewState)
				{
					VSI_OPERATOR_TYPE type = VSI_OPERATOR_TYPE_BASE_MASK & GetLoginOperatorType();

					if (m_bSecureModeToChange)
					{
						switch (type)
						{
						case VSI_OPERATOR_TYPE_ADMIN:
							// Can modify everyone
							bModify = TRUE;
							bExport = TRUE;
							if (!IsSelectedOperatorLogin())
							{
								// Cannot remove self
								bRemove = TRUE;
							}
							break;
						case VSI_OPERATOR_TYPE_STANDARD:
							if (IsSelectedOperatorLogin())
							{
								// Can only modify self
								bModify = TRUE;
								bExport = TRUE;
							}
							break;
						default:
							_ASSERT(0);	
						}
					}
					else
					{
						if (m_bSecureMode && IsSelectedOperatorLogin())  
						{
							// In transition state, cannot remove self
							bRemove = FALSE;
							bModify = TRUE;
							bExport = TRUE;
						}
						else
						{
							// Not User Management Mode, enable both
							bRemove = TRUE;
							bModify = TRUE;
							bExport = TRUE;
						}
					}
				}
			}

			GetDlgItem(IDC_OPERATOR_DELETE).EnableWindow(bRemove);
			GetDlgItem(IDC_OPERATOR_MODIFY).EnableWindow(bModify);
			GetDlgItem(IDC_OPERATOR_EXPORT).EnableWindow(bExport);
		}
	}

	return 0;
}

LRESULT CVsiOperatorAdmin::OnListColumnClick(int, LPNMHDR pnmh, BOOL&)
{
	NMLISTVIEW *pnmlv = (NMLISTVIEW*)pnmh;

	if (pnmlv->iSubItem == m_iSortColumn)
	{
		m_bSortDescending = !m_bSortDescending;
	}
	else
	{
		// Remove old sort indicator
		SetOperatorSortIndicator(m_iSortColumn, 0);

		m_iSortColumn = pnmlv->iSubItem;
		m_bSortDescending = TRUE;
	}

	SortOperatorList();

	return 0;
}

int CVsiOperatorAdmin::AddOperator(IVsiOperator *pOperator)
{
	HRESULT hr(S_OK);
	LVITEM lvi = { LVIF_TEXT | LVIF_PARAM };
	lvi.iItem = 0;
	try
	{
		// Name
		CComHeapPtr<OLECHAR> pszId;
		hr = pOperator->GetId(&pszId);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		lvi.lParam = (LPARAM)pszId.Detach();

		CComHeapPtr<OLECHAR> pszName;
		hr = pOperator->GetName(&pszName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		lvi.pszText = pszName;
		lvi.iSubItem = VSI_OPERATOR_COL_NAME;
		lvi.iItem = ListView_InsertItem(m_wndOperatorList, &lvi);

		lvi.mask = LVIF_TEXT;

		// Type
		VSI_OPERATOR_TYPE dwType;
		hr = pOperator->GetType(&dwType);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		lvi.iSubItem = VSI_OPERATOR_COL_TYPE;
		lvi.pszText = (VSI_OPERATOR_TYPE_ADMIN == dwType) ?
			VSI_LABEL_ADMINISTRATOR : VSI_LABEL_STANDARD;
		ListView_SetItem(m_wndOperatorList, &lvi);

		if (m_bSecureModeToChange)
		{
			CComHeapPtr<OLECHAR> pszGroup;
			hr = pOperator->GetGroup(&pszGroup);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			lvi.pszText = pszGroup;
			lvi.iSubItem = VSI_OPERATOR_COL_GROUP;
			ListView_SetItem(m_wndOperatorList, &lvi);
		}
	}
	VSI_CATCH(hr);

	return lvi.iItem;
}

void CVsiOperatorAdmin::UpdateOperator(int iIndex, IVsiOperator *pOperator)
{
	HRESULT hr(S_OK);

	try
	{
		LVITEM lvi = { LVIF_TEXT };
		lvi.iItem = iIndex;

		// Name
		CComHeapPtr<OLECHAR> pszName;
		hr = pOperator->GetName(&pszName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		lvi.pszText = pszName;
		lvi.iSubItem = VSI_OPERATOR_COL_NAME;
		ListView_SetItem(m_wndOperatorList, &lvi);

		// Type
		VSI_OPERATOR_TYPE dwType;
		hr = pOperator->GetType(&dwType);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		lvi.iSubItem = VSI_OPERATOR_COL_TYPE;
		lvi.pszText = (VSI_OPERATOR_TYPE_ADMIN == dwType) ?
			VSI_LABEL_ADMINISTRATOR : VSI_LABEL_STANDARD;
		ListView_SetItem(m_wndOperatorList, &lvi);

		if (m_bSecureModeToChange)
		{
			CComHeapPtr<OLECHAR> pszGroup;
			hr = pOperator->GetGroup(&pszGroup);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			lvi.pszText = pszGroup;
			lvi.iSubItem = VSI_OPERATOR_COL_GROUP;
			ListView_SetItem(m_wndOperatorList, &lvi);
		}
	}
	VSI_CATCH(hr);
}

int CALLBACK CVsiOperatorAdmin::CompareOperatorProc(LPARAM lParam1, LPARAM lParam2, LPARAM lparamSort)
{
	int iResult = 0;

	CVsiOperatorAdmin *pThis = (CVsiOperatorAdmin*)lparamSort;

	LPCWSTR pszId1 = (LPCWSTR)lParam1;
	LPCWSTR pszId2 = (LPCWSTR)lParam2;

	CComPtr<IVsiOperator> pOperator1;
	CComPtr<IVsiOperator> pOperator2;

	pThis->m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_ID, pszId1, &pOperator1);
	pThis->m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_ID, pszId2, &pOperator2);

	switch (pThis->m_iSortColumn)
	{
	case VSI_OPERATOR_COL_NAME:
		{
			CComHeapPtr<OLECHAR> pszName1;
			pOperator1->GetName(&pszName1);

			CComHeapPtr<OLECHAR> pszName2;
			pOperator2->GetName(&pszName2);

			iResult = lstrcmp(pszName1, pszName2);
		}
		break;
	case VSI_OPERATOR_COL_TYPE:
		{
			VSI_OPERATOR_TYPE dwType1;
			pOperator1->GetType(&dwType1);
			VSI_OPERATOR_TYPE dwType2;
			pOperator2->GetType(&dwType2);

			CString strType1 = (VSI_OPERATOR_TYPE_ADMIN == dwType1) ?
				VSI_LABEL_ADMINISTRATOR : VSI_LABEL_STANDARD;
			CString strType2 = (VSI_OPERATOR_TYPE_ADMIN == dwType2) ?
				VSI_LABEL_ADMINISTRATOR : VSI_LABEL_STANDARD;

			iResult = lstrcmp(strType1, strType2);
		}
		break;
	default:
		_ASSERT(0);
	}

	if (pThis->m_bSortDescending)
	{
		iResult *= -1;
	}

	return iResult;
}

void CVsiOperatorAdmin::SortOperatorList()
{
	ListView_SortItems(m_wndOperatorList, CompareOperatorProc, (LPARAM)this);

	SetOperatorSortIndicator(m_iSortColumn, m_bSortDescending ? HDF_SORTDOWN : HDF_SORTUP);
}

void CVsiOperatorAdmin::SetOperatorSortIndicator(int iCol, int iImage)
{
	CWindow wndHDR = ListView_GetHeader(m_wndOperatorList);

	HDITEM hdItem = { HDI_FORMAT };
	if (iImage == 0)
		hdItem.fmt = HDF_STRING;
	else
		hdItem.fmt = HDI_ORDER | HDF_STRING;

	if (iImage == HDF_SORTUP)
		hdItem.fmt |= HDF_SORTUP;
	else if (iImage == HDF_SORTDOWN)
		hdItem.fmt |= HDF_SORTDOWN;

	Header_SetItem(wndHDR, iCol, &hdItem);
}

HRESULT CVsiOperatorAdmin::UpdatePackageName(LPCWSTR pszFilePath, LPCWSTR pszOperatorName)
{
	HRESULT hr(S_OK);

	try
	{
		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Measurement Manager
		CComPtr<IVsiMsmntManager> pMsmntMgr;

		hr = pServiceProvider->QueryService(
			VSI_SERVICE_MEASUREMENT_MANAGER,
			__uuidof(IVsiMsmntManager),
			(IUnknown**)&pMsmntMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get the settings from the package
		CComPtr<IXMLDOMDocument> pPackageDoc;
		hr = VsiCreateDOMDocument(&pPackageDoc);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		VARIANT_BOOL bLoaded = VARIANT_FALSE;
		hr = pPackageDoc->load(CComVariant(pszFilePath), &bLoaded);
		if ((S_OK != hr) || (VARIANT_FALSE == bLoaded))
		{
			CComPtr<IXMLDOMParseError> pParseError;
			HRESULT hr = pPackageDoc->get_parseError(&pParseError);
			if (SUCCEEDED(hr))
			{
				long lLine = 0;
				CComBSTR bstrDesc;
				HRESULT hrLine = pParseError->get_line(&lLine);
				HRESULT hrDesc = pParseError->get_reason(&bstrDesc);
				if (SUCCEEDED(hrLine) && SUCCEEDED(hrDesc))
				{
					// Display error
					VSI_FAIL(VSI_LOG_ERROR,
						VsiFormatMsg(L"Read measurement package file '%s' failed - line %d: %s",
							pszFilePath, lLine, (LPCWSTR)bstrDesc));
				}
				else if (SUCCEEDED(hrDesc))
				{
					// Display error
					VSI_FAIL(VSI_LOG_ERROR,
						VsiFormatMsg(L"Read measurement package file '%s' failed - %s",
							pszFilePath, (LPCWSTR)bstrDesc));
				}
			}

			VSI_FAIL(VSI_LOG_ERROR, NULL);
		}

		// Read description
		CComPtr<IXMLDOMElement> pDocElement;
		hr = pPackageDoc->get_documentElement(&pDocElement);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failed to get document element");

		CComBSTR bstrDescription(VSI_MSMNT_XML_ATT_DESCRIPTION);

		CComVariant vDescription;
		hr = pDocElement->getAttribute(bstrDescription, &vDescription);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, L"Failed to get package description");

		CString strDescription = pszOperatorName;
		strDescription += L"-";
		strDescription += V_BSTR(&vDescription);

		hr = pDocElement->setAttribute(bstrDescription, CComVariant(strDescription));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pPackageDoc->save(CComVariant(pszFilePath));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiOperatorAdmin::ProcessCustomPackagesPresets()
{
	HRESULT hr(S_OK);

	try
	{
		WCHAR szDataPath[MAX_PATH];
		BOOL bResult = VsiGetApplicationDataDirectory(
			AtlFindStringResourceInstance(IDS_VSI_PRODUCT_NAME),
			szDataPath);
		VSI_VERIFY(bResult, VSI_LOG_ERROR, L"Failure getting application data directory");

		// Get the directory where the custom measurement packages are stored - to copy to each operator's folder
		CString strCustomMeasurementsDir;
		strCustomMeasurementsDir.Format(L"%s%s", szDataPath, VSI_FOLDER_MEASUREMENTS);

		if (0 != _waccess(strCustomMeasurementsDir, 0))
		{
			// Use the measurement folder under the exe path
			WCHAR szPath[MAX_PATH];
			GetModuleFileName(NULL, szPath, _countof(szPath));

			LPWSTR pszSpt = wcsrchr(szPath, L'\\');
			*(pszSpt + 1) = 0;

			PathAppend(szPath, VSI_FOLDER_MEASUREMENTS);
			strCustomMeasurementsDir = szPath;
		}

		strCustomMeasurementsDir += L"\\";

		// Get the directory where the custom presets are stored - to copy to each operator's folder
		// If there is no custom presets, don't fall back to system default presets
		CString strCustomPresetsDir;
		strCustomPresetsDir.Format(L"%s%s", szDataPath, VSI_FOLDER_PRESETS);

		strCustomPresetsDir += L"\\";

		// Get Operators root path
		CString strOperatorsRootDir;
		strOperatorsRootDir.Format(L"%s%s\\", szDataPath, VSI_OPERATORS_FOLDER);

		// Get all custom measurement settings files in system (wide) folder
		std::set<CString> setSettingFiles;
		if (!m_bSecureMode && m_bSecureModeToChange)  // Enable User Management Mode
		{
			CVsiFileIterator fileSettingIter(strCustomMeasurementsDir + L"*.sxml");

			LPWSTR pszFile;
			while (NULL != (pszFile = fileSettingIter.Next()))
			{
				CString strSettingFile(pszFile);
				if (strSettingFile != VSI_MSMNT_PACKAGE_FILE_GENERIC)
				{
					// Don't add system packages (those have .dxml file)
					CString strDefinitionFile(strCustomMeasurementsDir +
						strSettingFile.Left(strSettingFile.GetLength() - _countof(L".sxml") + 1));
					strDefinitionFile += L".dxml";

					bool bNotSysPackage = 0 != _waccess(strDefinitionFile, 0);
					if (bNotSysPackage)
					{
						setSettingFiles.insert(strSettingFile);
					}
				}
			}
		}

		// Loop through each operator's folder
		CComPtr<IVsiEnumOperator> pOperators;
		hr = m_pOperatorMgr->GetOperators(
			VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN,
			&pOperators);
		if (SUCCEEDED(hr))
		{
			CComPtr<IVsiOperator> pOperator;
			while (pOperators->Next(1, &pOperator, NULL) == S_OK)
			{
				CComHeapPtr<OLECHAR> pszName;
				hr = pOperator->GetName(&pszName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CString strOperatorMeasurementsDir;
				strOperatorMeasurementsDir.Format(
					L"%s%s\\" VSI_FOLDER_MEASUREMENTS L"\\",
					strOperatorsRootDir,
					pszName.operator wchar_t *());
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (0 != _waccess(strOperatorMeasurementsDir, 0))
				{
					// Copy settings files
					for (auto iter = setSettingFiles.cbegin(); iter != setSettingFiles.cend(); ++iter)
					{
						BOOL bRet = VsiCopyFiles(strCustomMeasurementsDir + *iter, strOperatorMeasurementsDir);
						VSI_VERIFY(bRet, VSI_LOG_ERROR, L"copy custom packages from system folder failed!");
					}

					// Copy DTD files
					CString strDtd(strCustomMeasurementsDir);
					strDtd += L"*.dtd";

					BOOL bRet = VsiCopyFiles(strDtd, strOperatorMeasurementsDir);
					VSI_VERIFY(bRet, VSI_LOG_ERROR, L"copy custom packages from system folder failed!");
				}

				CString strOperatorPresetsDir;
				strOperatorPresetsDir.Format(
					L"%s%s\\" VSI_FOLDER_PRESETS L"\\",
					strOperatorsRootDir,
					pszName.operator wchar_t *());
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if ((0 != _waccess(strOperatorPresetsDir, 0)) && (0 == _waccess(strCustomPresetsDir, 0)))
				{
					// Copy customized system presets and custom presets when enable User Management Mode
					VsiRemoveDirectory(strOperatorPresetsDir, FALSE);
					BOOL bRet = VsiCopyDirectory(strCustomPresetsDir, strOperatorPresetsDir);
					VSI_VERIFY(bRet, VSI_LOG_ERROR, L"copy custom presets from system folder failed!");
				}

				pOperator.Release();
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiOperatorAdmin::SaveOperatorsSettings()
{
	HRESULT hr(S_OK);

	try
	{
		// Get Operator Config Template Name
		WCHAR szTemplateFile[MAX_PATH];
		int iCount = _countof(szTemplateFile);
		DWORD dwRet = GetModuleFileName(NULL, szTemplateFile, iCount);
		VSI_WIN32_VERIFY(dwRet > 0, VSI_LOG_ERROR, NULL);

		LPWSTR pszSpt = wcsrchr(szTemplateFile, L'\\');
		lstrcpy(pszSpt + 1, VSI_FOLDER_SETTINGS L"\\");
 
		wcscat_s(szTemplateFile, iCount, VSI_OPERATOR_CONFIG_TEMPLATE);

		if (0 == _waccess(szTemplateFile, 0))
		{
			// Get Operators root path
			WCHAR szDataPath[MAX_PATH];
			BOOL bResult = VsiGetApplicationDataDirectory(
				AtlFindStringResourceInstance(IDS_VSI_PRODUCT_NAME),
				szDataPath);
			VSI_VERIFY(bResult, VSI_LOG_ERROR, L"Failure getting application data directory");

			CString strOperatorsRootDir;
			strOperatorsRootDir.Format(L"%s%s\\", szDataPath, VSI_OPERATORS_FOLDER);

			// Loop through each operator's folder
			CComPtr<IVsiEnumOperator> pOperators;
			hr = m_pOperatorMgr->GetOperators(
				VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN,
				&pOperators);
			if (SUCCEEDED(hr))
			{
				VSI_LOG_MSG(VSI_LOG_INFO, L"Generating operator Settings file.");
	
				// Get Operator Config File Name
				CString strConfigFile(VSI_OPERATOR_CONFIG_TEMPLATE);
				strConfigFile = strConfigFile.Left(strConfigFile.ReverseFind(L'.'));  //remove extension

				CString strConfigFileNamePath;
				CComQIPtr<IVsiPdmPersist> pPdmPersist(m_pPdm);
				if (pPdmPersist != NULL)
				{
					strConfigFileNamePath.Format(L"%s%s.config", strOperatorsRootDir, strConfigFile);

					// Persist settings
					hr = pPdmPersist->SaveSettings(szTemplateFile, strConfigFileNamePath);
					if (FAILED(hr))
					{
						VSI_LOG_MSG(VSI_LOG_WARNING, L"Failure persisting settings file.");
					}
				}

				CComPtr<IVsiOperator> pOperator;
				while (pOperators->Next(1, &pOperator, NULL) == S_OK)
				{
					CComHeapPtr<OLECHAR> pszName;
					hr = pOperator->GetName(&pszName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CString strOperatorDir;
					strOperatorDir.Format(
						L"%s%s\\",
						strOperatorsRootDir,
						pszName.operator wchar_t *());
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CString strOperatorConfigFileNamePath;
					strOperatorConfigFileNamePath.Format(L"%s%s.config", strOperatorDir, strConfigFile);

					if ((0 == _waccess(strConfigFileNamePath, 0)) && (0 != _waccess(strOperatorConfigFileNamePath, 0)))
					{
						CString strMsg;
						strMsg.Format(L"Saving Settings for %s.", pszName);

						// Save operator configuration
						if (NULL != pOperator)
						{
							BOOL bRet = VsiCopyFiles(strConfigFileNamePath, strOperatorDir);
							VSI_VERIFY(bRet, VSI_LOG_ERROR, L"copy custom packages from system folder failed!");
						}
					}
					
					pOperator.Release();
				}

				// Clean up
				VsiRemoveFile(strConfigFileNamePath, FALSE);  // Ignore returns - not critical
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

void CVsiOperatorAdmin::RefreshUsageLog()
{
	HRESULT hr(S_OK);

	try
	{
		BOOL bClear(FALSE);
		BOOL bExport(FALSE);

		int iCount = ListView_GetItemCount(m_wndUsageList);
		if (0 < iCount)
		{
			bExport = TRUE;

			CComPtr<IVsiOperator> pOperatorCurrent;
			hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorCurrent);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (NULL != pOperatorCurrent)
			{
				VSI_OPERATOR_TYPE dwType;
				hr = pOperatorCurrent->GetType(&dwType);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				BOOL bAdmin = 0 != (VSI_OPERATOR_TYPE_ADMIN & dwType);
				
				if (bAdmin)
				{
					GetDlgItem(IDC_OPERATOR_ADD).EnableWindow(TRUE);
					GetDlgItem(IDC_OPERATOR_TURNON_USAGE_CHECK).EnableWindow(TRUE);

					bClear = TRUE;
				}
			}
		}

		GetDlgItem(IDC_OPERATOR_USAGE_CLEAR).EnableWindow(bClear);
		GetDlgItem(IDC_OPERATOR_USAGE_EXPORT).EnableWindow(bExport);
	}
	VSI_CATCH(hr);
}
