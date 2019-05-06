/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorImport.cpp
**
**	Description:
**		Implementation of CVsiOperatorImport
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiSaxUtils.h>
#include <VsiWTL.h>
#include <Richedit.h>
#include <ATLComTime.h>
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
#include <VsiStudyXml.h>
#include <VsiStudyModule.h>
#include "VsiOperatorXml.h"
#include "VsiOperatorImport.h"


static const WCHAR g_szExt[] = L".operator.vbak";


CVsiOperatorImport::CVsiOperatorImport() :
	m_msgRefresh(0)
{
}

CVsiOperatorImport::~CVsiOperatorImport()
{
	_ASSERT(m_pApp == NULL);
	_ASSERT(m_pPdm == NULL);
	_ASSERT(m_pOperatorMgr == NULL);
	_ASSERT(m_pCmdMgr == NULL);
}

STDMETHODIMP CVsiOperatorImport::Activate(IVsiApp *pApp, HWND hwndParent)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_INTERFACE(pApp, VSI_LOG_ERROR, NULL);

		m_pApp = pApp;

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Get PDM
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Operator Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_OPERATOR_MGR,
			__uuidof(IVsiOperatorManager),
			(IUnknown**)&m_pOperatorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Command Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CMD_MGR,
			__uuidof(IVsiCommandManager),
			(IUnknown**)&m_pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HWND hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorImport::Deactivate()
{
	HRESULT hr = S_FALSE;

	if (IsWindow())
	{
		BOOL bRet = DestroyWindow();
		_ASSERT(bRet);

		hr = S_OK;
	}

	m_pCmdMgr.Release();
	m_pOperatorMgr.Release();
	m_pPdm.Release();
	m_pApp.Release();

	return hr;
}

STDMETHODIMP CVsiOperatorImport::GetWindow(HWND *phWnd)
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

STDMETHODIMP CVsiOperatorImport::GetAccelerator(HACCEL *phAccel)
{
	UNREFERENCED_PARAMETER(phAccel);
	return E_NOTIMPL;
}

STDMETHODIMP CVsiOperatorImport::GetIsBusy(
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

STDMETHODIMP CVsiOperatorImport::PreTranslateMessage(
	MSG *pMsg, BOOL *pbHandled)
{
	*pbHandled = IsDialogMessage(pMsg);

	return S_OK;
}

STDMETHODIMP CVsiOperatorImport::Initialize(IUnknown *pPropBag)
{
	HRESULT hr = S_OK;

	try
	{
		CComQIPtr<IVsiPropertyBag> pProperties(pPropBag);
		VSI_CHECK_INTERFACE(pProperties, VSI_LOG_ERROR, NULL);

		CComVariant vHwnd;
		hr = pProperties->Read(L"refreshHWND", &vHwnd);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		m_wndRefresh = reinterpret_cast<HWND>(V_UI8(&vHwnd));

		CComVariant vMsg;
		hr = pProperties->Read(L"refreshMsg", &vMsg);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		m_msgRefresh = V_I4(&vMsg);
	}
	VSI_CATCH(hr);

	return hr;
}

LRESULT CVsiOperatorImport::OnInitDialog(UINT, WPARAM, LPARAM, BOOL& bHandled)
{
	HRESULT hr = S_OK;
	bHandled = FALSE;

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
			hr = pLayout->Initialize(m_hWnd, VSI_WL_FLAG_AUTO_RELEASE);
			if (SUCCEEDED(hr))
			{
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

		SetWindowText(CString(MAKEINTRESOURCE(IDS_OPER_IMPORT_OPERATOR)));

		hr = m_pFExplorer.CoCreateInstance(__uuidof(VsiFolderExplorer));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		UpdateAvailableSpace();
		UpdateSizeUI();
		UpdateUI();

		// Refresh tree
		SetDlgItemText(IDC_IE_PATH, CString(MAKEINTRESOURCE(IDS_LOADING)));
		PostMessage(WM_VSI_REFRESH_TREE);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorImport::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	if (m_pFExplorer != NULL)
	{
		m_pFExplorer->Uninitialize();
		m_pFExplorer.Release();
	}

	return 0;
}

LRESULT CVsiOperatorImport::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Import Operator - OK");

	HRESULT hr = S_OK;
	HANDLE hTransaction(INVALID_HANDLE_VALUE);
	BOOL bRet;

	try
	{
		if (m_iAvailableSize < m_iRequiredSize)
		{
			ShowNotEnoughSpaceMessage();
			return 0;
		}

		// Overwriting current operator?
		CComHeapPtr<OLECHAR> pszSrcPath;
		hr = m_pFExplorer->GetSelectedFilePath(&pszSrcPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		VSI_OPERATOR_BACKUP_INFO info;
		hr = GetBackupInfo(pszSrcPath, &info);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
		CComPtr<IVsiOperator> pOperator;
		hr = m_pOperatorMgr->GetCurrentOperator(&pOperator);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComHeapPtr<OLECHAR> pszName;
		hr = pOperator->GetName(&pszName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		bool bCurrentOperator = info.strName == pszName;

		// Check rights
		{
			// Admin can import everything
			// Standard operator can only import himself

			BOOL bAdmin(FALSE);
			hr = m_pOperatorMgr->GetIsAdminAuthenticated(&bAdmin);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (!bAdmin && !bCurrentOperator)
			{
				CString strTitle;
				GetWindowText(strTitle);
	
				VsiMessageBox(
					*this,
					L"Only Administrator can import settings for other operators.",
					strTitle,
					MB_OK | MB_ICONINFORMATION);

				return 0;
			}
		}

		// Overwrite test
		int iOverwrite = CheckForOverwrite();
		if (iOverwrite > 0)
		{
			CString strTitle;
			GetWindowText(strTitle);

			CString strDate;
			BOOL bRet = SystemTimeToTzSpecificLocalTime(NULL, &info.stBackup, &info.stBackup);
			if (bRet)
			{
				CString strTime;

				GetDateFormat(
					LOCALE_USER_DEFAULT,
					DATE_SHORTDATE,
					&info.stBackup,
					NULL,
					strDate.GetBufferSetLength(100),
					100);
				strDate.ReleaseBuffer();

				GetTimeFormat(
					LOCALE_USER_DEFAULT,
					0,
					&info.stBackup,
					NULL,
					strTime.GetBufferSetLength(100),
					100);
				strTime.ReleaseBuffer();

				strDate += L" ";
				strDate += strTime;
			}
	
			CString strMsg;
			if (bCurrentOperator)
			{
				strMsg.Format(
					L"Your settings will be overwritten by this backup created on '%s'.\r\n\r\n"
					L"Are you sure you want to continue?",
					strDate);
			}
			else
			{
				strMsg.Format(
					L"%s's settings will be overwritten by this backup created on '%s'.\r\n\r\n"
					L"Are you sure you want to continue?",
					info.strName, strDate);
			}

			int iRet = VsiMessageBox(*this, strMsg, strTitle, MB_YESNO | MB_ICONWARNING);
			if (IDYES != iRet)
			{
				return 0;
			}
		}

		// Save last used path
		{
			CComHeapPtr<OLECHAR> pszPathName;
			hr = m_pFExplorer->GetSelectedNamespacePath(&pszPathName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			VsiSetRangeValue<LPCWSTR>(
				pszPathName,
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_SYS_PATH_EXPORT_BACKUP,
				m_pPdm);
		}

		CString strPath(pszSrcPath);
		CString strFile(PathFindFileName(pszSrcPath));

		PathRemoveFileSpec(strPath.GetBuffer());
		strPath.ReleaseBuffer();

		WCHAR szRestoreTmpPath[MAX_PATH];
		int iRet = GetTempPath(_countof(szRestoreTmpPath), szRestoreTmpPath);
		VSI_WIN32_VERIFY(0 != iRet, VSI_LOG_ERROR, NULL);

		PathAppend(szRestoreTmpPath, L"Operator.tmp");

		CVsiCabFile cab;
		hr = cab.ExtractFiles(strFile, strPath, szRestoreTmpPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CString strBackupInfo(szRestoreTmpPath);
		PathAppend(strBackupInfo.GetBufferSetLength(MAX_PATH), L"OperatorBackup.xml");
		strBackupInfo.ReleaseBuffer();

		CComPtr<IXMLDOMDocument> pXmlDocOperator;

		hr = VsiCreateDOMDocument(&pXmlDocOperator);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Load the XML document
		VARIANT_BOOL bLoaded = VARIANT_FALSE;
		hr = pXmlDocOperator->load(CComVariant(strBackupInfo), &bLoaded);
		// Handle error
		if (hr != S_OK || bLoaded == VARIANT_FALSE)
		{
			HRESULT hr1;
			CComPtr<IXMLDOMParseError> pParseError;
			hr1 = pXmlDocOperator->get_parseError(&pParseError);
			if (SUCCEEDED(hr1))
			{
				long lError;
				hr1 = pParseError->get_errorCode(&lError);
				CComBSTR bstrDesc;
				hr1 = pParseError->get_reason(&bstrDesc);
				if (SUCCEEDED(hr1))
				{
					// Display error
					VSI_LOG_MSG(VSI_LOG_WARNING,
						VsiFormatMsg(L"Read operator backup file '%s' failed - %s", strBackupInfo, bstrDesc));
				}
				else
				{
					VSI_LOG_MSG(VSI_LOG_WARNING,
						VsiFormatMsg(L"Read operator backup file '%s' failed", strBackupInfo));
				}
			}
			else
			{
				VSI_LOG_MSG(VSI_LOG_WARNING,
					VsiFormatMsg(L"Read operator backup file '%s' failed", strBackupInfo));
			}
		}
		else
		{
			CComPtr<IXMLDOMElement> pXmlElemRoot;
			hr = pXmlDocOperator->get_documentElement(&pXmlElemRoot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
			CComPtr<IXMLDOMNode> pNodeOperator;
			hr = pXmlElemRoot->selectSingleNode(CComBSTR(VSI_OPERATOR_ELM_OPERATOR), &pNodeOperator);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComQIPtr<IXMLDOMElement> pXmlElemOperator(pNodeOperator);
			VSI_CHECK_INTERFACE(pXmlElemOperator, VSI_LOG_ERROR, NULL);

			CComVariant vOperatorName;
			hr = pXmlElemOperator->getAttribute(CComBSTR(VSI_OPERATOR_ATT_NAME), &vOperatorName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IVsiOperator> pOperator;
			hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, V_BSTR(&vOperatorName), &pOperator);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			bool bNewOperator = S_OK != hr;

			if (bNewOperator)
			{
				hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			CComQIPtr<IVsiPersistOperator> pPersistOperator(pOperator);
			VSI_CHECK_INTERFACE(pPersistOperator, VSI_LOG_ERROR, NULL);
			hr = pPersistOperator->LoadXml(pNodeOperator);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (bNewOperator)
			{
				hr = m_pOperatorMgr->AddOperator(pOperator, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			CComHeapPtr<OLECHAR> pOperatorDataPath;
			hr = m_pOperatorMgr->GetOperatorDataPath(
				VSI_OPERATOR_PROP_NAME, V_BSTR(&vOperatorName), &pOperatorDataPath);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// This function will return NULL in WinXP, which is fine. The API will 
			// revert to the non-transacted methods internally if necessary.
			hTransaction = VsiTxF::VsiCreateTransaction(
				NULL,
				0,		
				TRANSACTION_DO_NOT_PROMOTE,	
				0,		
				0,		
				INFINITE,
				L"Operator import transaction");
			VSI_WIN32_VERIFY(INVALID_HANDLE_VALUE != hTransaction, VSI_LOG_ERROR, L"Failure creating transaction");

			if (VsiGetFileExists(pOperatorDataPath))
			{
				bRet = VsiTxF::VsiRemoveAllDirectoryTransacted(pOperatorDataPath, hTransaction);
				VSI_VERIFY(bRet, VSI_LOG_ERROR, NULL);
			}

			CString strBackupData(szRestoreTmpPath);
			PathAppend(strBackupData.GetBufferSetLength(MAX_PATH), L"\\Data");
			strBackupData.ReleaseBuffer();

			if (VsiGetFileExists(strBackupData))
			{
				bRet = VsiTxF::VsiMoveFileTransacted(
					strBackupData,
					pOperatorDataPath,	
					NULL,	
					NULL,	
					MOVEFILE_COPY_ALLOWED,
					hTransaction);
				VSI_VERIFY(bRet, VSI_LOG_ERROR, L"Move file failed");
			}

			bRet = VsiTxF::VsiCommitTransaction(hTransaction);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, L"Create transaction failed");
		}
		
		m_wndRefresh.SendMessage(m_msgRefresh, (WPARAM)info.strName.GetString());

		CComVariant vCurrentOperatorModified(bCurrentOperator);
		m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_OPERATOR_IMPORT, &vCurrentOperatorModified);
	}
	VSI_CATCH(hr);

	if (INVALID_HANDLE_VALUE != hTransaction)
	{
		BOOL bRet = VsiTxF::VsiCloseTransaction(hTransaction);
		if (!bRet)
		{
			VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure closing transaction");
		}
	}

	return 0;
}

LRESULT CVsiOperatorImport::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Import Operator - Cancel");

	CComVariant vCurrentOperatorModified(false);
	m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_OPERATOR_IMPORT, &vCurrentOperatorModified);

	return 0;
}

BOOL CVsiOperatorImport::GetAvailablePath(LPWSTR pszBuffer, int iBufferLength)
{
	HRESULT hr = S_OK;
	CComVariant vPath;

	try
	{
		CComHeapPtr<OLECHAR> pszPath;
		hr = m_pOperatorMgr->GetOperatorDataPath(VSI_OPERATOR_PROP_NONE, NULL, &pszPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		int iRet = wcscpy_s(pszBuffer, iBufferLength, pszPath);
		VSI_VERIFY(0 == iRet, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return (S_OK == hr);
}

int CVsiOperatorImport::CheckForOverwrite()
{
	HRESULT hr = S_OK;
	BOOL bOverwrite = FALSE;

	try
	{
		CComHeapPtr<OLECHAR> pszPath;
		hr = m_pFExplorer->GetSelectedFilePath(&pszPath);
		if ((S_FALSE == hr) || (0 == lstrlen(pszPath)))
		{
			return bOverwrite;  // this returns if nothing is selected
		}

		CString strFileName(PathFindFileName(pszPath));

		if (strFileName.GetLength() > (_countof(g_szExt) - 1))
		{
			VSI_OPERATOR_BACKUP_INFO info;
			hr = GetBackupInfo(pszPath, &info);
			if (SUCCEEDED(hr))
			{
				CComPtr<IVsiOperator> pOperator;
				hr = m_pOperatorMgr->GetOperator(
					VSI_OPERATOR_PROP_NAME,
					info.strName,
					&pOperator);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					bOverwrite = TRUE;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return bOverwrite ? 1 : 0;
}

void CVsiOperatorImport::UpdateUI()
{
	CVsiImportExportUIHelper<CVsiOperatorImport>::UpdateUI();

	HRESULT hr = S_OK;
	BOOL bEnableOK = FALSE;

	try
	{
		m_iRequiredSize = 0;

		WCHAR szPath[VSI_MAX_PATH];
		if (GetSelectedPath(szPath, _countof(szPath)))
		{
			if (IsBackup(szPath))
			{
				LARGE_INTEGER size = { 0 };
				hr = VsiGetFileSize(szPath, &size);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				m_iRequiredSize = size.LowPart;
				m_iRequiredSize *= 10;  // Compression

				bEnableOK = TRUE;
			}
		}

		// Check for sufficient space in the current location. If there is not enough
		// then enable the message and show them how much
		UpdateAvailableSpace();
		UpdateSizeUI();

		GetDlgItem(IDOK).EnableWindow(bEnableOK);
	}
	VSI_CATCH(hr);
}

LRESULT CVsiOperatorImport::OnTreeSelChanged(int, LPNMHDR, BOOL &)
{
	if (m_pFExplorer != NULL)
	{
		UpdateUI();
	}

	return 0;
}

LRESULT CVsiOperatorImport::OnCheckForBackup(int, LPNMHDR pnmh, BOOL &)
{
	HRESULT hr(S_OK);

	try
	{
		NM_FOLDER_GETITEMINFO* pInfo = (NM_FOLDER_GETITEMINFO*)pnmh;

		if (IsBackup(pInfo->pszFolderPath))
		{
			VSI_OPERATOR_BACKUP_INFO info;
			hr = GetBackupInfo(pInfo->pszFolderPath, &info);
			if (SUCCEEDED(hr))
			{
				BOOL bRet = SystemTimeToTzSpecificLocalTime(NULL, &info.stBackup, &info.stBackup);
				if (bRet)
				{
					CString strDate;
					CString strTime;

					GetDateFormat(
						LOCALE_USER_DEFAULT,
						DATE_SHORTDATE,
						&info.stBackup,
						NULL,
						strDate.GetBufferSetLength(100),
						100);
					strDate.ReleaseBuffer();

					GetTimeFormat(
						LOCALE_USER_DEFAULT,
						0,
						&info.stBackup,
						NULL,
						strTime.GetBufferSetLength(100),
						100);
					strTime.ReleaseBuffer();

					strDate += L" ";
					strDate += strTime;

					CString strName;
					strName.Format(L"%s (%s)", info.strName.GetString(), strDate.GetString());

					wcsncpy_s(
						pInfo->pszName,
						pInfo->iNameMax,
						strName,
						pInfo->iNameMax);

					LONG lVsiImageIndex = 0;
					hr = m_pFExplorer->GetVevoImageIndex(&lVsiImageIndex);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					pInfo->pItem->iImage = lVsiImageIndex;
					pInfo->pItem->iSelectedImage = lVsiImageIndex;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiOperatorImport::OnRefreshTree(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr(S_OK);

	try
	{
		VSI_FE_FLAG dwFlags(VSI_FE_FLAG_NONE);
		CComPtr<IFolderFilter> pFolderFilter;

		hr = pFolderFilter.CoCreateInstance(__uuidof(VsiFolderFilter));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiFolderFilter> pFilterSettings(pFolderFilter);
		VSI_CHECK_INTERFACE(pFilterSettings, VSI_LOG_ERROR, NULL);

		LPCOLESTR pszFiler = L".operator.vbak";
		hr = pFilterSettings->SetFileFilters(1, &pszFiler);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		DWORD dwState = VsiGetBitfieldValue<DWORD>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_STATE, m_pPdm);
		if (0 < (VSI_SYS_STATE_SYSTEM & dwState) &&
			0 == (VSI_SYS_STATE_ENGINEER_MODE & dwState))
		{
			// For Sierra system, shows removable and network drive
			dwFlags = VSI_FE_FLAG_SHOW_NETWORK | VSI_FE_FLAG_SHOW_REMOVEABLE;
		}
		else
		{
			// For workstation and eng mode, shows Desktop
			dwFlags = VSI_FE_FLAG_SHOW_DESKTOP | VSI_FE_FLAG_EXPLORER;

			hr = pFilterSettings->SetFolderFilter(VSI_FOLDER_FILTER_WORKSTATION);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
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

LRESULT CVsiOperatorImport::OnSetSelection(UINT, WPARAM, LPARAM, BOOL&)
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
			SetDlgItemText(IDC_IE_PATH, L"");
		}
	}
	VSI_CATCH(hr);

	return 0;
}

bool CVsiOperatorImport::IsBackup(LPCWSTR pszPath)
{
	bool bIsBackup(false);

	CString strTest(PathFindFileName(pszPath));

	if (strTest.GetLength() > (_countof(g_szExt) - 1))
	{
		if (strTest.Right(_countof(g_szExt) - 1) == g_szExt)
		{
			bIsBackup = true;
		}
	}

	return bIsBackup;
}

HRESULT CVsiOperatorImport::GetBackupInfo(LPCWSTR pszPath, VSI_OPERATOR_BACKUP_INFO *pInfo)
{
	HRESULT hr = S_OK;

	try
	{
		CString strPath(pszPath);
		CString strFile(PathFindFileName(pszPath));

		PathRemoveFileSpec(strPath.GetBuffer());
		strPath.ReleaseBuffer();

		WCHAR szBackupInfoPath[MAX_PATH];
		int iRet = GetTempPath(_countof(szBackupInfoPath), szBackupInfoPath);
		VSI_WIN32_VERIFY(0 != iRet, VSI_LOG_ERROR, NULL);

		CVsiCabFile cab;
		hr = cab.ExtractFile(L"OperatorBackup.xml", strFile, strPath, szBackupInfoPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		PathAppend(szBackupInfoPath, L"OperatorBackup.xml");

		CComPtr<IXMLDOMDocument> pXmlDocOperator;

		hr = VsiCreateDOMDocument(&pXmlDocOperator);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Load the XML document
		VARIANT_BOOL bLoaded = VARIANT_FALSE;
		hr = pXmlDocOperator->load(CComVariant(szBackupInfoPath), &bLoaded);
		// Handle error
		if (hr != S_OK || bLoaded == VARIANT_FALSE)
		{
			HRESULT hr1;
			CComPtr<IXMLDOMParseError> pParseError;
			hr1 = pXmlDocOperator->get_parseError(&pParseError);
			if (SUCCEEDED(hr1))
			{
				long lError;
				hr1 = pParseError->get_errorCode(&lError);
				CComBSTR bstrDesc;
				hr1 = pParseError->get_reason(&bstrDesc);
				if (SUCCEEDED(hr1))
				{
					// Display error
					VSI_LOG_MSG(VSI_LOG_WARNING,
						VsiFormatMsg(L"Read operator backup file '%s' failed - %s", szBackupInfoPath, bstrDesc));
				}
				else
				{
					VSI_LOG_MSG(VSI_LOG_WARNING,
						VsiFormatMsg(L"Read operator backup file '%s' failed", szBackupInfoPath));
				}
			}
			else
			{
				VSI_LOG_MSG(VSI_LOG_WARNING,
					VsiFormatMsg(L"Read operator backup file '%s' failed", szBackupInfoPath));
			}
		}
		else
		{
			CComPtr<IXMLDOMElement> pXmlElemRoot;
			hr = pXmlDocOperator->get_documentElement(&pXmlElemRoot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
			CComVariant vDate;
			hr = pXmlElemRoot->getAttribute(CComBSTR(VSI_OPERATOR_BACKUP_ATT_DATE), &vDate);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = VsiXmlReadDateTime(V_BSTR(&vDate), wcslen(V_BSTR(&vDate)), &pInfo->stBackup);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IXMLDOMNode> pNodeOperator;
			hr = pXmlElemRoot->selectSingleNode(CComBSTR(VSI_OPERATOR_ELM_OPERATOR), &pNodeOperator);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComQIPtr<IXMLDOMElement> pXmlElemOperator(pNodeOperator);
			VSI_CHECK_INTERFACE(pXmlElemOperator, VSI_LOG_ERROR, NULL);

			CComVariant vName;
			hr = pXmlElemOperator->getAttribute(CComBSTR(VSI_OPERATOR_ATT_NAME), &vName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			pInfo->strName = V_BSTR(&vName);
		}

		// Delete file, ignore errors
		DeleteFile(szBackupInfoPath);
	}
	VSI_CATCH_(err)
	{
		hr = err;
	}

	return hr;
}
