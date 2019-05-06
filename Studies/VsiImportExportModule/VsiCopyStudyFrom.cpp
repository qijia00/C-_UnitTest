/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiCopyStudyFrom.cpp
**
**	Description:
**		Implementation of CVsiCopyStudyFrom
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiWTL.h>
#include <VsiSaxUtils.h>
#include <Richedit.h>
#include <ATLComTime.h>
#include <VsiGlobalDef.h>
#include <VsiThemeColor.h>
#include <VsiServiceProvider.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiServiceKey.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiStudyXml.h>
#include <VsiStudyModule.h>
#include <VsiImportExportXml.h>
#include "VsiImpExpUtility.h"
#include "VsiReportExportStatus.h"
#include "VsiCopyStudyFrom.h"

class CVsiCopyStudyFromConfirmDlg :
	public CDialogImpl<CVsiCopyStudyFromConfirmDlg>
{
public:
	enum { IDD = IDD_COPY_STUDY_FROM_CONFIRM };

	enum
	{
		ACTION_OVERWRITE = 100,
		ACTION_SKIP,
		ACTION_COPY,
	};

protected:
	BEGIN_MSG_MAP(CVsiCopyStudyFromConfirmDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
	{
		HRESULT hr(S_OK);

		try
		{
			// Sets font
			HFONT hFont;
			VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
			VsiThemeRecurSetFont(m_hWnd, hFont);

			// Sets header and footer
			HIMAGELIST hilHeader(NULL);
			VsiThemeGetImageList(VSI_THEME_IL_HEADER_QUESTION, &hilHeader);
			VsiThemeDialogBox(*this, TRUE, TRUE, hilHeader);

			// Text branding
			{
				SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_COPY_STUDY_FROM)));
				GetDlgItem(IDC_COPYSTUDY_CONFIRM_LABEL).SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_ONE_OR_MORE_EXISTS)));
				GetDlgItem(IDC_CSFC_OVERWRITE).SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_OVERWRITE_SERIES)));
				GetDlgItem(IDC_CSFC_SKIP).SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_SKIP_SERIES)));
				GetDlgItem(IDC_CSFC_COPY).SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_CREATE_NEW_STUDY)));
			}

			CheckDlgButton(IDC_CSFC_OVERWRITE, BST_CHECKED);
		}
		VSI_CATCH(hr);

		if (FAILED(hr))
			EndDialog(IDCANCEL);

		return 0;
	}

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&)
	{
		VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Copy Study From Confirm - OK");

		int iRet(IDCANCEL);

		if (BST_CHECKED == IsDlgButtonChecked(IDC_CSFC_OVERWRITE))
		{
			iRet = ACTION_OVERWRITE;
		}
		else if (BST_CHECKED == IsDlgButtonChecked(IDC_CSFC_SKIP))
		{
			iRet = ACTION_SKIP;
		}
		else if (BST_CHECKED == IsDlgButtonChecked(IDC_CSFC_COPY))
		{
			iRet = ACTION_COPY;
		}

		EndDialog(iRet);

		return 0;
	}

	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
	{
		VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Copy Study From Confirm - Cancel");

		EndDialog(IDCANCEL);

		return 0;
	}
};

CVsiCopyStudyFrom::CVsiCopyStudyFrom() :
	m_dwSink(0),
	m_dwEventOperatorListUpdate(0),
	m_hbrushMandatory(NULL)
{
}

CVsiCopyStudyFrom::~CVsiCopyStudyFrom()
{
	_ASSERT(m_pApp == NULL);
	_ASSERT(m_pCmdMgr == NULL);
	_ASSERT(m_pPdm == NULL);
	_ASSERT(m_pOperatorMgr == NULL);
	_ASSERT(m_pStudyMgr == NULL);
	_ASSERT(NULL == m_hbrushMandatory);
}

STDMETHODIMP CVsiCopyStudyFrom::Activate(
	IVsiApp *pApp,
	HWND hwndParent)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_INTERFACE(pApp, VSI_LOG_ERROR, NULL);

		m_pApp = pApp;

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Get Study Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_STUDY_MGR,
			__uuidof(IVsiStudyManager),
			(IUnknown**)&m_pStudyMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Command Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CURSOR_MANAGER,
			__uuidof(IVsiCursorManager),
			(IUnknown**)&m_pCursorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

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

		// Get Operator Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_OPERATOR_MGR,
			__uuidof(IVsiOperatorManager),
			(IUnknown**)&m_pOperatorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Operator List update event id
		CComPtr<IVsiParameter> pParamEvent;
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
			VSI_PARAMETER_EVENTS_GENERAL_OPERATOR_LIST_UPDATE,
			&pParamEvent);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParamEvent->GetId(&m_dwEventOperatorListUpdate);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Events sink
		CComQIPtr<IVsiPdmEvents> pPdmEvents(m_pPdm);
		hr = pPdmEvents->AdviseParamUpdateEventSink(
			static_cast<IVsiParamUpdateEventSink*>(this), &m_dwSink);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HWND hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiCopyStudyFrom::Deactivate()
{
	HRESULT hr = S_FALSE;

	if (IsWindow())
	{
		BOOL bRet = DestroyWindow();
		_ASSERT(bRet);

		hr = S_OK;
	}

	// Unadvise
	if (0 != m_dwSink)
	{
		CComQIPtr<IVsiPdmEvents> pPdmEvents(m_pPdm);
		hr = pPdmEvents->UnadviseParamUpdateEventSink(m_dwSink);
		if (FAILED(hr))
		{
			VSI_LOG_MSG(VSI_LOG_WARNING, L"UnadviseParamUpdateEventSink failed");
		}
		m_dwSink = 0;
	}

	m_pItemPreview.Release();

	m_pOperatorMgr.Release();

	m_pPdm.Release();

	m_pCmdMgr.Release();

	m_pStudyMgr.Release();

	m_pApp.Release();

	return hr;
}

STDMETHODIMP CVsiCopyStudyFrom::GetWindow(HWND *phWnd)
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

STDMETHODIMP CVsiCopyStudyFrom::GetAccelerator(HACCEL *phAccel)
{
	UNREFERENCED_PARAMETER(phAccel);
	return E_NOTIMPL;
}

STDMETHODIMP CVsiCopyStudyFrom::GetIsBusy(
	DWORD dwStateCurrent,
	DWORD dwState,
	BOOL bTryRelax,
	BOOL *pbBusy)
{
	UNREFERENCED_PARAMETER(dwStateCurrent);
	UNREFERENCED_PARAMETER(dwState);
	UNREFERENCED_PARAMETER(bTryRelax);

	*pbBusy = FALSE;

	return S_OK;
}

STDMETHODIMP CVsiCopyStudyFrom::PreTranslateMessage(
	MSG *pMsg, BOOL *pbHandled)
{
	switch (pMsg->message)
	{
	case WM_KEYDOWN:
		{
			if (GetFocus() == GetDlgItem(IDC_CSF_STUDY_LIST))
			{
				switch (pMsg->wParam)
				{
				case L'A':
					{
						if (GetKeyState(VK_CONTROL) < 0)
						{
							m_pDataList->SelectAll();
						}
					}
					break;
				}
			}
		}
		break;

	default:
		// Nothing
		break;
	}

	*pbHandled = IsDialogMessage(pMsg);

	return S_OK;
}

STDMETHODIMP CVsiCopyStudyFrom::OnParamUpdate(DWORD *pdwParamIds, DWORD dwCount)
{
	for (DWORD i = 0; i < dwCount; ++i, ++pdwParamIds)
	{
		if (*pdwParamIds == m_dwEventOperatorListUpdate)
		{
			m_wndOperator.Fill(m_strOwnerOperator);

			break;
		}
	}

	return S_OK;
}

LRESULT CVsiCopyStudyFrom::OnInitDialog(UINT, WPARAM, LPARAM, BOOL& bHandled)
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

				pLayout->AddControl(0, IDC_IE_PATH, VSI_WL_SIZE_X);
				pLayout->AddControl(0, IDC_IE_FOLDER_TREE, VSI_WL_SIZE_XY);
				pLayout->AddControl(0, IDC_CSF_STUDY_LIST_LABEL, VSI_WL_MOVE_Y);
				pLayout->AddControl(0, IDC_CSF_STUDY_LIST, VSI_WL_MOVE_Y | VSI_WL_SIZE_X);
				pLayout->AddControl(0, IDC_SB_PREVIEW, VSI_WL_MOVE_X | VSI_WL_SIZE_Y);
				pLayout->AddControl(0, IDC_SB_INFO, VSI_WL_MOVE_X | VSI_WL_SIZE_Y);
				pLayout->AddControl(0, IDC_CSF_REMOVE, VSI_WL_MOVE_XY);

				pLayout->AddControl(0, IDC_IE_SELECTED, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_REQUIRED_LABEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_REQUIRED, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_AVAILABLE_LABEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_AVAILABLE, VSI_WL_MOVE_X);

				pLayout.Detach();
			}
		}

		m_hbrushMandatory = CreateSolidBrush(VSI_COLOR_MANDATORY);
		VSI_WIN32_VERIFY(NULL != m_hbrushMandatory, VSI_LOG_ERROR, NULL);

		// Checkboxes in folder tree
		CWindow wndTree(GetDlgItem(IDC_IE_FOLDER_TREE));
		TreeView_SetExtendedStyle(wndTree, TVS_EX_PARTIALCHECKBOXES, TVS_EX_PARTIALCHECKBOXES);
		wndTree.ModifyStyle(0, TVS_CHECKBOXES);

		// Text Branding
		{
			GetDlgItem(IDC_IE_SELECTED).SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_STUDIES_SELECTED)));
			GetDlgItem(IDC_CSF_STUDY_LIST_LABEL).SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_STUDIES_SELECTED)));
		}

		// Initialize Study List
		static const VSI_LVCOLUMN plvc[] =
		{
			{ LVCF_WIDTH | LVCF_TEXT, 0, 260, IDS_CPYSTY_COLUMN_NAME, VSI_DATA_LIST_COL_NAME, -1, 0, TRUE },
			{ LVCF_WIDTH | LVCF_TEXT, 0, 120, IDS_CPYSTY_COLUMN_DATE, VSI_DATA_LIST_COL_DATE, -1, 1, TRUE },
			{ LVCF_WIDTH | LVCF_TEXT | LVCF_FMT, LVCFMT_RIGHT, 100, IDS_CPYSTY_COLUMN_SIZE, VSI_DATA_LIST_COL_SIZE, -1, 2, TRUE },
			{ LVCF_WIDTH | LVCF_TEXT, 0, 120, IDS_CPYSTY_COLUMN_OWNER, VSI_DATA_LIST_COL_OWNER, -1, 3, TRUE },
		};

		hr = m_pDataList.CoCreateInstance(__uuidof(VsiDataListWnd));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating Data List");

		hr = m_pDataList->Initialize(
			GetDlgItem(IDC_CSF_STUDY_LIST),
			VSI_DATA_LIST_FLAG_ITEM_STATUS_CALLBACK | VSI_DATA_LIST_FLAG_NO_SESSION_LINK,
			VSI_DATA_LIST_STUDY, m_pApp, plvc, _countof(plvc));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = m_pDataList->SetEmptyMessage(CString(MAKEINTRESOURCE(IDS_CPYSTY_NO_STUDIES_SELECTED)));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Init study overview control
		m_studyOverview.SubclassWindow(GetDlgItem(IDC_SB_INFO));
		m_studyOverview.Initialize(CVsiStudyOverviewCtl::VSI_STUDY_OVERVIEW_DISPLAY_DETAILS);

		SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_COPY_STUDY_FROM)));

		hr = m_pFExplorer.CoCreateInstance(__uuidof(VsiFolderExplorer));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		m_wndImageList = GetDlgItem(IDC_SB_PREVIEW);

		// Operator
		{
			CComHeapPtr<OLECHAR> pszCurrent;

			CComPtr<IVsiOperator> pOperatorCurrent;
			hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorCurrent);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr)
			{
				hr = pOperatorCurrent->GetName(&pszCurrent);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				m_strOwnerOperator = pszCurrent;
			}

			// Init operator
			hr = m_wndOperator.Initialize(
				FALSE,
				TRUE,
				FALSE,
				GetDlgItem(IDC_CF_OPERATOR_LIST),
				m_pOperatorMgr,
				m_pCmdMgr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_wndOperator.SendMessage(CB_SETEXTENDEDUI, (WPARAM)TRUE, 0L);

			m_wndOperator.Fill(m_strOwnerOperator);

			BOOL bSecureMode(FALSE);
			bSecureMode = VsiGetBooleanValue<BOOL>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
				VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
			BOOL bAdmin(FALSE);
			if (bSecureMode)
			{
				hr = m_pOperatorMgr->GetIsAdminAuthenticated(&bAdmin);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			m_wndOperator.EnableWindow((bSecureMode && (!bAdmin))? FALSE : TRUE);
		}

		UpdateAvailableSpace();
		UpdateUI();

		// Refresh tree
		SetDlgItemText(IDC_IE_PATH, L"Loading...");
		PostMessage(WM_VSI_REFRESH_TREE);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	StopPreview();

	if (NULL != m_hbrushMandatory)
	{
		DeleteObject(m_hbrushMandatory);
		m_hbrushMandatory = NULL;
	}

	// Clean up operator list
	m_wndOperator.Uninitialize();

	m_setSelectedStudy.clear();
	m_setRemovedSeries.clear();
	m_setSelectedStudyId.clear();

	m_studyOverview.UnsubclassWindow();
	m_studyOverview.Uninitialize();

	m_pDataList->Uninitialize();

	if (m_pFExplorer != NULL)
	{
		m_pFExplorer->Uninitialize();
		m_pFExplorer.Release();

		HRESULT hr(S_OK);
		try
		{
			if (!m_strLastAddPath.IsEmpty())
			{
				CVsiRange<LPCWSTR> pParamPath;
				hr = m_pPdm->GetParameter(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
					VSI_PARAMETER_SYS_PATH_IMPORT_STUDY, &pParamPath);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				pParamPath.SetValue((LPWSTR)(LPCWSTR)m_strLastAddPath);
			}
		}
		VSI_CATCH(hr);
	}

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnSettingChange(UINT, WPARAM, LPARAM, BOOL &bHandled)
{
	bHandled = FALSE;

	// Refresh preview
	if (VSI_DATA_LIST_STUDY != m_typeSel)
	{
		StopPreview(VSI_PREVIEW_LIST_CLEAR);
		m_pItemPreview.Release();
		m_typeSel = VSI_DATA_LIST_NONE;
	}
	else
	{
		m_studyOverview.Refresh();
	}

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnCtlColorEdit(UINT, WPARAM wp, LPARAM lp, BOOL &bHandled)
{
	BOOL bMandatory(FALSE);

	if (GetDlgItem(IDC_CF_OPERATOR_LIST) == (HWND)lp)
	{
		int index = (int)::SendMessage((HWND)lp, CB_GETCURSEL, 0, 0);
		if (CB_ERR != index)
		{
			LPCWSTR pszId = (LPCWSTR)::SendMessage((HWND)lp, CB_GETITEMDATA, index, 0);
			if (IS_INTRESOURCE(pszId))
			{
				bMandatory = TRUE;
			}
		}
		else
		{
			bMandatory = TRUE;
		}
	}

	if (bMandatory)
	{
		SetBkColor((HDC)wp, VSI_COLOR_MANDATORY);

		return (LRESULT)m_hbrushMandatory;
	}
	else
	{
		bHandled = FALSE;
	}

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnRefreshTree(UINT, WPARAM, LPARAM, BOOL&)
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
			dwFlags = VSI_FE_FLAG_SHOW_REMOVEABLE |
				VSI_FE_FLAG_SHOW_NETWORK;

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

LRESULT CVsiCopyStudyFrom::OnSetSelection(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	CWaitCursor wait;

	try
	{
		// Set default selected folder
		CVsiRange<LPCWSTR> pParamPath;
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_SYS_PATH_IMPORT_STUDY, &pParamPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiParameterRange> pRange(pParamPath.m_pParam);
		VSI_CHECK_INTERFACE(pRange, VSI_LOG_ERROR, NULL);

		CComVariant vPath;
		hr = pRange->GetValue(&vPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = m_pFExplorer->SetSelectedNamespacePath(V_BSTR(&vPath));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK != hr)
		{
			UpdateUI();
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnUpdatePreview(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr;

	try
	{
		CComPtr<IVsiEnumStudy> pEnumStudyPreFiltered;
		CComPtr<IVsiEnumSeries> pEnumSeriesPreFiltered;
		CComPtr<IVsiEnumImage> pEnumImagePreFiltered;

		int iStudyPreFiltered = 0;
		int iSeriesPreFiltered = 0;
		int iImagePreFiltered = 0;

		CComPtr<IVsiEnumStudy> pEnumStudy;
		CComPtr<IVsiEnumSeries> pEnumSeries;
		CComPtr<IVsiEnumImage> pEnumImage;

		int iStudy = 0;
		int iSeries = 0;
		int iImage = 0;

		VSI_DATA_LIST_COLLECTION listSel, listSelPreFiltered;

		memset((void*)&listSelPreFiltered, 0, sizeof(listSelPreFiltered));
		listSelPreFiltered.dwFlags = VSI_DATA_LIST_FILTER_NONE;
		listSelPreFiltered.ppEnumStudy = &pEnumStudyPreFiltered;
		listSelPreFiltered.piStudyCount = &iStudyPreFiltered;
		listSelPreFiltered.ppEnumSeries = &pEnumSeriesPreFiltered;
		listSelPreFiltered.piSeriesCount = &iSeriesPreFiltered;
		listSelPreFiltered.ppEnumImage = &pEnumImagePreFiltered;
		listSelPreFiltered.piImageCount = &iImagePreFiltered;
		hr = m_pDataList->GetSelection(&listSelPreFiltered);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		memset((void*)&listSel, 0, sizeof(listSel));
		listSel.dwFlags = VSI_DATA_LIST_FILTER_SELECT_PARENT;
		listSel.ppEnumStudy = &pEnumStudy;
		listSel.piStudyCount = &iStudy;
		listSel.ppEnumSeries = &pEnumSeries;
		listSel.piSeriesCount = &iSeries;
		listSel.ppEnumImage = &pEnumImage;
		listSel.piImageCount = &iImage;
		hr = m_pDataList->GetSelection(&listSel);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Gets focused item
		int iFocused(-1);
		VSI_DATA_LIST_TYPE type = VSI_DATA_LIST_NONE;
		CComPtr<IUnknown> pUnkItemFocused;
		{
			hr = m_pDataList->GetSelectedItem(&iFocused);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (iFocused >= 0)
			{
				hr = m_pDataList->GetItem(iFocused, &type, &pUnkItemFocused);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Check for corruption
				if (VSI_DATA_LIST_STUDY == type)
				{
					CComQIPtr<IVsiStudy> pStudy(pUnkItemFocused);
					VSI_CHECK_INTERFACE(pStudy, VSI_LOG_ERROR, NULL);

					VSI_STUDY_ERROR dwError(VSI_STUDY_ERROR_NONE);
					pStudy->GetErrorCode(&dwError);

					// Export is still supported when some of the series have problem
					if (VSI_STUDY_ERROR_NONE != dwError && VSI_STUDY_ERROR_LOAD_SERIES != dwError)
					{
						StopPreview(VSI_PREVIEW_LIST_CLEAR);
						m_pItemPreview = NULL;
						SetStudyInfo(NULL, false);
						iFocused = -1;
					}
				}
				else if (VSI_DATA_LIST_SERIES == type)
				{
					CComQIPtr<IVsiSeries> pSeries(pUnkItemFocused);
					VSI_CHECK_INTERFACE(pSeries, VSI_LOG_ERROR, NULL);

					VSI_SERIES_ERROR dwError(VSI_SERIES_ERROR_NONE);
					pSeries->GetErrorCode(&dwError);

					// Export is still supported when some of the image have problem
					if (VSI_SERIES_ERROR_NONE != dwError && VSI_SERIES_ERROR_LOAD_IMAGES != dwError)
					{
						StopPreview(VSI_PREVIEW_LIST_CLEAR);
						m_pItemPreview = NULL;
						SetStudyInfo(NULL, false);
						iFocused = -1;
					}
				}
			}
		}

		if (iFocused >= 0)
		{
			if (VSI_DATA_LIST_STUDY == type)
			{
				CComQIPtr<IVsiStudy> pStudy(pUnkItemFocused);
				VSI_CHECK_INTERFACE(pStudy, VSI_LOG_ERROR, NULL);

				if (VSI_DATA_LIST_SERIES == m_typeSel)
				{
					StopPreview(VSI_PREVIEW_LIST_CLEAR);
				}

				m_pItemPreview = pStudy;
				m_typeSel = VSI_DATA_LIST_STUDY;

				// Preview study
				SetStudyInfo(pStudy);
			}
			else if (VSI_DATA_LIST_SERIES == type)
			{
				CComQIPtr<IVsiSeries> pSeries(pUnkItemFocused);
				VSI_CHECK_INTERFACE(pSeries, VSI_LOG_ERROR, NULL);

				if (m_pItemPreview != NULL)
				{
					if (VSI_DATA_LIST_STUDY == m_typeSel)
					{
						SetStudyInfo(NULL);

						m_pItemPreview = NULL;
					}
					else
					{
						if (! m_pItemPreview.IsEqualObject(pSeries))
						{
							StopPreview(VSI_PREVIEW_LIST_CLEAR);
							m_pItemPreview = NULL;
						}
					}
				}

				if (NULL == m_pItemPreview)
				{
					SetStudyInfo(NULL);
					m_pItemPreview = pSeries;
					m_typeSel = VSI_DATA_LIST_SERIES;

					CComQIPtr<IVsiPersistSeries> pPersist(m_pItemPreview);
					hr = pPersist->LoadSeriesData(NULL, VSI_DATA_TYPE_IMAGE_LIST, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					StartPreview();
				}
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Copy Study From - OK");

	HRESULT hr = S_OK;
	INT64 iTotalSizeToCopy = 0, iTotalSizeStudy = 0, iTotalSizeSeries = 0;
	CComHeapPtr<OLECHAR> pszPath;
	CComPtr<IXMLDOMElement> pElmRoot;
	LPWSTR pszSpt = NULL;
	CComVariant vId, vNewId, vLocked, var;
	BOOL bOverwrite = FALSE;
	CComVariant vOwner;
	CString strSearchSeriesFolder, strSearchTargetSeriesPath;

	try
	{
		if (m_iAvailableSize < m_iRequiredSize)
		{
			ShowNotEnoughSpaceMessage();
			return 0;
		}

		int iOverwriteAction(0);
		CString strAction;
		int iOverwrite = CheckForOverwrite();
		if (0 < iOverwrite)
		{
			// Overwrite, Skip, Create Copy or  Cancel
			CVsiCopyStudyFromConfirmDlg dlg;
			iOverwriteAction = (int)dlg.DoModal();
			if (IDCANCEL == iOverwriteAction)
				return 0;

			if (CVsiCopyStudyFromConfirmDlg::ACTION_OVERWRITE == iOverwriteAction)
			{
				strAction = L"Overwrite";
			}
			else if (CVsiCopyStudyFromConfirmDlg::ACTION_SKIP == iOverwriteAction)
			{
				strAction = L"Skip";
			}
			else if (CVsiCopyStudyFromConfirmDlg::ACTION_COPY == iOverwriteAction)
			{
				strAction = L"Copy";
			}
		}
		else
		{
			strAction = L"New";
		}

		VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Copy study from: started (action = '%s')", (LPCWSTR)strAction));

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiSession> pSession;
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_SESSION_MGR,
			__uuidof(IVsiSession),
			(IUnknown**)&pSession);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
		{
			CComPtr<IVsiImage> pImage;
			pSession->GetImage(i, &pImage);
			if (NULL != pImage)
			{
				CComVariant vParam(pImage.p);
				hr = m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_IMAGE, &vParam);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				vParam = i;
				hr = m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_MODE_VIEW, &vParam);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

		// Create XML doc
		CComPtr<IXMLDOMDocument> pXmlDoc;
		hr = VsiCreateDOMDocument(&pXmlDoc);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Root element
		hr = pXmlDoc->createElement(CComBSTR(VSI_STUDY_XML_ELM_STUDIES), &pElmRoot);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pXmlDoc->appendChild(pElmRoot, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vJobType(VSI_JOB_STUDY_IMPORT);
		hr = pElmRoot->setAttribute(CComBSTR(VSI_STUDIES_XML_ATT_JOB), vJobType);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiEnumStudy> pEnumStudy;
		VSI_DATA_LIST_COLLECTION items = { 0 };
		items.ppEnumStudy = &pEnumStudy;

		hr = m_pDataList->GetItems(&items);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get New Study Access Default from new Owner
		bool bNewOwnerStudyAccess = false;
		CComPtr<IVsiOperator> pOperator;
		hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, m_strOwnerOperator, &pOperator);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		VSI_OPERATOR_DEFAULT_STUDY_ACCESS dwStudyAccess;
		hr = pOperator->GetDefaultStudyAccess(&dwStudyAccess);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);
		if (S_OK == hr)
		{
			bNewOwnerStudyAccess = true;
		}

		// Loop selected studies
		CComPtr<IXMLDOMElement> pStudyElement;
		CComPtr<IVsiStudy> pStudy;
		while (S_OK == pEnumStudy->Next(1, &pStudy, NULL))
		{
			CComPtr<IVsiStudy> pExistingStudy;

			// Set Overwrite attribute
			bOverwrite = FALSE;

			BOOL bNewOwner = FALSE;
			BOOL bNewCopy = FALSE;
			BOOL bSkip = FALSE;
			CComPtr<IVsiStudy> pStudyLocal;

			// New owner?
			{
				CComVariant vOwner;
				hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					bNewOwner = m_strOwnerOperator != V_BSTR(&vOwner);
				}
				else
				{
					// No owner assigned (it is an "open" quick study with no operator assigned)
					bNewOwner = TRUE;
				}
			}

			vId.Clear();
			hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Try VSI_PROP_STUDY_ID_COPIED 1st
			hr = m_pStudyMgr->GetIsStudyPresent(
				VSI_PROP_STUDY_ID_COPIED,
				V_BSTR(&vId),
				m_strOwnerOperator,
				&pExistingStudy);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK != hr)
			{
				// Try VSI_PROP_STUDY_ID 2nd
				hr = m_pStudyMgr->GetIsStudyPresent(
					VSI_PROP_STUDY_ID,
					V_BSTR(&vId),
					m_strOwnerOperator,
					&pExistingStudy);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			if (S_OK == hr)  // Study is present
			{
				vId.Clear();
				hr = pExistingStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				// Check if there are any series that need to be overwritten
				if (CVsiCopyStudyFromConfirmDlg::ACTION_OVERWRITE == iOverwriteAction)
				{
					bOverwrite = TRUE;
					bNewOwner = FALSE;

					// Get total lock state
					BOOL bLockAll = VsiGetBooleanValue<BOOL>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
						VSI_PARAMETER_SYS_LOCK_ALL, m_pPdm);
 
					// Check study lock state
					if (bLockAll)
					{
						hr = pExistingStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
 
						if (VARIANT_TRUE == V_BOOL(&vLocked))
						{
							bOverwrite = FALSE;
							bSkip = TRUE;
						}
					}
				}
				else if (CVsiCopyStudyFromConfirmDlg::ACTION_SKIP == iOverwriteAction)
				{
					int iNumNewSeriesToImport = 0;

					// Enumerate all the series in the source study
					CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
					hr = pPersist->LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST | VSI_DATA_TYPE_NO_SESSION_LINK, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiEnumSeries> pEnum;
					hr = pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_CREATED_DATE, FALSE, &pEnum, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiSeries> pSeries;
					while (pEnum->Next(1, &pSeries, NULL) == S_OK)
					{
						CComVariant vSeriesId;
						hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vSeriesId);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IUnknown> pUnkSeries;
						hr = pExistingStudy->GetItem(V_BSTR(&vSeriesId), &pUnkSeries);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (S_OK != hr)
						{
							++iNumNewSeriesToImport;
						}

						pSeries.Release();
					}

					// Skip importing any study that doesn't have any new series to copy over
					if (0 == iNumNewSeriesToImport)
					{
						pStudy.Release();
						continue;
					}

					bSkip = TRUE;
					bNewOwner = FALSE;
				}
				else if (CVsiCopyStudyFromConfirmDlg::ACTION_COPY == iOverwriteAction)
				{
					bOverwrite = FALSE;
					bNewCopy = TRUE;
				}
				else
				{
					bSkip = TRUE;
				}
			}

			// Create an element for the new study
			pStudyElement.Release();
			hr = VsiCreateStudyElement(
				pXmlDoc,
				pElmRoot,
				pStudy,
				&pStudyElement,
				m_strOwnerOperator,
				TRUE,
				FALSE);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (bNewOwner || bNewCopy)
			{
				// Assigns new id
				WCHAR szStudyId[128];
				VsiGetGuid(szStudyId, _countof(szStudyId));
				CComVariant vStudyId(szStudyId);
				hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_ID), vStudyId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (bNewCopy)
				{
					CComVariant vName;
					hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_NAME), &vName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vCreatedId;
					hr = pStudy->GetProperty(VSI_PROP_STUDY_ID_CREATED, &vCreatedId);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					int iNum(0);
					hr = m_pStudyMgr->GetStudyCopyNumber(V_BSTR(&vCreatedId), m_strOwnerOperator, &iNum);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CString strName;
					if (iNum != 0)
					{
						strName.Format(L"Copy of %s (%d)", V_BSTR(&vName), iNum);
						if (VSI_STUDY_NAME_MAX < strName.GetLength())
						{
							// Fallback to the old name
							strName = V_BSTR(&vName);
						}
					}
					else
					{
						strName = V_BSTR(&vName);
					}

					hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_NAME), CComVariant(strName));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
				else if (bNewOwner)
				{
					// A new copy because of a new owner
					hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_NEW_OWNER_COPY), CComVariant(true));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}
			else if (bOverwrite || bSkip)
			{
				_ASSERT(pExistingStudy != NULL);
				if (NULL == pExistingStudy)
				{
					// Something wrong - shouldn't happen
					pStudy.Release();
					continue;
				}

				// Set Target Id to reconstruct id's
				hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_ID), vId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Set Target to the folder name of the existing study
				pszPath.Free();
				hr = pExistingStudy->GetDataPath(&pszPath);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				pszSpt = wcsrchr(pszPath, L'\\');
				CString strFolderName(pszPath, (int)(pszSpt - pszPath));
				pszSpt = wcsrchr((LPWSTR)(LPCWSTR)strFolderName, L'\\');
				strFolderName = strFolderName.Mid((int)(pszSpt - (LPWSTR)(LPCWSTR)strFolderName) + 1);

				CComVariant vOverwrite(true);
				hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_OVERWRITE), vOverwrite);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				strFolderName += L".tmp";
				var = (LPCWSTR)strFolderName;

				hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_TARGET), var);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Set Target Locked attribute
				vLocked.Clear();
				hr = pExistingStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (VT_EMPTY == V_VT(&vLocked))
					vLocked = false;

				hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_LOCKED), vLocked);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			iTotalSizeStudy = 0;

			CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
			hr = pPersist->LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST | VSI_DATA_TYPE_NO_SESSION_LINK, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IVsiEnumSeries> pEnum;
			hr = pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_CREATED_DATE, FALSE, &pEnum, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr)
			{
				BOOL bSeriesToOverwrite = FALSE;

				CComPtr<IVsiSeries> pSeries;
				while (pEnum->Next(1, &pSeries, NULL) == S_OK)
				{
					bSeriesToOverwrite = FALSE;

					pszPath.Free();
					hr = pSeries->GetDataPath(&pszPath);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					pszSpt = wcsrchr(pszPath, L'\\');
					CString strPath(pszPath, (int)(pszSpt - pszPath));

					// Check if the series is in the exclude set
					auto iter = m_setRemovedSeries.find(strPath);

					if (iter == m_setRemovedSeries.end())  // Is not in the exclude set
					{
						// Process overwritten series separately
						if (bOverwrite || bSkip)
						{
							// Check if series is going to be overwritten
							if (NULL != pExistingStudy)
							{
								CComVariant vSeriesId;
								hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vSeriesId);
								VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

								CComPtr<IUnknown> pUnkSeries;
								hr = pExistingStudy->GetItem(V_BSTR(&vSeriesId), &pUnkSeries);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								if (S_OK == hr)
								{
									bSeriesToOverwrite = TRUE;
								}
							}

							if (bSeriesToOverwrite && bSkip)
							{
								pSeries.Release();
								continue;
							}
						}

						CComPtr<IXMLDOMElement> pImportSeriesElement;
						hr = VsiCreateSeriesElement(pXmlDoc, pStudyElement, pSeries, &pImportSeriesElement, &iTotalSizeSeries);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (bSeriesToOverwrite)
						{
							CComVariant vOverwrite(true);
							hr = pImportSeriesElement->setAttribute(CComBSTR(VSI_SERIES_XML_ATT_OVERWRITE), vOverwrite);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}

						iTotalSizeStudy += iTotalSizeSeries;
					}

					pSeries.Release();
				}
			}

			iTotalSizeToCopy += iTotalSizeStudy;
			pStudy.Release();
		}

#ifdef _DEBUG
{
		WCHAR szPath[MAX_PATH];
		GetModuleFileName(NULL, szPath, _countof(szPath));
		LPWSTR pszSpt = wcsrchr(szPath, L'\\');
		lstrcpy(pszSpt + 1, VSI_FOLDER_DEBUG);
		lstrcat(szPath, L"\\copyfrom1.xml");
		pXmlDoc->save(CComVariant(szPath));
}
#endif

		// Create StudyMover object
		CComPtr<IVsiStudyMover> pStudyMover;
		hr = pStudyMover.CoCreateInstance(__uuidof(VsiStudyMover));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant v(pXmlDoc);
		hr = pStudyMover->Initialize(m_hWnd, &v, iTotalSizeToCopy, VSI_STUDY_MOVER_OVERWRITE);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get data path
		CString strPathData = VsiGetRangeValue<CString>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_PATH_DATA, m_pPdm);

		hr = pStudyMover->Import(strPathData);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		int iCopied(0);

		// Assigns new study id; Changes owner; Handles overwrite
		{
			CComPtr<IXMLDOMNodeList> pListStudy;
			hr = pXmlDoc->selectNodes(
				CComBSTR(L"//" VSI_STUDY_XML_ELM_STUDIES L"/" VSI_STUDY_XML_ELM_STUDY),
				&pListStudy);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting studies");

			long lLength = 0;
			hr = pListStudy->get_length(&lLength);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

			CComVariant vCopyStatus;
			CComVariant vTargetName;
			for (long l = 0; l < lLength; l++)
			{
				CComPtr<IXMLDOMNode> pNode;
				hr = pListStudy->get_item(l, &pNode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

				CComQIPtr<IXMLDOMElement> pElemStudy(pNode);
				VSI_CHECK_INTERFACE(pElemStudy, VSI_LOG_ERROR, NULL);

				vCopyStatus.Clear();
				hr = pElemStudy->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS), &vCopyStatus);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				vCopyStatus.ChangeType(VT_UI4);
				if (S_OK != hr || VSI_SEI_COPY_OK != V_UI4(&vCopyStatus))
				{
					// Copy failed - skip
					continue;
				}

				++iCopied;

				// Checks overwrite
				CString strStudyFolderDel;
				BOOL bOverwrt = FALSE;
				CComVariant vOverwrite;
				hr = pElemStudy->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_OVERWRITE), &vOverwrite);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				if (S_OK == hr)
				{
					vOverwrite.ChangeType(VT_BOOL);
					if (VARIANT_FALSE != V_BOOL(&vOverwrite))
					{
						// Overwrites old study
						bOverwrt = TRUE;
					}
				}

				// Folder
				vTargetName.Clear();
				hr = pElemStudy->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_TARGET), &vTargetName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CString strStudyFolderTmp;
				strStudyFolderTmp.Format(L"%s\\%s", strPathData.GetString(), V_BSTR(&vTargetName));

				CString strStudyFolder(strStudyFolderTmp, strStudyFolderTmp.ReverseFind(L'.'));

				// Loads properties
				CComPtr<IVsiStudy> pStudy;
				hr = m_pStudyMgr->LoadStudy(strStudyFolderTmp, TRUE, &pStudy);
				if (FAILED(hr))
				{
					vCopyStatus = VSI_SEI_COPY_FAILED;
					hr = pElemStudy->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS), vCopyStatus);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);
					continue;
				}

				// Checks owner
				{
					CComVariant vOwner;
					hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (0 != _wcsicmp(V_BSTR(&vOwner), m_strOwnerOperator))
					{
						// Sets owner
						vOwner = (LPCWSTR)m_strOwnerOperator;
						hr = pStudy->SetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}

				// Assigns study instance id
				{
					CComVariant vIdOrg;
					hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vIdOrg);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vId;
					hr = pElemStudy->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_ID), &vId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (vIdOrg != vId)
					{
						// This is a new copy
						hr = pStudy->SetProperty(VSI_PROP_STUDY_ID, &vId);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						// Is this a new copy because of a new owner?
						CComVariant vNewOwnerCopy(false);
						hr = pElemStudy->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_NEW_OWNER_COPY), &vNewOwnerCopy);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (S_OK == hr)
						{
							hr = vNewOwnerCopy.ChangeType(VT_BOOL);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
						else
						{
							vNewOwnerCopy = false;
						}

						if (VARIANT_FALSE != V_BOOL(&vNewOwnerCopy))
						{
							// New copy because of a new owner - Set VSI_PROP_STUDY_ID_COPIED
							hr = pStudy->SetProperty(VSI_PROP_STUDY_ID_COPIED, &vIdOrg);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
						else
						{
							// New copy requested by the user - Clear VSI_PROP_STUDY_ID_COPIED
							hr = pStudy->SetProperty(VSI_PROP_STUDY_ID_COPIED, NULL);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
					}
				}

				// Assigns study name
				{
					CComVariant vNameProp;
					hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vNameProp);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vName;
					hr = pElemStudy->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_NAME), &vName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (vNameProp != vName)
					{
						hr = pStudy->SetProperty(VSI_PROP_STUDY_NAME, &vName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}

				// Reset Study Access
				if (bNewOwnerStudyAccess)
				{
					CComVariant vStudyAccess;
					switch (dwStudyAccess)
					{
					case VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE:
						vStudyAccess = VSI_STUDY_ACCESS_PRIVATE;
						break;
			
					case VSI_OPERATOR_DEFAULT_STUDY_ACCESS_GROUP:
						vStudyAccess = VSI_STUDY_ACCESS_GROUP;
						break;

					case VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PUBLIC:
						vStudyAccess = VSI_STUDY_ACCESS_PUBLIC;
						break;
					}

					if (VT_BSTR == V_VT(&vStudyAccess))
					{
						hr = pStudy->SetProperty(VSI_PROP_STUDY_ACCESS, &vStudyAccess);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}

				// Resaves study file
				CComQIPtr<IVsiPersistStudy> pStudyPersist(pStudy);
				hr = pStudyPersist->SaveStudyData(NULL, 0);
				if (FAILED(hr))
				{
					vCopyStatus = VSI_SEI_COPY_FAILED;
					hr = pElemStudy->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS), vCopyStatus);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);
					continue;
				}

				strStudyFolderDel.Empty();

				// Checks overwrite
				if (!bOverwrt)
				{
					// Hide the original study from the system by rename to ".del" extension
					if (0 == _waccess(strStudyFolder, 0))
					{
						strStudyFolderDel = strStudyFolder + L".del";

						BOOL bRet = MoveFile(strStudyFolder, strStudyFolderDel);
						if (!bRet)
						{
							vCopyStatus = VSI_SEI_COPY_FAILED;
							hr = pElemStudy->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS), vCopyStatus);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

							// Remove Temp study folder
							VsiRemoveDirectory(strStudyFolderTmp);

							continue;
						}
					}

					// Activates new study by removing ".tmp" extension
					{
						BOOL bRet = TRUE;

						// Try 2 sec incase the files are still lock
						for (int i = 0; i < 5; ++i)
						{
							bRet = MoveFile(strStudyFolderTmp, strStudyFolder);
							if (bRet)
							{
								break;
							}

							Sleep(400);
						}

						if (!bRet)
						{
							vCopyStatus = VSI_SEI_COPY_FAILED;
							hr = pElemStudy->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS), vCopyStatus);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);
							continue;
						}
					}
				}
				else
				{
					strStudyFolderDel = strStudyFolderTmp;
				}

				if (!strStudyFolderDel.IsEmpty())
				{
					// Delete the ".del" folder
					BOOL bRet = VsiRemoveDirectory(strStudyFolderDel);
					if (!bRet)
					{
						VSI_LOG_MSG(
							VSI_LOG_WARNING,
							VsiFormatMsg(
								L"Failed to delete overwritten study folder - %s",
								(LPCWSTR)strStudyFolderDel));
					}
				}
			}
		}

		// Display status dialog
		CVsiReportExportStatus statusReport(
			V_UNKNOWN(&v),
			CVsiReportExportStatus::REPORT_TYPE_STUDY_IMPORT);
		statusReport.DoModal(m_hWnd);

#ifdef _DEBUG
{
		WCHAR szPath[MAX_PATH];
		GetModuleFileName(NULL, szPath, _countof(szPath));
		LPWSTR pszSpt = wcsrchr(szPath, L'\\');
		lstrcpy(pszSpt + 1, VSI_FOLDER_DEBUG);
		lstrcat(szPath, L"\\copyfrom2.xml");
		pXmlDoc->save(CComVariant(szPath));
}
#endif

		// Refresh UI
		if (iCopied > 0)
		{
			// XML for UI update
			CComPtr<IXMLDOMDocument> pUpdateXmlDoc;
			hr = VsiCreateDOMDocument(&pUpdateXmlDoc);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating DOMDocument");

			// Root element
			CComPtr<IXMLDOMElement> pElmUpdate;
			hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_UPDATE), &pElmUpdate);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vResetSel(true);
			hr = pElmUpdate->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_RESET_SELECTION), vResetSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pUpdateXmlDoc->appendChild(pElmUpdate, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IXMLDOMElement> pElmAdd;
			hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_ADD), &pElmAdd);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pElmUpdate->appendChild(pElmAdd, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IXMLDOMElement> pElmRefresh;
			hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_REFRESH), &pElmRefresh);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pElmUpdate->appendChild(pElmRefresh, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// For each study
			CComPtr<IXMLDOMNodeList> pListStudy;
			hr = pXmlDoc->selectNodes(
				CComBSTR(L"//" VSI_STUDY_XML_ELM_STUDIES L"/" VSI_STUDY_XML_ELM_STUDY),
				&pListStudy);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting studies");

			long lLength = 0;
			hr = pListStudy->get_length(&lLength);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

			CComVariant vCopyStatus;
			CComVariant vTargetName;
			for (long l = 0; l < lLength; l++)
			{
				CComPtr<IXMLDOMNode> pNode;
				hr = pListStudy->get_item(l, &pNode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

				CComQIPtr<IXMLDOMElement> pElemSeries(pNode);

				vCopyStatus.Clear();
				hr = pElemSeries->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS), &vCopyStatus);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));

				if (S_OK != hr)
				{
					// Not process at all
					continue;
				}

				vCopyStatus.ChangeType(VT_UI4);
				if (VSI_SEI_COPY_OK != V_UI4(&vCopyStatus))
				{
					// Copy failed - skip
					continue;
				}

				// Checks overwrite
				BOOL bOverwrite = FALSE;
				CComVariant vOverwrite;
				hr = pElemSeries->getAttribute(CComBSTR(VSI_SERIES_XML_ATT_OVERWRITE), &vOverwrite);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_OVERWRITE));
				if (S_OK == hr)
				{
					vOverwrite.ChangeType(VT_BOOL);
					if (VARIANT_FALSE != V_BOOL(&vOverwrite))
					{
						// Overwrites old study
						bOverwrite = TRUE;
					}
				}

				CComVariant vId;
				hr = pElemSeries->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_ID), &vId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComPtr<IXMLDOMElement> pElmItem;
				hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_ITEM), &pElmItem);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), vId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vSelected(true);
				hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_SELECT), vSelected);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (bOverwrite)
				{
					// Collapse overwritten study to refresh it's view on the next open action
					CComVariant vCollapse(true);
					hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_COLLAPSE), vCollapse);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Refresh
					hr = pElmRefresh->appendChild(pElmItem, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
				else
				{
					// Add

					CComVariant vTargetName;
					hr = pElemSeries->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_TARGET), &vTargetName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_TARGET));

					CString strStudyFolderTmp;
					strStudyFolderTmp.Format(L"%s\\%s", strPathData.GetString(), V_BSTR(&vTargetName));

					CString strStudyFolder(strStudyFolderTmp, strStudyFolderTmp.ReverseFind(L'.'));
					vTargetName = (LPCWSTR)strStudyFolder;

					hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_PATH), vTargetName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pElmAdd->appendChild(pElmItem, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}

#ifdef _DEBUG
{
			WCHAR szPath[MAX_PATH];
			GetModuleFileName(NULL, szPath, _countof(szPath));
			LPWSTR pszSpt = wcsrchr(szPath, L'\\');
			lstrcpy(pszSpt + 1, VSI_FOLDER_DEBUG);
			lstrcat(szPath, L"\\copyfromUpdate.xml");
			pUpdateXmlDoc->save(CComVariant(szPath));
}
#endif

			// Updates Study Manager
			hr = m_pStudyMgr->Update(pUpdateXmlDoc);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Close
		m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_COPY_STUDY_FROM, NULL);
	}
	VSI_CATCH(hr);

	VSI_LOG_MSG(VSI_LOG_INFO, L"Copy study from: ended");

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Copy Study From - Cancel");

	m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_COPY_STUDY_FROM, NULL);

	return 0;
}

int CVsiCopyStudyFrom::GetNumberOfSelectedItems()
{
	return (int)m_setSelectedStudy.size();
}

BOOL CVsiCopyStudyFrom::GetAvailablePath(LPWSTR pszBuffer, int iBufferLength)
{
	HRESULT hr = S_OK;
	CComVariant vPath;

	try
	{
		// Get data path
		CVsiRange<LPCWSTR> pParamPathData;
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_PATH_DATA, &pParamPathData);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiParameterRange> pRange(pParamPathData.m_pParam);
		VSI_CHECK_INTERFACE(pRange, VSI_LOG_ERROR, NULL);

		hr = pRange->GetValue(&vPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return ((S_OK == hr) && (0 == wcscpy_s(pszBuffer, iBufferLength, (LPWSTR)V_BSTR(&vPath))));
}

int CVsiCopyStudyFrom::CheckForOverwrite()
{
	HRESULT hr = S_OK;
	int iOverwrite = 0;

	try
	{
		// If operator is not selected - no overwrite checking possible
		if (m_strOwnerOperator.IsEmpty())
		{
			return 0;
		}

		CComPtr<IVsiEnumStudy> pEnumStudy;
		VSI_DATA_LIST_COLLECTION items = { 0 };
		items.ppEnumStudy = &pEnumStudy;

		hr = m_pDataList->GetItems(&items);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiStudy> pStudy;
		while (pEnumStudy->Next(1, &pStudy, NULL) == S_OK)
		{
			CComPtr<IVsiStudy> pStudyLocal;

			CComVariant vStudyId;
			hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vStudyId);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			// Try VSI_PROP_STUDY_ID_COPIED first
			hr = m_pStudyMgr->GetIsStudyPresent(
				VSI_PROP_STUDY_ID_COPIED,
				V_BSTR(&vStudyId),
				m_strOwnerOperator,
				&pStudyLocal);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK != hr)
			{
				// Try VSI_PROP_STUDY_ID next
				hr = m_pStudyMgr->GetIsStudyPresent(
					VSI_PROP_STUDY_ID,
					V_BSTR(&vStudyId),
					m_strOwnerOperator,
					&pStudyLocal);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			if (S_OK == hr)  // Study is present
			{
				// Check if there are any series that need to be overwritten

				// Enumerate all the series in the source study
				CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
				hr = pPersist->LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST | VSI_DATA_TYPE_NO_SESSION_LINK, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComPtr<IVsiEnumSeries> pEnum;
				hr = pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_CREATED_DATE, FALSE, &pEnum, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					CComPtr<IVsiSeries> pSeries;
					while (pEnum->Next(1, &pSeries, NULL) == S_OK)
					{
						CComVariant vSeriesId;
						hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vSeriesId);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IUnknown> pUnkSeries;
						hr = pStudyLocal->GetItem(V_BSTR(&vSeriesId), &pUnkSeries);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (S_OK == hr)
						{
							++iOverwrite;
						}

						pSeries.Release();
					}
				}
			}

			pStudy.Release();
		}
	}
	VSI_CATCH(hr);

	return iOverwrite;
}

void CVsiCopyStudyFrom::UpdateUI()
{
	HRESULT hr = S_OK;
	BOOL bEnableOK = TRUE;

	try
	{
		CVsiImportExportUIHelper<CVsiCopyStudyFrom>::UpdateUI();

		// Owner
		m_strOwnerOperator.Empty();

		int iIndex = (int)m_wndOperator.SendMessage(CB_GETCURSEL);
		if (CB_ERR != iIndex)
		{
			LPCWSTR pszOperator = (LPCWSTR)m_wndOperator.SendMessage(CB_GETITEMDATA, iIndex);
			if (! IS_INTRESOURCE(pszOperator))
			{
				m_wndOperator.GetWindowText(m_strOwnerOperator);
			}
		}

		if (m_strOwnerOperator.IsEmpty())
		{
			bEnableOK = FALSE;
		}
		if (0 == m_setSelectedStudy.size())
		{
			bEnableOK = FALSE;
		}

		// Check for sufficient space in the current location. If there is not enough
		// then enable the message and show them how much
		UpdateSizeUI();

		GetDlgItem(IDOK).EnableWindow(bEnableOK);

		int iStudyPreFiltered = 0, iSeriesPreFiltered = 0;
		VSI_DATA_LIST_COLLECTION selection;
		memset(&selection, 0, sizeof(VSI_DATA_LIST_COLLECTION));
		selection.dwFlags = VSI_DATA_LIST_FILTER_NONE;
		selection.piStudyCount = &iStudyPreFiltered;
		selection.piSeriesCount = &iSeriesPreFiltered;
		hr = m_pDataList->GetSelection(&selection);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		GetDlgItem(IDC_CSF_REMOVE).EnableWindow((iStudyPreFiltered + iSeriesPreFiltered) > 0);
	}
	VSI_CATCH(hr);
}

LRESULT CVsiCopyStudyFrom::OnTreeClick(int, LPNMHDR pnmh, BOOL &)
{
	DWORD dwPos = GetMessagePos();
	TVHITTESTINFO tviti = { 0 };
	tviti.pt.x = GET_X_LPARAM(dwPos);
	tviti.pt.y = GET_Y_LPARAM(dwPos);
	::ScreenToClient(pnmh->hwndFrom, &tviti.pt);

	TreeView_HitTest(pnmh->hwndFrom, &tviti);

	if ((TVHT_ONITEMSTATEICON == tviti.flags) || (TVHT_ONITEMRIGHT == tviti.flags))
	{
		TreeView_SelectItem(pnmh->hwndFrom, tviti.hItem);

		int iState = TreeView_GetCheckState(pnmh->hwndFrom, tviti.hItem);
		switch (iState)
		{
		case VSI_CHECK_STATE_UNCHECKED:
			RemoveSelectedItems();
			break;
		case VSI_CHECK_STATE_CHECKED:
			AddSelectedItems();
			break;
		}
	}

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnTreeSelChanged(int, LPNMHDR, BOOL &bHandled)
{
	if (m_pFExplorer != NULL)
	{
		CComHeapPtr<OLECHAR> pszFolderName;
		HRESULT hr = m_pFExplorer->GetSelectedFilePath(&pszFolderName);
		if (S_FALSE == hr)
			return hr;

		PathSetDlgItemPath(*this, IDC_IE_PATH, pszFolderName);

		UpdateFEStudyDetails();

		UpdateUI();

		bHandled = TRUE;
	}

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnTreeGetInfoTip(int, LPNMHDR pnmhr, BOOL &bHandled)
{
	HRESULT hr = S_OK;
	bHandled = TRUE;

	try
	{
		LPNMTVGETINFOTIP pgit = (LPNMTVGETINFOTIP)pnmhr;

		CComHeapPtr<OLECHAR> pszFolderPath;
		hr = m_pFExplorer->GetItemPath((ULONG_PTR)pgit->hItem, &pszFolderPath);
		// It may not be a physical folder
		if (S_OK == hr)
		{
			WCHAR szFileName[MAX_PATH];

			// Study
			_snwprintf_s(szFileName, _countof(szFileName),
				L"%s\\%s", (LPCWSTR)pszFolderPath, CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));

			if (0 == _waccess(szFileName, 0))
			{
				CComPtr<IVsiStudy> pStudy;
				hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
				hr = pPersist->LoadStudyData(szFileName, 0, NULL);
				if (SUCCEEDED(hr))
				{
					CString strTips;
					CString strItem;

					CComVariant v;
					hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &v);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CString strTextStudy(MAKEINTRESOURCE(IDS_CPYSTY_STUDY_NAME));
					strItem.Format(strTextStudy, V_BSTR(&v));
					strTips += strItem;

					v.Clear();
					hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &v);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					strItem.Format(L"Owner:\t\t%s\r\n", V_BSTR(&v));
					strTips += strItem;

					// Created date
					WCHAR szTempBuffer[256];
					SYSTEMTIME stDate;
					CComVariant vDate;
					hr = pStudy->GetProperty(VSI_PROP_STUDY_CREATED_DATE, &vDate);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					COleDateTime date(vDate);
					date.GetAsSystemTime(stDate);

					SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

					GetDateFormat(
						LOCALE_USER_DEFAULT,
						DATE_SHORTDATE,
						&stDate,
						NULL,
						szTempBuffer,
						_countof(szTempBuffer));
					strItem.Format(L"Date:\t\t%s\r\n", szTempBuffer);
					strTips += strItem;

					lstrcpyn(pgit->pszText, strTips, pgit->cchTextMax);
				}
			}

			_snwprintf_s(szFileName, _countof(szFileName),
				L"%s\\%s", (LPCWSTR)pszFolderPath, CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME)));

			if (0 == _waccess(szFileName, 0))
			{
				CComPtr<IVsiSeries> pSeries;
				hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComQIPtr<IVsiPersistSeries> pPersist(pSeries);
				hr = pPersist->LoadSeriesData(szFileName, 0, NULL);
				if (SUCCEEDED(hr))
				{
					CString strTips;
					CString strItem;

					CComVariant v;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_NAME, &v);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CString strTextSeries(MAKEINTRESOURCE(IDS_CPYSTY_SERIES_NAME));
					strItem.Format(strTextSeries, V_BSTR(&v));
					strTips += strItem;

					v.Clear();
					hr = pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &v);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					strItem.Format(L"Acquired By:\t%s\r\n", V_BSTR(&v));
					strTips += strItem;

					// Created date
					WCHAR szTempBuffer[256];
					SYSTEMTIME stDate;
					CComVariant vDate;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_CREATED_DATE, &vDate);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					COleDateTime date(vDate);
					date.GetAsSystemTime(stDate);

					SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

					GetDateFormat(
						LOCALE_USER_DEFAULT,
						DATE_SHORTDATE,
						&stDate,
						NULL,
						szTempBuffer,
						_countof(szTempBuffer));
					strItem.Format(L"Date:\t\t%s\r\n", szTempBuffer);
					strTips += strItem;

					lstrcpyn(pgit->pszText, strTips, pgit->cchTextMax);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnCheckForStudySeries(int, LPNMHDR pnmh, BOOL &)
{
	HRESULT hr = S_OK;

	try
	{
		NM_FOLDER_GETITEMINFO* pInfo = (NM_FOLDER_GETITEMINFO*)pnmh;
		pInfo->pItem->mask |= TVIF_STATE;
		pInfo->pItem->stateMask = TVIS_STATEIMAGEMASK;
		pInfo->pItem->state = 0;

		// Skip directory root to improve performance (e.g. testing A:\ is very slow)
		// Study and series XML file should not be under directory root
		if (!PathIsRoot(pInfo->pszFolderPath))
		{
			// Study
			CString strStudyFileName;
			strStudyFileName.Format(L"%s\\%s", pInfo->pszFolderPath, CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));

			// Series
			CString strSeriesFileName;
			strSeriesFileName.Format(L"%s\\%s", pInfo->pszFolderPath, CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME)));

			if (0 == _waccess(strStudyFileName, 0))
			{
				int iCheck(VSI_CHECK_STATE_UNCHECKED);

				if (IsStudyAdded(pInfo->pszFolderPath))
				{
					if (IsStudyFullyAdded(pInfo->pszFolderPath))
					{
						iCheck = VSI_CHECK_STATE_CHECKED;
					}
					else
					{
						iCheck = VSI_CHECK_STATE_PARTIAL;
					}
				}

				LONG lStudyFolderImageIndex = 0;
				hr = m_pFExplorer->GetVevoImageIndex(&lStudyFolderImageIndex);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				pInfo->pItem->iImage = lStudyFolderImageIndex;
				pInfo->pItem->iSelectedImage = lStudyFolderImageIndex;
				pInfo->pItem->state = INDEXTOSTATEIMAGEMASK(iCheck + 1);
			}
			else if (0 == _waccess(strSeriesFileName, 0))
			{
				// Get the true series name
				CComPtr<IVsiSeries> pSeries;
				hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComQIPtr<IVsiPersistSeries> pPersist(pSeries);
				hr = pPersist->LoadSeriesData(strSeriesFileName, 0, NULL);
				if (SUCCEEDED(hr))
				{
					int iCheck(VSI_CHECK_STATE_UNCHECKED);

					if (IsSeriesAdded(pInfo->pszFolderPath))
					{
						iCheck = VSI_CHECK_STATE_CHECKED;
					}

					LONG lStudyFolderImageIndex = 0;
					hr = m_pFExplorer->GetVevoImageIndex(&lStudyFolderImageIndex);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					pInfo->pItem->iImage = lStudyFolderImageIndex;
					pInfo->pItem->iSelectedImage = lStudyFolderImageIndex;
					pInfo->pItem->state = INDEXTOSTATEIMAGEMASK(iCheck + 1);

					CComVariant vSeriesName;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_NAME, &vSeriesName);
					if (S_OK == hr)
					{
						wcscpy_s(pInfo->pszName, pInfo->iNameMax, V_BSTR(&vSeriesName));
					}
				}
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnItemChanged(int, LPNMHDR, BOOL &bHandled)
{
	MSG msg;
	PeekMessage(&msg, *this, WM_VSI_UPDATE_PREVIEW, WM_VSI_UPDATE_PREVIEW, PM_REMOVE);
	PostMessage(WM_VSI_UPDATE_PREVIEW);

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnCheckItemStatus(int, LPNMHDR pnmh, BOOL &)
{
	NM_DL_ITEM* pCheck = (NM_DL_ITEM*)pnmh;

	// Here only series is of a concern
	if (pCheck->type != VSI_DATA_LIST_SERIES)
		return VSI_DATA_LIST_ITEM_STATUS_NORMAL;

	LPCWSTR pszSpt = wcsrchr(pCheck->pszItemPath, L'\\');
	CString strPath(pCheck->pszItemPath, (int)(pszSpt - pCheck->pszItemPath));

	auto iter = m_setRemovedSeries.find(strPath);
	return (iter == m_setRemovedSeries.end()) ?
		VSI_DATA_LIST_ITEM_STATUS_NORMAL :
		VSI_DATA_LIST_ITEM_STATUS_HIDDEN;
}

LRESULT CVsiCopyStudyFrom::OnSelChanged(int, LPNMHDR, BOOL &bHandled)
{
	CWindow wndFE(GetDlgItem(IDC_IE_FOLDER_TREE));
	CWindow wndSelectedStudies(GetDlgItem(IDC_CSF_STUDY_LIST));

	if ((HWND)wndFE == GetFocus())
	{
		UpdateFEStudyDetails();
	}
	else if ((HWND)wndSelectedStudies == GetFocus())
	{
		MSG msg;
		PeekMessage(&msg, *this, WM_VSI_UPDATE_PREVIEW, WM_VSI_UPDATE_PREVIEW, PM_REMOVE);
		PostMessage(WM_VSI_UPDATE_PREVIEW);
	}

	UpdateUI();

	bHandled = TRUE;

	return 0;
}

HRESULT CVsiCopyStudyFrom::AddSelectedItems()
{
	HRESULT hr = S_OK;
	HANDLE hFile(INVALID_HANDLE_VALUE);

	try
	{
		CComHeapPtr<OLECHAR> pszSelectedPath;
		hr = m_pFExplorer->GetSelectedFilePath(&pszSelectedPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CString strTmp(MAKEINTRESOURCE(IDS_CPYSTY_STUDY_CORRUPTED));
		if ((NULL != pszSelectedPath) && (0 != *pszSelectedPath))
		{
			CString strStudyTest;
			strStudyTest.Format(L"%s\\%s", pszSelectedPath, CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));
			CString strSeriesTest;
			strSeriesTest.Format(L"%s\\%s", pszSelectedPath, CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME)));
			if (0 == _waccess(strStudyTest, 0))
			{
				// Cache last used path
				{
					CComHeapPtr<OLECHAR> pszPathName;
					hr = m_pFExplorer->GetSelectedNamespacePath(&pszPathName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Remove study folder
					LPWSTR pszSlash = wcsrchr(pszPathName, L'\\');
					if (NULL != pszSlash)
					{
						CString strText(MAKEINTRESOURCE(IDS_CPYSTY_STUDYS_CORRUPTED));
						strTmp.Format(strText, pszSlash + 1);
						*pszSlash = 0;
						m_strLastAddPath = pszPathName;
					}
				}

				CString strSelectedPath((LPCWSTR)pszSelectedPath);

				// Check whether selected path has already been copied
				auto iter = m_setSelectedStudy.find(strSelectedPath);
				if (iter != m_setSelectedStudy.end())
				{
					// Updates m_setRemovedSeries
					for (iter = m_setRemovedSeries.begin(); iter != m_setRemovedSeries.end();)
					{
						const CString &strSeriesPath = *iter;

						if (0 == strSeriesPath.Find(pszSelectedPath))
						{
							iter = m_setRemovedSeries.erase(iter);
						}
						else
						{
							++iter;
						}
					}
				}
				else
				{
					CComPtr<IVsiStudy> pStudy;
					hr = m_pStudyMgr->LoadStudy(pszSelectedPath, TRUE, &pStudy);
					if (SUCCEEDED(hr))
					{
						CComVariant vId;
						hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
						if (SUCCEEDED(hr))
						{
							auto iterId = m_setSelectedStudyId.find(V_BSTR(&vId));
							if (iterId == m_setSelectedStudyId.end())
							{
								// Not added before
								m_setSelectedStudy.insert(strSelectedPath);
								m_setSelectedStudyId.insert(V_BSTR(&vId));

								hr = m_pDataList->AddItem(
									pStudy,
									VSI_DATA_LIST_SELECTION_CLEAR |
										VSI_DATA_LIST_SELECTION_UPDATE |
										VSI_DATA_LIST_SELECTION_ENSURE_VISIBLE);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							}
							else
							{
								// Study added already
								VsiMessageBox(
									GetTopLevelParent(),
									CString(MAKEINTRESOURCE(IDS_CPYSTY_PREVENT_OVERWRITE)),
									CString(MAKEINTRESOURCE(IDS_CPYSTY_COPY_STUDY_FROM)),
									MB_OK | MB_ICONWARNING);

								m_pFExplorer->Refresh();
							}
						}
					}

					if (FAILED(hr))
					{
						VsiMessageBox(
							GetTopLevelParent(),
							strTmp,
							CString(MAKEINTRESOURCE(IDS_CPYSTY_COPY_STUDY_FROM)),
							MB_OK | MB_ICONWARNING);

						m_pFExplorer->Refresh();
					}
				}

				m_pDataList->UpdateItem(NULL, TRUE);
			}

			if (0 == _waccess(strSeriesTest, 0))
			{
				// Cache last used path
				{
					CComHeapPtr<OLECHAR> pszPathName;
					hr = m_pFExplorer->GetSelectedNamespacePath(&pszPathName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Remove study folder
					LPWSTR pszSlash = wcsrchr(pszPathName, L'\\');
					if (NULL != pszSlash)
					{
						CString strText(MAKEINTRESOURCE(IDS_CPYSTY_STUDYS_CORRUPTED));
						strTmp.Format(strText, pszSlash + 1);
						*pszSlash = 0;
						pszSlash = wcsrchr(pszPathName, L'\\');
						if (NULL != pszSlash)
						{
							*pszSlash = 0;
							m_strLastAddPath = pszPathName;
						}
					}
				}

				CString strSelectedPath((LPCWSTR)pszSelectedPath);

				// Check whether selected path has already been copied
				LPWSTR pszSlash = wcsrchr(pszSelectedPath, L'\\');
				CString strStudyPath(pszSelectedPath, (int)(pszSlash - pszSelectedPath));

				auto iter = m_setSelectedStudy.find(strStudyPath);
				if (iter != m_setSelectedStudy.end())
				{
					// Updates m_setRemovedSeries
					iter = m_setRemovedSeries.find(strSelectedPath);
					if (iter != m_setRemovedSeries.end())
					{
						m_setRemovedSeries.erase(iter);
					}
				}
				else
				{
					CComPtr<IVsiStudy> pStudy;
					hr = m_pStudyMgr->LoadStudy(strStudyPath, TRUE, &pStudy);
					if (SUCCEEDED(hr))
					{
						CComVariant vId;
						hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
						if (SUCCEEDED(hr))
						{
							auto iterId = m_setSelectedStudyId.find(V_BSTR(&vId));
							if (iterId == m_setSelectedStudyId.end())
							{
								// New study
								CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
								hr = pPersist->LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST | VSI_DATA_TYPE_NO_SESSION_LINK, NULL);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								hr = m_pDataList->AddItem(
									pStudy,
									VSI_DATA_LIST_SELECTION_CLEAR |
										VSI_DATA_LIST_SELECTION_UPDATE |
										VSI_DATA_LIST_SELECTION_ENSURE_VISIBLE);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								m_setSelectedStudy.insert(strStudyPath);
								m_setSelectedStudyId.insert(V_BSTR(&vId));

								// Remove all series but one
								CComPtr<IVsiEnumSeries> pEnum;
								hr = pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_CREATED_DATE, FALSE, &pEnum, NULL);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								CComPtr<IVsiSeries> pSeries;
								while (S_OK == pEnum->Next(1, &pSeries, NULL))
								{
									CComHeapPtr<OLECHAR> pszPath;
									hr = pSeries->GetDataPath(&pszPath);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									LPWSTR pszSpt = wcsrchr(pszPath, L'\\');
									CString strSeriesPath(pszPath, (int)(pszSpt - pszPath));

									if (strSeriesPath != strSelectedPath)
									{
										m_setRemovedSeries.insert(strSeriesPath);
									}

									pSeries.Release();
								}
							}
							else
							{
								// Study added already
								VsiMessageBox(
									GetTopLevelParent(),
									CString(MAKEINTRESOURCE(IDS_CPYSTY_PREVENT_OVERWRITE)),
									CString(MAKEINTRESOURCE(IDS_CPYSTY_COPY_STUDY_FROM)),
									MB_OK | MB_ICONWARNING);

								m_pFExplorer->Refresh();
							}
						}
					}

					if (FAILED(hr))
					{
						VsiMessageBox(
							GetTopLevelParent(),
							strTmp,
							CString(MAKEINTRESOURCE(IDS_CPYSTY_COPY_STUDY_FROM)),
							MB_OK | MB_ICONWARNING);

						m_pFExplorer->Refresh();
					}
				}

				m_pDataList->UpdateItem(NULL, TRUE);
			}
			else
			{
				BOOL bAddStudy = TRUE;

				// Cache last used path
				{
					CComHeapPtr<OLECHAR> pszPathName;
					hr = m_pFExplorer->GetSelectedNamespacePath(&pszPathName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					m_strLastAddPath = pszPathName;
				}

				// Construct a string to represent the base path to use for
				// finding all of the studies at the selected full path
				CString strFolderName;
				WCHAR szSourceFolder[VSI_MAX_PATH];
				wcscpy_s(szSourceFolder, _countof(szSourceFolder), pszSelectedPath);
				wcscat_s(szSourceFolder, _countof(szSourceFolder), L"\\*.*");

				WIN32_FIND_DATA ff;
				hFile = FindFirstFile(szSourceFolder, &ff);
				BOOL bWorking = (INVALID_HANDLE_VALUE != hFile);

				VSI_DATA_LIST_SELECTION dwSelectFlags =
					VSI_DATA_LIST_SELECTION_CLEAR |
					VSI_DATA_LIST_SELECTION_UPDATE |
					VSI_DATA_LIST_SELECTION_ENSURE_VISIBLE;

				while (bWorking)  // For each file in the study.
				{
					// Skip the dots, include subdirectories.
					if (L'.' == ff.cFileName[0])
					{
						NULL;
					}
					else if (ff.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
					{
						// File copy
						strFolderName.Format(
							L"%s\\%s\\%s",
							pszSelectedPath,
							ff.cFileName,
							CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));
						if (0 == _waccess(strFolderName, 0))
						{
							bAddStudy = TRUE;

							strFolderName.Format(L"%s\\%s",	pszSelectedPath, ff.cFileName);

							// Check whether selected path has already included
							auto iter = m_setSelectedStudy.find(strFolderName);
							bAddStudy = iter == m_setSelectedStudy.end();

							// Add this study to the list
							if (bAddStudy)
							{
								CComPtr<IVsiStudy> pStudy;
								hr = m_pStudyMgr->LoadStudy(strFolderName, TRUE, &pStudy);
								if (SUCCEEDED(hr))
								{
									CComVariant vId;
									hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
									if (SUCCEEDED(hr))
									{
										auto iterId = m_setSelectedStudyId.find(V_BSTR(&vId));
										if (iterId == m_setSelectedStudyId.end())
										{
											// New study
											m_pDataList->AddItem((LPUNKNOWN)pStudy, dwSelectFlags);
											VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

											dwSelectFlags &= ~VSI_DATA_LIST_SELECTION_CLEAR;

											m_setSelectedStudy.insert(strFolderName);
											m_setSelectedStudyId.insert(V_BSTR(&vId));
										}
										else
										{
											// Already added
										}
									}
								}

								if (FAILED(hr))
								{
									CString strTmp;
									CString strText(MAKEINTRESOURCE(IDS_CPYSTY_STUDYS_CORRUPTED));
									strTmp.Format(strText, ff.cFileName);
									VsiMessageBox(
										GetTopLevelParent(),
										strTmp,
										CString(MAKEINTRESOURCE(IDS_CPYSTY_COPY_STUDY_FROM)),
										MB_OK | MB_ICONWARNING);

									m_pFExplorer->Refresh();
								}
							}
						}
					}

					bWorking = FindNextFile(hFile, &ff);
				}

				// Done with the current directory. Close the handle and return.
				FindClose(hFile);
				hFile = INVALID_HANDLE_VALUE;
			}
		}

		CalculateRequiredSize();
		UpdateAvailableSpace();
		UpdateUI();

		// Sets focus
		PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)GetDlgItem(IDC_IE_SELECTED), TRUE);
	}
	VSI_CATCH(hr);

	if (INVALID_HANDLE_VALUE != hFile)
	{
		FindClose(hFile);
	}

	return hr;
}

HRESULT CVsiCopyStudyFrom::RemoveSelectedItems()
{
	HRESULT hr(S_OK);
	HANDLE hFile(INVALID_HANDLE_VALUE);

	try
	{
		CComHeapPtr<OLECHAR> pszSelectedPath;
		hr = m_pFExplorer->GetSelectedFilePath(&pszSelectedPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if ((NULL != pszSelectedPath) && (0 != *pszSelectedPath))
		{
			CString strStudyTest;
			strStudyTest.Format(L"%s\\%s", pszSelectedPath, CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));
			CString strSeriesTest;
			strSeriesTest.Format(L"%s\\%s", pszSelectedPath, CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME)));
			if (0 == _waccess(strStudyTest, 0))
			{
				CString strSelectedPath(pszSelectedPath);

				CComPtr<IVsiStudy> pStudy;
				hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
				hr = pPersist->LoadStudyData(strStudyTest, 0, NULL);
				if (SUCCEEDED(hr))
				{
					CComVariant vNs;
					hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vNs);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vId;
					hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = m_pDataList->RemoveItem(V_BSTR(&vNs), VSI_DATA_LIST_SELECTION_UPDATE);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					auto iter = m_setSelectedStudy.find(strSelectedPath);
					if (iter != m_setSelectedStudy.end())
					{
						m_setSelectedStudy.erase(iter);
					}

					auto iterId = m_setSelectedStudyId.find(V_BSTR(&vId));
					if (iterId != m_setSelectedStudyId.end())
					{
						m_setSelectedStudyId.erase(iterId);
					}

					for (auto iter = m_setRemovedSeries.begin(); iter != m_setRemovedSeries.end(); )
					{
						const CString &strSeriesPath = *iter;

						if (0 == strSeriesPath.Find(strSelectedPath))
						{
							iter = m_setRemovedSeries.erase(iter);
						}
						else
						{
							++iter;
						}
					}
				}
			}
			else if (0 == _waccess(strSeriesTest, 0))
			{
				CString strSelectedPath(pszSelectedPath);

				// Get study folder
				LPCWSTR pszSlash = wcsrchr(pszSelectedPath, L'\\');
				if (NULL != pszSlash)
				{
					CString strStudyPath(pszSelectedPath, (int)(pszSlash - pszSelectedPath));

					strStudyTest.Format(L"%s\\%s", strStudyPath, CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));

					CComPtr<IVsiStudy> pStudy;
					hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
					hr = pPersist->LoadStudyData(strStudyTest, 0, NULL);
					if (SUCCEEDED(hr))
					{
						// Check whether all the series are excluded
						int iSeriesExcluded(0);
						int iSeriesTotal(0);
						CComPtr<IVsiEnumSeries> pEnum;
						hr = pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_CREATED_DATE, FALSE, &pEnum, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IVsiSeries> pSeries;
						while (pEnum->Next(1, &pSeries, NULL) == S_OK)
						{
							++iSeriesTotal;

							CComHeapPtr<OLECHAR> pszSeriesPath;
							hr = pSeries->GetDataPath(&pszSeriesPath);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							LPCWSTR pszSpt = wcsrchr(pszSeriesPath, L'\\');
							CString strSeriesPath(pszSeriesPath, (int)(pszSpt - pszSeriesPath));

							auto iter = m_setRemovedSeries.find(strSeriesPath);
							if (iter != m_setRemovedSeries.end())
							{
								++iSeriesExcluded;
							}

							pSeries.Release();
						}

						if (iSeriesTotal == iSeriesExcluded + 1)
						{
							// Excluding all series, remove the study
							CComVariant vNs;
							hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vNs);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							CComVariant vId;
							hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							hr = m_pDataList->RemoveItem(V_BSTR(&vNs), VSI_DATA_LIST_SELECTION_UPDATE);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							auto iter = m_setSelectedStudy.find(strStudyPath);
							if (iter != m_setSelectedStudy.end())
							{
								m_setSelectedStudy.erase(iter);
							}

							auto iterId = m_setSelectedStudyId.find(V_BSTR(&vId));
							if (iterId != m_setSelectedStudyId.end())
							{
								m_setSelectedStudyId.erase(iterId);
							}

							for (auto iter = m_setRemovedSeries.begin(); iter != m_setRemovedSeries.end(); )
							{
								const CString &strSeriesPath = *iter;

								if (0 == strSeriesPath.Find(strStudyPath))
								{
									iter = m_setRemovedSeries.erase(iter);
								}
								else
								{
									++iter;
								}
							}
						}
						else
						{
							m_setRemovedSeries.insert(strSelectedPath);
						}
					}
				}
			}
			else
			{
				// Construct a string to represent the base path to use for
				// finding all of the studies at the selected full path
				CString strStudyTest;
				WCHAR szSourceFolder[VSI_MAX_PATH];
				wcscpy_s(szSourceFolder, _countof(szSourceFolder), pszSelectedPath);
				wcscat_s(szSourceFolder, _countof(szSourceFolder), L"\\*.*");

				WIN32_FIND_DATA ff;
				hFile = FindFirstFile(szSourceFolder, &ff);
				BOOL bWorking = (INVALID_HANDLE_VALUE != hFile);

				while (bWorking)  // For each file in the study.
				{
					// Skip the dots, include subdirectories.
					if (L'.' == ff.cFileName[0])
					{
						NULL;
					}
					else if (ff.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
					{
						// File copy
						strStudyTest.Format(
							L"%s\\%s\\%s",
							pszSelectedPath,
							ff.cFileName,
							CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));
						if (0 == _waccess(strStudyTest, 0))
						{
							CString strSelectedPath;

							strSelectedPath.Format(L"%s\\%s", pszSelectedPath, ff.cFileName);

							// Check whether selected path has already included
							auto iter = m_setSelectedStudy.find(strSelectedPath);

							// Remove this study from the list
							if (iter != m_setSelectedStudy.end())
							{
								CComPtr<IVsiStudy> pStudy;
								hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
								hr = pPersist->LoadStudyData(strStudyTest, 0, NULL);
								if (SUCCEEDED(hr))
								{
									CComVariant vNs;
									hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vNs);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									CComVariant vId;
									hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									hr = m_pDataList->RemoveItem(V_BSTR(&vNs), VSI_DATA_LIST_SELECTION_UPDATE);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									auto iter = m_setSelectedStudy.find(strSelectedPath);
									if (iter != m_setSelectedStudy.end())
									{
										m_setSelectedStudy.erase(iter);
									}

									auto iterId = m_setSelectedStudyId.find(V_BSTR(&vId));
									if (iterId != m_setSelectedStudyId.end())
									{
										m_setSelectedStudyId.erase(iterId);
									}

									for (auto iter = m_setRemovedSeries.begin(); iter != m_setRemovedSeries.end(); )
									{
										const CString &strSeriesPath = *iter;

										if (0 == strSeriesPath.Find(strSelectedPath))
										{
											iter = m_setRemovedSeries.erase(iter);
										}
										else
										{
											++iter;
										}
									}
								}
							}
						}
					}

					bWorking = FindNextFile(hFile, &ff);
				}

				// Done with the current directory. Close the handle and return.
				FindClose(hFile);
				hFile = INVALID_HANDLE_VALUE;
			}
		}

		hr = m_pDataList->UpdateItem(NULL, TRUE);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CalculateRequiredSize();
		UpdateAvailableSpace();
		UpdateUI();
	}
	VSI_CATCH(hr);

	if (INVALID_HANDLE_VALUE != hFile)
	{
		FindClose(hFile);
	}

	return hr;
}

BOOL CVsiCopyStudyFrom::IsFolderHasStudy()
{
	HRESULT hr = S_OK;
	HANDLE hFile = INVALID_HANDLE_VALUE;
	BOOL bHasStudy = FALSE;

	try
	{
		CComHeapPtr<OLECHAR> pszSelectedPath;
		hr = m_pFExplorer->GetSelectedFilePath(&pszSelectedPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if ((NULL != pszSelectedPath) && (0 != *pszSelectedPath))
		{
			// Construct a string to represent the base path to use for
			// finding all of the studies at the selected full path
			WCHAR szFolderName[VSI_MAX_PATH];
			WCHAR szSourceFolder[VSI_MAX_PATH];
			wcscpy_s(szSourceFolder, _countof(szSourceFolder), pszSelectedPath);
			wcscat_s(szSourceFolder, _countof(szSourceFolder), L"\\*.*");

			WIN32_FIND_DATA ff;
			hFile = FindFirstFile(szSourceFolder, &ff);
			BOOL bWorking = (INVALID_HANDLE_VALUE != hFile);

			while (bWorking)	// For each file in the study.
			{
				// Skip the dots, include subdirectories.
				if (ff.cFileName[0] == L'.')
				{
					NULL;
				}
				else if (ff.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
				{
					_snwprintf_s(szFolderName, _countof(szFolderName), L"%s\\%s\\%s",
						pszSelectedPath,
						ff.cFileName,
						CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));
					if (0 == _waccess(szFolderName, 0))
					{
						bHasStudy = TRUE;
						break;
					}
				}

				bWorking = FindNextFile(hFile, &ff);
			}
		}
	}
	VSI_CATCH(hr);

	if (INVALID_HANDLE_VALUE != hFile)
	{
		FindClose(hFile);
	}

	return bHasStudy;
}

LRESULT CVsiCopyStudyFrom::OnBnClickedRemove(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		CComPtr<IVsiEnumStudy> pEnumStudy;
		CComPtr<IVsiEnumSeries> pEnumSeries;

		int iStudy = 0, iSeries = 0;
		VSI_DATA_LIST_COLLECTION selection;
		memset(&selection, 0, sizeof(VSI_DATA_LIST_COLLECTION));
		selection.dwFlags = VSI_DATA_LIST_FILTER_SELECT_PARENT;
		selection.piStudyCount = &iStudy;
		selection.piSeriesCount = &iSeries;
		selection.ppEnumStudy = &pEnumStudy;
		selection.ppEnumSeries = &pEnumSeries;
		hr = m_pDataList->GetSelection(&selection);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (NULL != pEnumStudy)
		{
			CComPtr<IVsiStudy> pStudy;
			while (S_OK == pEnumStudy->Next(1, &pStudy, NULL))
			{
				CComVariant vNs;
				hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vId;
				hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = m_pDataList->RemoveItem(V_BSTR(&vNs), VSI_DATA_LIST_SELECTION_UPDATE);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComHeapPtr<OLECHAR> pszPath;
				hr = pStudy->GetDataPath(&pszPath);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				LPWSTR pszSlash = wcsrchr(pszPath, L'\\');
				CString strStudyPath(pszPath, (int)(pszSlash - pszPath));

				auto iter = m_setSelectedStudy.find(strStudyPath);
				if (iter != m_setSelectedStudy.end())
				{
					m_setSelectedStudy.erase(iter);
				}

				auto iterId = m_setSelectedStudyId.find(V_BSTR(&vId));
				if (iterId != m_setSelectedStudyId.end())
				{
					m_setSelectedStudyId.erase(iterId);
				}

				for (auto iter = m_setRemovedSeries.begin(); iter != m_setRemovedSeries.end(); )
				{
					const CString &strSeriesPath = *iter;

					if (0 == strSeriesPath.Find(strStudyPath))
					{
						iter = m_setRemovedSeries.erase(iter);
					}
					else
					{
						++iter;
					}
				}

				pStudy.Release();
			}
		}

		if (NULL != pEnumSeries)
		{
			CComPtr<IVsiSeries> pSeries;
			while (S_OK == pEnumSeries->Next(1, &pSeries, NULL))
			{
				CComHeapPtr<OLECHAR> pszPath;
				hr = pSeries->GetDataPath(&pszPath);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CString strStudyPath;
				CString strSeriesPath;

				// Is study still around
				LPWSTR pszSlash = wcsrchr(pszPath, L'\\');
				if (NULL != pszSlash)
				{
					*pszSlash = 0;
					strSeriesPath.SetString(pszPath);

					pszSlash = wcsrchr(pszPath, L'\\');
					if (NULL != pszSlash)
					{
						*pszSlash = 0;
						strStudyPath.SetString(pszPath);
					}
				}

				auto iter = m_setSelectedStudy.find(strStudyPath);
				if (iter != m_setSelectedStudy.end())
				{
					m_setRemovedSeries.insert(strSeriesPath);
				}
				else
				{
					auto iter = m_setRemovedSeries.find(strSeriesPath);
					if (iter != m_setRemovedSeries.end())
					{
						m_setRemovedSeries.erase(iter);
					}
				}

				pSeries.Release();
			}
		}

		hr = m_pDataList->UpdateItem(NULL, TRUE);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = m_pFExplorer->Refresh();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CalculateRequiredSize();
		UpdateAvailableSpace();
		UpdateUI();

		// Sets focus
		PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)GetDlgItem(IDC_IE_SELECTED), TRUE);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCopyStudyFrom::OnOperatorSelenOk(WORD, WORD, HWND, BOOL& bHandled)
{
	HRESULT hr(S_OK);
	bHandled = FALSE;

	try
	{
		int i = (int)m_wndOperator.SendMessage(CB_GETCURSEL);
		if (CB_ERR != i)
		{
			LPCWSTR pszId = (LPCWSTR)m_wndOperator.SendMessage(CB_GETITEMDATA, i, 0);
			if (! IS_INTRESOURCE(pszId))
			{
				UpdateUI();
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

BOOL CVsiCopyStudyFrom::IsStudyAdded(LPCWSTR pszFilePath)
{
	auto iter = m_setSelectedStudy.find(pszFilePath);
	return iter != m_setSelectedStudy.end();
}

BOOL CVsiCopyStudyFrom::IsStudyFullyAdded(LPCWSTR pszFilePath)
{
	HRESULT hr = S_OK;
	BOOL bFullyAdded = FALSE;

	try
	{
		CString strSelectedPath(pszFilePath);

		auto iter = m_setSelectedStudy.find(strSelectedPath);
		if (iter != m_setSelectedStudy.end())
		{
			bFullyAdded = TRUE;

			for (iter = m_setRemovedSeries.begin(); iter != m_setRemovedSeries.end(); ++iter)
			{
				const CString &strSeriesPath = *iter;

				if (0 == strSeriesPath.Find(strSelectedPath))
				{
					bFullyAdded = FALSE;
					break;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return bFullyAdded;
}

BOOL CVsiCopyStudyFrom::IsSeriesAdded(LPCWSTR pszFilePath)
{
	HRESULT hr = S_OK;
	BOOL bAdded(FALSE);

	try
	{
		CString strSeriesPath((LPCWSTR)pszFilePath);

		// Get study folder
		// Remove series.xml
		LPCWSTR pszSlash = wcsrchr(pszFilePath, L'\\');
		if (NULL != pszSlash)
		{
			CString strStudyPath(pszFilePath, (int)(pszSlash - pszFilePath));

			auto iter = m_setSelectedStudy.find(strStudyPath);
			if (iter != m_setSelectedStudy.end())
			{
				bAdded = TRUE;

				for (iter = m_setRemovedSeries.begin(); iter != m_setRemovedSeries.end(); ++iter)
				{
					const CString &strTest = *iter;

					if (strSeriesPath == strTest)
					{
						bAdded = FALSE;
						break;
					}
				}
			}
		}
	}
	VSI_CATCH(hr);

	return bAdded;
}

void CVsiCopyStudyFrom::SetStudyInfo(IVsiStudy *pStudy, bool bUpdateUI)
{
	HRESULT hr(S_OK);

	try
	{
		int iCol = VsiGetDiscreteValue<long>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_SERIES_ITEM_STUDYBROWSER, m_pPdm);
		int iOrder = VsiGetDiscreteValue<long>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_SERIES_ORDER_STUDYBROWSER, m_pPdm);

		m_studyOverview.SetStudy(
			pStudy,
			GetSeriesPropIdFromColumnId(iCol),
			(VSI_SYS_SORT_ORDER_DESCENDING == iOrder));
	}
	VSI_CATCH(hr);

	if (bUpdateUI)
	{
		bool bUpdate(false);
		if (pStudy != NULL)
		{
			if (!m_studyOverview.IsWindowVisible() || m_wndImageList.IsWindowVisible())
			{
				bUpdate = true;
			}
		}
		else
		{
			if (m_studyOverview.IsWindowVisible() || !m_wndImageList.IsWindowVisible())
			{
				bUpdate = true;
			}
		}

		if (bUpdate)
		{
			HDWP hdwp = BeginDeferWindowPos(2);

			if (pStudy != NULL)
			{
				hdwp = m_studyOverview.DeferWindowPos(hdwp, NULL, 0, 0, 0, 0,
					SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
				m_studyOverview.EnableWindow(TRUE);
				hdwp = m_wndImageList.DeferWindowPos(hdwp, NULL, 0, 0, 0, 0,
					SWP_NOSIZE | SWP_NOMOVE | SWP_HIDEWINDOW);
				m_wndImageList.EnableWindow(FALSE);
			}
			else
			{
				hdwp = m_wndImageList.DeferWindowPos(hdwp, NULL, 0, 0, 0, 0,
					SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
				m_wndImageList.EnableWindow(TRUE);
				hdwp = m_studyOverview.DeferWindowPos(hdwp, NULL, 0, 0, 0, 0,
					SWP_NOSIZE | SWP_NOMOVE | SWP_HIDEWINDOW);
				m_studyOverview.EnableWindow(FALSE);
			}

			EndDeferWindowPos(hdwp);
		}
	}
}

VSI_PROP_SERIES CVsiCopyStudyFrom::GetSeriesPropIdFromColumnId(int iCol)
{
	VSI_PROP_SERIES prop;

	switch (iCol)
	{
	case VSI_SYS_SORT_ITEM_SB_NAME:
		prop = VSI_PROP_SERIES_NAME;
		break;
	default:
	case VSI_SYS_SORT_ITEM_SB_DATE:
		prop = VSI_PROP_SERIES_CREATED_DATE;
		break;
	}

	return prop;
}

void CVsiCopyStudyFrom::UpdateFEStudyDetails()
{
	HRESULT hr = S_OK;
	WCHAR szStudyFileName[MAX_PATH];
	WCHAR szSeriesFileName[MAX_PATH];
	BOOL bStudyFound = FALSE, bSeriesFound = FALSE;

	try
	{
		CComHeapPtr<OLECHAR> pszSelectedPath;
		hr = m_pFExplorer->GetSelectedFilePath(&pszSelectedPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// check for a study
		// if positive - update study info
		_snwprintf_s(szStudyFileName, _countof(szStudyFileName),
			L"%s\\%s", pszSelectedPath, CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));
		if (0 == _waccess(szStudyFileName, 0))
		{
			CComPtr<IVsiStudy> pStudy; // don't keep the pointer around
			hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			StopPreview(VSI_PREVIEW_LIST_CLEAR);

			CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
			hr = pPersist->LoadStudyData(szStudyFileName, 0, NULL);
			if (SUCCEEDED(hr))
			{
				m_pItemPreview = pStudy;

				SetStudyInfo(pStudy);

				m_typeSel = VSI_DATA_LIST_STUDY;

				bStudyFound = TRUE;
			}
			else
			{
				// TODO: Corrupted study - show info
			}
		}

		// check for a series
		// if positive - fill up preview image list
		_snwprintf_s(szSeriesFileName, _countof(szSeriesFileName),
			L"%s\\%s", pszSelectedPath, CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME)));
		if (0 == _waccess(szSeriesFileName, 0))
		{
			m_pItemPreview.Release();
			hr = m_pItemPreview.CoCreateInstance(__uuidof(VsiSeries));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (VSI_DATA_LIST_STUDY == m_typeSel)
			{
				SetStudyInfo(NULL);
			}
			else
			{
				StopPreview(VSI_PREVIEW_LIST_CLEAR);
			}

			m_typeSel = VSI_DATA_LIST_SERIES;

			CComQIPtr<IVsiPersistSeries> pPersist(m_pItemPreview);
			hr = pPersist->LoadSeriesData(szSeriesFileName, VSI_DATA_TYPE_IMAGE_LIST | VSI_DATA_TYPE_NO_SESSION_LINK, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			StartPreview();

			bSeriesFound = TRUE;
		}
	}
	VSI_CATCH(hr);

	if (!bStudyFound && !bSeriesFound)
	{
		if (VSI_DATA_LIST_STUDY == m_typeSel)
		{
			SetStudyInfo(NULL, false);
		}
		else
		{
			StopPreview(VSI_PREVIEW_LIST_CLEAR);
		}
	}
}

void CVsiCopyStudyFrom::PreviewImages()
{
	HRESULT hr = S_OK;

	try
	{
		if (m_pItemPreview == NULL)
		{
			StopPreview();
		}
		else
		{
			if (m_pEnumImage == NULL)
			{
				CComQIPtr<IVsiSeries> pSeries(m_pItemPreview);
				if (pSeries != NULL)
				{
					hr = pSeries->GetImageEnumerator(
						VSI_PROP_IMAGE_ACQUIRED_DATE,
						TRUE,
						&m_pEnumImage, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					LONG lCount(0);
					hr = pSeries->GetImageCount(&lCount, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					ListView_SetItemCount(m_wndImageList, lCount);
				}
			}

			IVsiImage *ppImage[2];
			ULONG ulCount = 0;
			hr = m_pEnumImage->Next(_countof(ppImage), ppImage, &ulCount);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			for (ULONG l = 0; l < ulCount; ++l)
			{
				CComHeapPtr<OLECHAR> pFile;
				ppImage[l]->GetThumbnailFile(&pFile);

				SYSTEMTIME stDate;
				CComVariant vDate;
				hr = ppImage[l]->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vDate);
				if (FAILED(hr))
				{
					ppImage[l]->Release();
					continue;  // skip
				}

				COleDateTime date(vDate);
				date.GetAsSystemTime(stDate);

				// The time is stored as UTC. Convert it to local time for display.
				BOOL bRet = SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);
				if (!bRet)
				{
					ppImage[l]->Release();
					continue;  // skip
				}

				WCHAR szText[500], szDate[50], szTime[50];
				GetDateFormat(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &stDate, NULL, szDate, _countof(szDate));
				GetTimeFormat(LOCALE_USER_DEFAULT, 0, &stDate, NULL, szTime, _countof(szTime));

				CComVariant vName;
				ppImage[l]->GetProperty(VSI_PROP_IMAGE_NAME, &vName);

				if (VT_EMPTY != V_VT(&vName) && (lstrlen(V_BSTR(&vName)) != 0))
					swprintf_s(szText, _countof(szText), L"%s\r\n%s %s", V_BSTR(&vName), szTime, szDate);
				else
					swprintf_s(szText, _countof(szText), L"%s %s", szTime, szDate);

				LVITEM lvi = { LVIF_TEXT | LVIF_IMAGE | LVIF_PARAM };
				lvi.pszText = szText;
				lvi.iImage = (int)(INT_PTR)(LPCWSTR)pFile;
				ppImage[l]->AddRef();
				lvi.lParam = (LPARAM)(ppImage[l]);
				ListView_InsertItem(m_wndImageList, &lvi);

				ppImage[l]->Release();
			}

			if (S_OK != hr)
			{
				StopPreview();
			}

			m_wndImageList.Invalidate(FALSE);
		}
	}
	VSI_CATCH(hr);
}

void CVsiCopyStudyFrom::CalculateRequiredSize()
{
	HRESULT hr = S_OK;
	INT64 iTotalSize = 0;

	try
	{
		for (auto iter = m_setSelectedStudy.begin(); iter != m_setSelectedStudy.end(); ++iter)
		{
			const CString &strStudyPath = *iter;

			if (strStudyPath.GetLength() > 0)
			{
				// Construct a string to represent the base path to use for
				// finding all of the series at the selected study path
				CString strSeriesPath;
				CString strStudyFile(strStudyPath);
				strStudyFile += L"\\*.*";

				WIN32_FIND_DATA fd;
				HANDLE hFile = FindFirstFile(strStudyFile, &fd);
				BOOL bWorking = (INVALID_HANDLE_VALUE != hFile);

				while (bWorking)	// For each file in the study.
				{
					// Skip the dots, include subdirectories.
					if (L'.' == fd.cFileName[0])
					{
						NULL;
					}
					else if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
					{
						// Check series
						strSeriesPath.Format(L"%s\\%s",	(LPCWSTR)strStudyPath, fd.cFileName);

						// Check whether series folder is not excluded from the list
						auto iterSeries = m_setRemovedSeries.find(strSeriesPath);
						if (iterSeries == m_setRemovedSeries.end())
						{
							iTotalSize += VsiGetDirectorySize(strSeriesPath);
						}
					}

					bWorking = FindNextFile(hFile, &fd);
				}

				// Done with the current directory. Close the handle and return.
				if (INVALID_HANDLE_VALUE != hFile)
					FindClose(hFile);
			}
		}
	}
	VSI_CATCH(hr);

	if (SUCCEEDED(hr))
		m_iRequiredSize = iTotalSize;
	else
		m_iRequiredSize = 0;
}

LRESULT CVsiCopyStudyFrom::OnWindowPosChanged(UINT, WPARAM, LPARAM lp, BOOL &bHandled)
{
	WINDOWPOS *pPos = (WINDOWPOS*)lp;

	HRESULT hr = S_OK;

	try
	{
		// Updates preview list when turn visible
		if (SWP_SHOWWINDOW & pPos->flags)
		{
			// Refresh preview
			if (VSI_DATA_LIST_STUDY != m_typeSel)
			{
				StopPreview(VSI_PREVIEW_LIST_CLEAR);
				m_pItemPreview.Release();
				m_typeSel = VSI_DATA_LIST_NONE;
			}
			else
			{
				SetStudyInfo(NULL, false);
			}

			HRESULT hr = m_pCursorMgr->SetState(VSI_CURSOR_STATE_ON, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (SWP_HIDEWINDOW & pPos->flags)
		{
			// Clear preview to save memory
			if (VSI_DATA_LIST_STUDY != m_typeSel)
			{
				StopPreview(VSI_PREVIEW_LIST_CLEAR);
				m_pItemPreview.Release();
				m_typeSel = VSI_DATA_LIST_NONE;
			}
		}
	}
	VSI_CATCH(hr);

	bHandled = FALSE;

	return 0;
}