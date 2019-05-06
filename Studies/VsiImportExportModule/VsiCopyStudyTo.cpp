/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiCopyStudyTo.cpp
**
**	Description:
**		Implementation of CVsiCopyStudyTo
**
*******************************************************************************/

#include "stdafx.h"
#include <shlguid.h>
#include <VsiWTL.h>
#include <ATLComTime.h>
#include <VsiGlobalDef.h>
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
#include <VsiImportExportXml.h>
#include "VsiImpExpUtility.h"
#include "VsiReportExportStatus.h"
#include "VsiCopyStudyTo.h"


#define VSI_AUTOCOMPLETE_NAME L"VsiHistoryCopyTo.xml"


class CVsiCopyStudyToConfirmDlg :
	public CDialogImpl<CVsiCopyStudyToConfirmDlg>
{
public:
	enum { IDD = IDD_COPY_STUDY_TO_CONFIRM };

	enum
	{
		ACTION_OVERWRITE = 100,
		ACTION_SKIP,
	};

protected:
	BEGIN_MSG_MAP(CVsiCopyStudyToConfirmDlg)
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
			
			CheckDlgButton(IDC_CSTC_SKIP, BST_CHECKED);

			// Text branding
			{
				SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_COPY_STUDY_TO)));
				GetDlgItem(IDC_CSTC_LABEL).SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_STUDIES_EXISTS_CONF)));
				GetDlgItem(IDC_CSTC_OVERWRITE).SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_STUDIES_OVERWRITE)));
				GetDlgItem(IDC_CSTC_SKIP).SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_STUDIES_SKIP)));				
			}
		}
		VSI_CATCH(hr);

		if (FAILED(hr))
			EndDialog(IDCANCEL);

		return 0;
	}

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&)
	{
		VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Copy Study To Confirm - OK");

		int iRet(IDCANCEL);

		if (BST_CHECKED == IsDlgButtonChecked(IDC_CSTC_OVERWRITE))
		{
			iRet = ACTION_OVERWRITE;
		}
		else if (BST_CHECKED == IsDlgButtonChecked(IDC_CSTC_SKIP))
		{
			iRet = ACTION_SKIP;
		}

		EndDialog(iRet);

		return 0;		
	}

	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
	{
		VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Copy Study To Confirm - Cancel");

		EndDialog(IDCANCEL);

		return 0;		
	}
};


CVsiCopyStudyTo::CVsiCopyStudyTo() :
	m_dwFlags(0),
	m_bWriterSelected(FALSE)
{
}

CVsiCopyStudyTo::~CVsiCopyStudyTo()
{
	_ASSERT(m_pApp == NULL);
	_ASSERT(m_pCmdMgr == NULL);
	_ASSERT(m_pPdm == NULL);
	_ASSERT(m_pXmlDoc == NULL);
	_ASSERT(m_pStudyMgr == NULL);
	_ASSERT(m_pWriter == NULL);
}

STDMETHODIMP CVsiCopyStudyTo::Activate(
	IVsiApp *pApp,
	HWND hwndParent)
{
	HRESULT hr = S_OK;

	try
	{
		m_pApp = pApp;
		VSI_CHECK_INTERFACE(m_pApp, VSI_LOG_ERROR, NULL);

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

		// CD/DVD
		hr = m_pWriter.CoCreateInstance(__uuidof(VsiCdDvdWriter));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating DVD writer object");

		HWND hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

LRESULT CVsiCopyStudyTo::OnWindowPosChanged(UINT, WPARAM, LPARAM lp, BOOL &bHandled)
{
	WINDOWPOS *pPos = (WINDOWPOS*)lp;

	HRESULT hr = S_OK;

	try
	{
		// Updates preview list when turn visible
		if (SWP_SHOWWINDOW & pPos->flags)
		{
			HRESULT hr = m_pCursorMgr->SetState(VSI_CURSOR_STATE_ON, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	bHandled = FALSE;

	return 1;//act as if nothing happened
}

STDMETHODIMP CVsiCopyStudyTo::Deactivate()
{
	HRESULT hr = S_FALSE;

	if (IsWindow())
	{
		BOOL bRet = DestroyWindow();
		_ASSERT(bRet);

		hr = S_OK;
	}

	m_pWriter.Release();
	m_pXmlDoc.Release();
	m_pStudyMgr.Release();
	m_pCmdMgr.Release();
	m_pPdm.Release();
	m_pApp.Release();

	return hr;
}

STDMETHODIMP CVsiCopyStudyTo::GetWindow(HWND *phWnd)
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

STDMETHODIMP CVsiCopyStudyTo::GetAccelerator(HACCEL *phAccel)
{
	UNREFERENCED_PARAMETER(phAccel);
	return E_NOTIMPL;
}

STDMETHODIMP CVsiCopyStudyTo::GetIsBusy(
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

STDMETHODIMP CVsiCopyStudyTo::PreTranslateMessage(
	MSG *pMsg, BOOL *pbHandled)
{
	*pbHandled = IsDialogMessage(pMsg);

	return S_OK;
}

LRESULT CVsiCopyStudyTo::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		OSVERSIONINFO osv = {0};
		osv.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
		GetVersionEx((LPOSVERSIONINFO) &osv);
		if (VER_PLATFORM_WIN32_NT == osv.dwPlatformId && osv.dwMajorVersion > 5)//vista or later
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
				pLayout->AddControl(0, IDC_IE_PATH, VSI_WL_SIZE_X);
				pLayout->AddControl(0, IDC_IE_FOLDER_TREE, VSI_WL_SIZE_XY);
				pLayout->AddControl(0, IDC_IE_SELECTED, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_REQUIRED_LABEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_REQUIRED, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_AVAILABLE_LABEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_AVAILABLE, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_OPTIONS, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_CST_AUTO_DEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_CST_SAVEAS_LABEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_SAVEAS, VSI_WL_MOVE_X);

				pLayout.Detach();
			}
		}

		SetWindowText(CString(MAKEINTRESOURCE(IDS_CPYSTY_COPY_STUDY_TO)));


		m_wndStudyName.SubclassWindow(GetDlgItem(IDC_IE_SAVEAS));

		// Auto Complete - Keep going on error
		try
		{
			hr = m_pAutoComplete.CoCreateInstance(CLSID_AutoComplete);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pAutoCompleteSource.CoCreateInstance(__uuidof(VsiAutoCompleteSource));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			WCHAR pszPath[MAX_PATH];
			BOOL bRet = VsiGetApplicationDataDirectory(
				AtlFindStringResourceInstance(IDS_VSI_PRODUCT_NAME),
				pszPath);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, NULL);
			
			PathAppend(pszPath, VSI_AUTOCOMPLETE_NAME);

			hr = m_pAutoCompleteSource->LoadSettings(pszPath);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
			CComPtr<IEnumString> pEnumString;
			hr = m_pAutoCompleteSource->GetSource(&pEnumString);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pAutoComplete->Init(GetDlgItem(IDC_IE_SAVEAS), pEnumString, NULL, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
			CComQIPtr<IAutoComplete2> pAutoComplete2(m_pAutoComplete);
			if (pAutoComplete2 != NULL)
			{
				hr = pAutoComplete2->SetOptions(
					ACO_AUTOSUGGEST | ACO_UPDOWNKEYDROPSLIST);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}
		VSI_CATCH(hr);

		// Get required size
		INT64 iTotalSize = 0;
		CComVariant vSize, vLocked;
		CComPtr<IXMLDOMNodeList> pListStudy;
		hr = m_pXmlDoc->selectNodes(
			CComBSTR(L"//" VSI_STUDY_XML_ELM_STUDIES L"/" VSI_STUDY_XML_ELM_STUDY),
			&pListStudy);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting studies");

		BOOL bLocked = FALSE;
		long lNumUnlocked = 0;
		long lLength = 0;
		hr = pListStudy->get_length(&lLength);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");
		for (long i = 0; i < lLength; ++i)
		{
			CComPtr<IXMLDOMNode> pNode;
			hr = pListStudy->get_item(i, &pNode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

			CComQIPtr<IXMLDOMElement> pElemParam(pNode);

			hr = pElemParam->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), &vSize);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_SIZE));

			hr = VariantChangeTypeEx(
				&vSize, &vSize,
				MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US), SORT_DEFAULT),
				0, VT_UI8);

			iTotalSize += V_UI8(&vSize);

			hr = pElemParam->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_LOCKED), &vLocked);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_LOCKED));

			bLocked = FALSE;
			if (S_OK == hr)
			{
				bLocked = (0 == lstrcmpi(VSI_VAL_TRUE, V_BSTR(&vLocked)));
			}

			if (V_BOOL(&vLocked) == VARIANT_FALSE)
				lNumUnlocked++;
		}

		m_iStudyCount = (int)lLength;
		m_iRequiredSize = iTotalSize;

		hr = m_pFExplorer.CoCreateInstance(__uuidof(VsiFolderExplorer));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Initialize Folder Name field
		if (1 == lLength)
		{
			CComVariant vTarget;
			CComPtr<IXMLDOMNode> pNode;
			hr = pListStudy->get_item(0, &pNode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting studies");

			CComQIPtr<IXMLDOMElement> pElemParam(pNode);

			hr = pElemParam->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_TARGET), &vTarget);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_TARGET));

			SetDlgItemText(IDC_IE_SAVEAS, V_BSTR(&vTarget));
		}

		UpdateSizeUI();
		UpdateUI();

		// Refresh tree
		SetDlgItemText(IDC_IE_PATH, L"Loading...");
		PostMessage(WM_VSI_REFRESH_TREE);

		// Sets focus
		PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)GetDlgItem(IDC_IE_SAVEAS), TRUE);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCopyStudyTo::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	if (m_pFExplorer != NULL)
	{
		HRESULT hr = S_OK;
		try
		{
			CComHeapPtr<OLECHAR> pszPathName;
			hr = m_pFExplorer->GetSelectedNamespacePath(&pszPathName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_pFExplorer->Uninitialize();
			m_pFExplorer.Release();

			CVsiRange<LPCWSTR> pParamPath;
			hr = m_pPdm->GetParameter(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_PATH_EXPORT_STUDY, &pParamPath);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			pParamPath.SetValue(pszPathName);
		}
		VSI_CATCH(hr)
	}

	return 0;
}

LRESULT CVsiCopyStudyTo::OnRefreshTree(UINT, WPARAM, LPARAM, BOOL&)
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

LRESULT CVsiCopyStudyTo::OnSetSelection(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	CWaitCursor wait;

	try
	{
		// Set default selected folder
		CVsiRange<LPCWSTR> pParamPath;
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_SYS_PATH_EXPORT_STUDY, &pParamPath);
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

LRESULT CVsiCopyStudyTo::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Copy Study To - OK");

	HRESULT hr = S_OK;

	try
	{
		WCHAR szFolderName[MAX_PATH];
		GetDlgItemText(IDC_IE_SAVEAS, szFolderName, _countof(szFolderName));

		// Save Auto Complete settings
		if (*szFolderName != 0 && NULL != m_pAutoCompleteSource)
		{
			WCHAR pszPath[MAX_PATH];
			BOOL bRet = VsiGetApplicationDataDirectory(
				AtlFindStringResourceInstance(IDS_VSI_PRODUCT_NAME),
				pszPath);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, NULL);

			PathAppend(pszPath, VSI_AUTOCOMPLETE_NAME);

			hr = m_pAutoCompleteSource->Push(szFolderName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pAutoCompleteSource->SaveSettings(pszPath, 10);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Get selected folder from Folder Explorer
		CComHeapPtr<OLECHAR> pszPath;
		hr = m_pFExplorer->GetSelectedFilePath(&pszPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vDoc(m_pXmlDoc);

		if (m_bWriterSelected)
		{
			CComPtr<IVsiCdDvdWriterView> pCdDvdWriterView;
			hr = pCdDvdWriterView.CoCreateInstance(__uuidof(VsiCdDvdWriterView));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			pCdDvdWriterView->Initialize(m_strDriveLetter, m_pWriter, &vDoc);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Pass IXMLDOMDocument
			CComVariant vCdDvdView(pCdDvdWriterView);
			m_pCmdMgr->InvokeCommand(ID_CMD_OPEN_BURN_DISC, &vCdDvdView);

			return 0;
		}
		else
		{
			// Read only test
			if (CheckReadOnly(pszPath))
			{
				ShowReadOnlyMessage();
				return 0;
			}

			if (m_iAvailableSize < m_iRequiredSize)
			{
				ShowNotEnoughSpaceMessage();
				return 0;
			}

			int iOverwriteAction(0);
			CString strAction;
			DWORD dwFlags = 0;

			// Overwrite test
			int iOverwrite = CheckForOverwrite();
			if (0 < iOverwrite)
			{
				// Overwrite, Skip or  Cancel
				CVsiCopyStudyToConfirmDlg dlg;
				iOverwriteAction = (int)dlg.DoModal();
				if (IDCANCEL == iOverwriteAction)
					return 0;

				if (CVsiCopyStudyToConfirmDlg::ACTION_OVERWRITE == iOverwriteAction)
				{
					dwFlags |= VSI_STUDY_MOVER_OVERWRITE;
					strAction = L"Overwrite";
				}
				else if (CVsiCopyStudyToConfirmDlg::ACTION_SKIP == iOverwriteAction)
				{
					strAction = L"Skip";
				}
			}

			VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Copy study to: started (action = '%s')", strAction.GetString()));

			// Create StudyMover object
			CComPtr<IVsiStudyMover> pStudyMover;
			hr = pStudyMover.CoCreateInstance(__uuidof(VsiStudyMover));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = UpdateJobXml(0 != (dwFlags & VSI_STUDY_MOVER_OVERWRITE));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pStudyMover->Initialize(m_hWnd, &vDoc, m_iRequiredSize, dwFlags);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pStudyMover->Export(pszPath, (LPCOLESTR)szFolderName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Display status dialog
		CVsiReportExportStatus statusReport(
			V_UNKNOWN(&vDoc),
			CVsiReportExportStatus::REPORT_TYPE_STUDY_EXPORT);
		statusReport.DoModal(m_hWnd);

		// Close
		m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_COPY_STUDY_TO, NULL);
	}
	VSI_CATCH(hr);

	VSI_LOG_MSG(VSI_LOG_INFO, L"Copy study to: ended");

	return 0;
}

LRESULT CVsiCopyStudyTo::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Copy Study To - Cancel");

	m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_COPY_STUDY_TO, NULL);

	return 0;
}

STDMETHODIMP CVsiCopyStudyTo::Initialize(const VARIANT *pvParam)
{
	HRESULT hr(S_OK);

	try
	{
		VSI_VERIFY(VT_DISPATCH == V_VT(pvParam), VSI_LOG_ERROR, NULL);

		m_pXmlDoc = V_DISPATCH(pvParam);
		VSI_CHECK_INTERFACE(m_pXmlDoc, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

BOOL CVsiCopyStudyTo::GetAvailablePath(LPWSTR pszBuffer, int iBufferLength)
{
	return GetSelectedPath(pszBuffer, iBufferLength);
}

LRESULT CVsiCopyStudyTo::OnTreeSelChanged(int, LPNMHDR, BOOL&)
{
	if (m_pFExplorer != NULL)
	{
		m_strDriveLetter.Empty();

		WCHAR szPath[VSI_MAX_PATH];
		if (GetSelectedPath(szPath, _countof(szPath)))
		{
			m_bWriterSelected = FALSE;
			if (szPath[0] >= L'A' && szPath[0] <= L'Z')
			{
				CString strDriveLetter(szPath, 1);
				m_strDriveLetter = strDriveLetter;
				HRESULT hr = m_pWriter->IsDeviceWriter(m_strDriveLetter, &m_bWriterSelected);
				if (FAILED(hr))
					VSI_LOG_MSG(VSI_LOG_ERROR, L"CVsiCdDvdWriter::IsDeviceWriter failed");
			}
		}

		UpdateAvailableSpace();
		UpdateSizeUI();
		UpdateUI();
	}
	return 0;
}

LRESULT CVsiCopyStudyTo::OnCheckAllowChildren(int, LPNMHDR pnmh, BOOL &)
{
	NMTESTREMOVABLEDEVICE* pCheck = (NMTESTREMOVABLEDEVICE*)pnmh;

	if (lstrlen(pCheck->pszDevicePath) > 0)
	{
		HRESULT hr = S_OK;

		try
		{
			CString strDriveLetter(pCheck->pszDevicePath, 1);

			// test for DVD writer
			VSI_WRITE_DEVICE_INFO DeviceInfo;
			memset(&DeviceInfo, 0, sizeof(DeviceInfo));

			hr = m_pWriter->GetDeviceInfo(strDriveLetter, &DeviceInfo);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr && DeviceInfo.bBurner)
				pCheck->bAllowChildren = FALSE;
		}
		VSI_CATCH(hr);
	}

	return 0;
}


/// <summary>
///	The user has done something to change the text in the "Save As" edit control. Update the buttons
///	because the new name may have changed the overwrite status.
/// </summary>
///
///
/// <returns>0, always</returns>
LRESULT CVsiCopyStudyTo::OnStudyNameChanged(WORD, WORD, HWND, BOOL&)
{
	UpdateUI();

	return 0;
}

void CVsiCopyStudyTo::UpdateUI()
{
	CVsiImportExportUIHelper<CVsiCopyStudyTo>::UpdateUI();

	SetTextBoxLimit();

	int iItems = GetNumberOfSelectedItems();
	if (1 < iItems)
	{
		GetDlgItem(IDC_IE_OPTIONS).ShowWindow(SW_HIDE);
		GetDlgItem(IDC_CST_SAVEAS_LABEL).ShowWindow(SW_HIDE);
		GetDlgItem(IDC_IE_SAVEAS).ShowWindow(SW_HIDE);

		GetDlgItem(IDC_IE_SAVEAS).EnableWindow(FALSE);
	}
	else
	{
		GetDlgItem(IDC_IE_SAVEAS).EnableWindow(TRUE);

		GetDlgItem(IDC_IE_OPTIONS).ShowWindow(SW_SHOWNA);
		GetDlgItem(IDC_CST_SAVEAS_LABEL).ShowWindow(SW_SHOWNA);
		GetDlgItem(IDC_IE_SAVEAS).ShowWindow(SW_SHOWNA);
	}

	BOOL bOkayEnabled(FALSE);
	WCHAR szPath[VSI_MAX_PATH];
	if (GetSelectedPath(szPath, _countof(szPath)))
	{
		bOkayEnabled = 0 != *szPath;
		if (bOkayEnabled && 1 == iItems)
		{
			bOkayEnabled = 0 < m_wndStudyName.GetWindowTextLength();
		}
	}

	GetDlgItem(IDOK).EnableWindow(bOkayEnabled);
}

void CVsiCopyStudyTo::UpdateAvailableSpace()
{
	if (m_bWriterSelected)
	{
		HRESULT hr = S_OK;
		m_iAvailableSize = 0;

		CWaitCursor wait;

		try
		{
			VSI_WRITE_DEVICE_INFO DeviceInfo;
			memset(&DeviceInfo, 0, sizeof(DeviceInfo));

			hr = m_pWriter->GetDeviceInfo(m_strDriveLetter, &DeviceInfo);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if ((S_OK == hr) && DeviceInfo.bBurner)
			{
				CWaitCursor wait;

				VSI_WRITE_MEDIA_INFO MediaInfo;
				memset(&MediaInfo, 0, sizeof(MediaInfo));

				hr = m_pWriter->GetMediaInfo(m_strDriveLetter, &MediaInfo);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if ((S_OK == hr) && MediaInfo.bWritable)
				{
					m_iAvailableSize = MediaInfo.iTotalSizeBytes;
				}
			}
		}
		VSI_CATCH(hr);
	}
	else
	{
		CVsiImportExportUIHelper<CVsiCopyStudyTo>::UpdateAvailableSpace();
	}
}

BOOL CVsiCopyStudyTo::UpdateSizeUI()
{
	BOOL bRet = CVsiImportExportUIHelper<CVsiCopyStudyTo>::UpdateSizeUI();

	int iItems = GetNumberOfSelectedItems();

	// Update the number of items currently selected
	WCHAR szText[200];
	WCHAR szNumItems[20];
	if (iItems > 0)
		swprintf_s(szNumItems, _countof(szNumItems), L"%d", iItems);
	else
		lstrcpy(szNumItems, L"No");

	if (iItems == 1)
	{
		CString strText(MAKEINTRESOURCE(IDS_CPYSTY_ONE_STUDY_SELECTED));
		swprintf_s(szText, _countof(szText), strText);
	}
	else
	{
		CString strText(MAKEINTRESOURCE(IDS_CPYSTY_N_STUDIES_SELECTED));
		swprintf_s(szText, _countof(szText), strText, szNumItems);
	}

	GetDlgItem(IDC_IE_SELECTED).SetWindowText(szText);

	return bRet;
}

int CVsiCopyStudyTo::CheckForOverwrite()
{
	HRESULT hr = S_OK;
	int iNum = 0;
	BOOL bOverwrite = FALSE;

	if (m_bWriterSelected)
		return iNum;

	try
	{
		CComVariant vTarget;
		CComPtr<IXMLDOMNodeList> pListStudy;
		hr = m_pXmlDoc->selectNodes(
			CComBSTR(L"//" VSI_STUDY_XML_ELM_STUDIES L"/" VSI_STUDY_XML_ELM_STUDY),
			&pListStudy);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting studies");

		long lLength = 0;
		hr = pListStudy->get_length(&lLength);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");
		for (long l = 0; l < lLength; l++)
		{
			CComPtr<IXMLDOMNode> pNode;
			hr = pListStudy->get_item(l, &pNode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

			CComQIPtr<IXMLDOMElement> pElemParam(pNode);

			if (1 == GetNumberOfSelectedItems())
			{
				WCHAR szBuffer[VSI_MAX_PATH];
				GetUserSuppliedName(szBuffer, _countof(szBuffer));
				bOverwrite = WouldOverwrite(pElemParam, szBuffer);
			}
			else
			{
				hr = pElemParam->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_TARGET), &vTarget);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_TARGET));
				bOverwrite = WouldOverwrite(pElemParam, V_BSTR(&vTarget));
			}

			if (bOverwrite)
			{
				iNum++;
			}
		}
	}
	VSI_CATCH(hr);

	return iNum;
}

int CVsiCopyStudyTo::GetNumberOfSelectedItems()
{
	return m_iStudyCount;
}

BOOL CVsiCopyStudyTo::WouldOverwrite(IXMLDOMElement *pXmlElm, LPCWSTR pszTargetFolderName)
{
	BOOL bOverwrite(FALSE);
	HRESULT hr(S_OK);

	try
	{
		// Target
		CComHeapPtr<OLECHAR> pszPathDest;
		hr = m_pFExplorer->GetSelectedFilePath(&pszPathDest);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CString strPathDest;
		CString strFileStudy(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME));
		strPathDest.Format(L"%s\\%s\\%s", pszPathDest, pszTargetFolderName, strFileStudy);

		if (0 == _waccess(strPathDest, 0))
		{
			CComPtr<IVsiStudy> pStudyDest;
			hr = pStudyDest.CoCreateInstance(__uuidof(VsiStudy));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComQIPtr<IVsiPersistStudy> pPersistStudy(pStudyDest);
			VSI_CHECK_INTERFACE(pPersistStudy, VSI_LOG_ERROR, NULL);

			hr = pPersistStudy->LoadStudyData(strPathDest, VSI_DATA_TYPE_SERIES_LIST, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Source
			CComPtr<IXMLDOMNodeList> pListSeries;
			hr = pXmlElm->selectNodes(CComBSTR(VSI_SERIES_XML_ELM_SERIES), &pListSeries);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			long lSeries(0);
			hr = pListSeries->get_length(&lSeries);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			for (long i = 0; i < lSeries; ++i)
			{
				CComPtr<IXMLDOMNode> pNode;
				hr = pListSeries->get_item(i, &pNode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

				CComQIPtr<IXMLDOMElement> pElemParam(pNode);
				VSI_CHECK_INTERFACE(pElemParam, VSI_LOG_ERROR, NULL);

				CComVariant vSeriesId;
				hr = pElemParam->getAttribute(CComBSTR(VSI_SERIES_XML_ATT_ID), &vSeriesId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_SERIES_XML_ATT_ID));

				CComPtr<IVsiSeries> pSeries;
				hr = pStudyDest->GetSeries(V_BSTR(&vSeriesId), &pSeries);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				bOverwrite = (S_OK == hr);
				if (bOverwrite)
				{
					break;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return bOverwrite;
}

HRESULT CVsiCopyStudyTo::UpdateJobXml(BOOL bOverwrite)
{
	HRESULT hr = S_OK;

	try
	{
		CComHeapPtr<OLECHAR> pszPathDest;
		hr = m_pFExplorer->GetSelectedFilePath(&pszPathDest);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vTarget;
		CComPtr<IXMLDOMNodeList> pListStudy;
		hr = m_pXmlDoc->selectNodes(
			CComBSTR(L"//" VSI_STUDY_XML_ELM_STUDIES L"/" VSI_STUDY_XML_ELM_STUDY),
			&pListStudy);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting studies");

		long lLength = 0;
		hr = pListStudy->get_length(&lLength);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");
		for (long l = 0; l < lLength; l++)
		{
			CComPtr<IXMLDOMNode> pNode;
			hr = pListStudy->get_item(l, &pNode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

			CComQIPtr<IXMLDOMElement> pElemStudy(pNode);
			VSI_CHECK_INTERFACE(pElemStudy, VSI_LOG_ERROR, NULL);

			CString strTarget;

			if (1 == GetNumberOfSelectedItems())
			{
				WCHAR szBuffer[VSI_MAX_PATH];
				GetUserSuppliedName(szBuffer, _countof(szBuffer));
				strTarget = szBuffer;
			}
			else
			{
				hr = pElemStudy->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_TARGET), &vTarget);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_TARGET));
				strTarget = V_BSTR(&vTarget);
			}

			CString strPathDest;
			CString strFileStudy(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME));
			strPathDest.Format(L"%s\\%s\\%s", pszPathDest, strTarget.GetString(), strFileStudy);

			if (0 == _waccess(strPathDest, 0))
			{
				CComPtr<IVsiStudy> pStudyDest;
				hr = pStudyDest.CoCreateInstance(__uuidof(VsiStudy));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComQIPtr<IVsiPersistStudy> pPersistStudy(pStudyDest);
				VSI_CHECK_INTERFACE(pPersistStudy, VSI_LOG_ERROR, NULL);

				hr = pPersistStudy->LoadStudyData(strPathDest, VSI_DATA_TYPE_SERIES_LIST, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Source
				CComPtr<IXMLDOMNodeList> pListSeries;
				hr = pElemStudy->selectNodes(CComBSTR(VSI_SERIES_XML_ELM_SERIES), &pListSeries);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				long lSeries(0);
				hr = pListSeries->get_length(&lSeries);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				for (long i = 0; i < lSeries; ++i)
				{
					CComPtr<IXMLDOMNode> pNode;
					hr = pListSeries->get_item(i, &pNode);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

					CComQIPtr<IXMLDOMElement> pElemSeries(pNode);
					VSI_CHECK_INTERFACE(pElemSeries, VSI_LOG_ERROR, NULL);

					CComVariant vSeriesId;
					hr = pElemSeries->getAttribute(CComBSTR(VSI_SERIES_XML_ATT_ID), &vSeriesId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure getting attribute (%s)", VSI_SERIES_XML_ATT_ID));

					CComPtr<IVsiSeries> pSeries;
					hr = pStudyDest->GetSeries(V_BSTR(&vSeriesId), &pSeries);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						if (bOverwrite)
						{
							hr = pElemSeries->setAttribute(CComBSTR(VSI_SERIES_XML_ATT_OVERWRITE),
								CComVariant(bOverwrite ? L"true" : L"false"));
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,
								VsiFormatMsg(L"Failure setting attribute (%s)", VSI_SERIES_XML_ATT_OVERWRITE));
						}
						else
						{
							// Skip series
							hr = pElemStudy->removeChild(pNode, NULL);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
					}
				}
			}
		}

#ifdef _DEBUG
{
		WCHAR szPath[MAX_PATH];
		GetModuleFileName(NULL, szPath, _countof(szPath));
		LPWSTR pszSpt = wcsrchr(szPath, L'\\');
		lstrcpy(pszSpt + 1, VSI_FOLDER_DEBUG);
		lstrcat(szPath, L"\\copyto1.xml");
		m_pXmlDoc->save(CComVariant(szPath));
}
#endif

	}
	VSI_CATCH(hr);

	return hr;
}

