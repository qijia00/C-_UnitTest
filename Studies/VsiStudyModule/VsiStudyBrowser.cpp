/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudyBrowser.cpp
**
**	Description:
**		Implementation of CVsiStudyBrowser
**
*******************************************************************************/

#include "stdafx.h"
#include <shellapi.h>
#include <VsiGlobalDef.h>
#include <VsiSaxUtils.h>
#include <VsiWtl.h>
#include <VsiServiceProvider.h>
#include <VsiServiceKey.h>
#include <VsiCommUtl.h>
#include <VsiLicenseUtils.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiPdmModule.h>
#include <VsiModeUtils.h>
#include <VsiModeModule.h>
#include <VsiAnalRsltXmlTags.h>
#include <VsiAnalysisResultsModule.h>
#include <VsiMeasurementModule.h>
#include <VsiImportExportModule.h>
#include <VsiImportExportXml.h>
#include <VsiAppViewModule.h>
#include <VsiEkvUtils.h>
#include "VsiStudyXml.h"
#include "VsiDeleteConfirmation.h"
#include "VsiChangeMsmntPackageDlg.h"
#include "VsiChangeStudyInfoDlg.h"
#include "VsiStudyBrowser.h"


// Commands that this view listen
#define VSI_CMD(_a) _a,
static const DWORD g_pStudyBrowserCmds[] =
{
	VSI_CMDS()
};


// Subclass list
// NOTE: Default values of column width and order came from Data/VsiConfigurations/UI.xml
//       Please refer to parameter name of "Study-Browser-Columns-Width" and "Study-Browser-Columns-Order"
static VSI_LVCOLUMN g_plvcStudyBrowserListView[] =
{
	{ LVCF_WIDTH | LVCF_TEXT, 0, 0, IDS_STYBRW_COLUMN_NAME, VSI_DATA_LIST_COL_NAME, -1, -1, TRUE },
	{ LVCF_WIDTH | LVCF_TEXT, 0, 18, IDS_STYBRW_COLUMN_LOCKED, VSI_DATA_LIST_COL_LOCKED, 0, -1, TRUE },
	{ LVCF_WIDTH | LVCF_TEXT, 0, 0, IDS_STYBRW_COLUMN_DATE, VSI_DATA_LIST_COL_DATE, -1, -1, TRUE },
	{ LVCF_WIDTH | LVCF_TEXT, 0, 0, IDS_STYBRW_COLUMN_STUDY_OWNER, VSI_DATA_LIST_COL_OWNER, -1, -1, TRUE },
	{ LVCF_WIDTH | LVCF_TEXT, 0, 0, IDS_STYBRW_COLUMN_ACQUIRED_BY, VSI_DATA_LIST_COL_ACQ_BY, -1, -1, TRUE },
	{ LVCF_WIDTH | LVCF_TEXT, 0, 0, IDS_STYBRW_COLUMN_MODE, VSI_DATA_LIST_COL_MODE, -1, -1, TRUE },
	{ LVCF_WIDTH | LVCF_TEXT, 0, 0, IDS_STYBRW_COLUMN_LENGTH, VSI_DATA_LIST_COL_LENGTH, -1, -1, TRUE },
	{ LVCF_WIDTH | LVCF_TEXT, 0, 0, IDS_STYBRW_COLUMN_ANIMAL_ID, VSI_DATA_LIST_COL_ANIMAL_ID, -1, -1, FALSE },
	{ LVCF_WIDTH | LVCF_TEXT, 0, 0, IDS_STYBRW_COLUMN_PROTOCOL_ID, VSI_DATA_LIST_COL_PROTOCOL_ID, -1, -1, FALSE },
};

CVsiStudyBrowser::CVsiStudyBrowser() :
	m_dwAppMode(VSI_APP_MODE_NONE),
	m_bStopSearch(FALSE),
	m_iTracker(0),
	m_iPreview(0),
	m_iBorder(0),
	m_bVertical(false),
	m_State(VSI_SB_STATE_START),
	m_typeSel(VSI_DATA_LIST_NONE),
	m_dwSink(0),
	m_dwEventImagePropUpdate(0),
	m_dwStudiesPathUpdate(0),
	m_dwStudiesAccessFilter(0),
	m_dwOperatorListUpdate(0),
	m_dwSecureModeChanged(0),
	m_ViewState(VSI_VIEW_STATE_HIDE)
{
	CVsiLogoViewBase<CVsiStudyBrowser>::SetSmallLogoWidth(1180);
}

CVsiStudyBrowser::~CVsiStudyBrowser()
{
	_ASSERT(m_pApp == NULL);
	_ASSERT(m_pPdm == NULL);
	_ASSERT(m_pStudyMgr == NULL);
	_ASSERT(m_pModeMgr == NULL);
	_ASSERT(m_pDataList == NULL);
	_ASSERT(m_pSession == NULL);
}

STDMETHODIMP CVsiStudyBrowser::OnPreStateUpdate(
	VSI_APP_STATE dwStateCurrent,
	VSI_APP_STATE dwState,
	BOOL bTryRelax,
	BOOL *pbBusy,
	const VARIANT *pvParam)
{
	UNREFERENCED_PARAMETER(dwStateCurrent);
	UNREFERENCED_PARAMETER(dwState);
	UNREFERENCED_PARAMETER(bTryRelax);
	UNREFERENCED_PARAMETER(pbBusy);
	UNREFERENCED_PARAMETER(pvParam);

	if (bTryRelax)
	{
		CancelSearchIfRunning();
	}

	return S_OK;
}

STDMETHODIMP CVsiStudyBrowser::OnStateUpdate(VSI_APP_STATE dwState, const VARIANT *pvParam)
{
	return HandleSetState(dwState, pvParam);
}

STDMETHODIMP CVsiStudyBrowser::Activate(
	IVsiApp *pApp,
	HWND hwndParent)
{
	HRESULT hr = S_OK;

	try
	{
		m_pApp = pApp;
		VSI_CHECK_INTERFACE(m_pApp, VSI_LOG_ERROR, NULL);

		hr = pApp->GetAppMode(&m_dwAppMode);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Cache App Controller
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_APP_CONTROLLER,
			__uuidof(IVsiAppController),
			(IUnknown**)&m_pAppController);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Register as a event listener
		hr = m_pAppController->AdviseStateUpdateSink(
			static_cast<IVsiAppStateSink*>(this));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Cache App View
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_APP_VIEW,
			__uuidof(IVsiAppView),
			(IUnknown**)&m_pAppView);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
		// Get PDM
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Command Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CMD_MGR,
			__uuidof(IVsiCommandManager),
			(IUnknown**)&m_pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Study Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_STUDY_MGR,
			__uuidof(IVsiStudyManager),
			(IUnknown**)&m_pStudyMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Mode Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_MODE_MANAGER,
			__uuidof(IVsiModeManager),
			(IUnknown**)&m_pModeMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Operator Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_OPERATOR_MGR,
			__uuidof(IVsiOperatorManager),
			(IUnknown**)&m_pOperatorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Advise as a sink for commands
		hr = m_pCmdMgr->AdviseCommandSink(
			g_pStudyBrowserCmds,
			_countof(g_pStudyBrowserCmds),
			VSI_CMD_SINK_TYPE_EXECUTE,
			this);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure advising command sink");

		// Get image prop update event id
		CComPtr<IVsiParameter> pParamEvent;
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
			VSI_PARAMETER_EVENTS_GENERAL_IMAGE_PROP_UPDATE,
			&pParamEvent);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParamEvent->GetId(&m_dwEventImagePropUpdate);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		pParamEvent.Release();
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_SYS_PATH_DATA,
			&pParamEvent);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParamEvent->GetId(&m_dwStudiesPathUpdate);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		pParamEvent.Release();
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_UI_STUDY_ACCESS_FILTER,
			&pParamEvent);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParamEvent->GetId(&m_dwStudiesAccessFilter);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		pParamEvent.Release();
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
			VSI_PARAMETER_EVENTS_GENERAL_OPERATOR_LIST_UPDATE,
			&pParamEvent);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParamEvent->GetId(&m_dwOperatorListUpdate);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		pParamEvent.Release();
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
			VSI_PARAMETER_SYS_SECURE_MODE,
			&pParamEvent);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParamEvent->GetId(&m_dwSecureModeChanged);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Events sink
		CComQIPtr<IVsiPdmEvents> pPdmEvents(m_pPdm);
		hr = pPdmEvents->AdviseParamUpdateEventSink(
			static_cast<IVsiParamUpdateEventSink*>(this), &m_dwSink);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Gets session
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_SESSION_MGR,
			__uuidof(IVsiSession),
			(IUnknown**)&m_pSession);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Gets cursor manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CURSOR_MANAGER,
			__uuidof(IVsiCursorManager),
			(IUnknown**)&m_pCursorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HWND hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
			VSI_ERROR_LOG(err);

		Deactivate();
	}

	return hr;
}

STDMETHODIMP CVsiStudyBrowser::Deactivate()
{
	HRESULT hr = S_FALSE;

	if (IsWindow())
	{
		BOOL bRet = DestroyWindow();
		_ASSERT(bRet);

		hr = S_OK;
	}

	m_pCursorMgr.Release();

	// Release AnalysisSet object
	if (NULL != m_pAnalysisSet)
	{
		m_pAnalysisSet->Uninitialize();
		m_pAnalysisSet.Release();
	}

	m_pDataList.Release();
	m_pSession.Release();

	// Un-advise
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

	m_pOperatorMgr.Release();
	m_pModeMgr.Release();
	m_pStudyMgr.Release();

	if (m_pCmdMgr != NULL)
	{
		// Un-advise as a sink for mode commands
		HRESULT hrCommandSink = m_pCmdMgr->UnadviseCommandSink(
			g_pStudyBrowserCmds,
			_countof(g_pStudyBrowserCmds),
			VSI_CMD_SINK_TYPE_EXECUTE,
			this);
		if (FAILED(hrCommandSink))
			VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure unadvising command sink");

		m_pCmdMgr.Release();
	}

	m_pPdm.Release();
	m_pAppView.Release();

	if (m_pAppController != NULL)
	{
		// Un-register event listener
		m_pAppController->UnadviseStateUpdateSink(
			static_cast<IVsiAppStateSink*>(this));

		m_pAppController.Release();
	}

	m_pApp.Release();

	return hr;
}

STDMETHODIMP CVsiStudyBrowser::GetWindow(HWND *phWnd)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(phWnd, VSI_LOG_ERROR, L"phWnd");

		*phWnd = m_hWnd;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyBrowser::GetAccelerator(HACCEL *phAccel)
{
	UNREFERENCED_PARAMETER(phAccel);
	return E_NOTIMPL;
}

STDMETHODIMP CVsiStudyBrowser::GetIsBusy(
	DWORD dwStateCurrent,
	DWORD dwState,
	BOOL bTryRelax,
	BOOL *pbBusy)
{
	UNREFERENCED_PARAMETER(bTryRelax);

	switch (dwStateCurrent)
	{
	case VSI_APP_STATE_STUDY_BROWSER:
		{
			switch (dwState)
			{
			case VSI_APP_STATE_DO_EXPORT_IMAGE:  // VevoStrain
				*pbBusy = FALSE;
				break;

			default:
				*pbBusy = !GetTopLevelParent().IsWindowEnabled();
				break;
			}
		}
		break;
		
	case VSI_APP_STATE_IMAGE_LABEL:
		{
			switch (dwState)
			{
			case VSI_APP_STATE_IMAGE_LABEL_CLOSE:
				*pbBusy = FALSE;
				break;
			}
		}	
		break;
	
	default:
		*pbBusy = !GetTopLevelParent().IsWindowEnabled();
		break;
	}

	return S_OK;
}

STDMETHODIMP CVsiStudyBrowser::OnViewStateChange(VSI_VIEW_STATE state)
{
	m_ViewState = state;

	return S_OK;
}

STDMETHODIMP CVsiStudyBrowser::PreTranslateMessage(
	MSG *pMsg, BOOL *pbHandled)
{
	switch (pMsg->message)
	{
	case WM_KEYDOWN:
		{
			// Handles key messages if the app has the focus
			// and the focus is no on the search field
			CWindow wndFocus(GetFocus());
			if ((wndFocus.GetTopLevelParent() == GetTopLevelParent()) &&
				(wndFocus.GetParent() != m_wndSearch))
			{
				switch (pMsg->wParam)
				{
				case VK_DELETE:
					{
						if (GetDlgItem(IDC_SB_DELETE).IsWindowEnabled())
						{
							BOOL bHandled(TRUE);
							OnBnClickedDelete(BN_CLICKED, IDC_SB_DELETE, NULL, bHandled);
						}
					}
					break;

				case VK_RETURN:
					{
						if (GetFocus() == GetDlgItem(IDC_SB_STUDY_LIST))
						{
							OpenItem();
						}
					}
					break;

				case L'A':
					{
						if (GetKeyState(VK_CONTROL) < 0)
						{
							if (GetFocus() == GetDlgItem(IDC_SB_STUDY_LIST))
							{
								CVsiLockData lock(m_pDataList);
								if (lock.LockData())
								{
									m_pDataList->SelectAll();
								}
							}
						}
					}
					break;
				}
			}
		}
		break;

	default:
		// Nothing to do
		break;
	}

	*pbHandled = IsDialogMessage(pMsg);

	return S_OK;
}

STDMETHODIMP CVsiStudyBrowser::SetSelection(LPCOLESTR pszSelection)
{
	m_strSelection = pszSelection;

	if (IsWindow())
	{
		UpdateSelection();
	}

	return S_OK;
}

STDMETHODIMP CVsiStudyBrowser::GetLastSelectedItem(
	VSI_DATA_LIST_TYPE *pType, IUnknown **ppUnk)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pType, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(ppUnk, VSI_LOG_ERROR, NULL);

		hr = m_pDataList->GetLastSelectedItem(pType, ppUnk);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyBrowser::Execute(
	DWORD dwCmd, const VARIANT *pvParam, BOOL *pbContinueExecute)
{
	HRESULT hr = S_OK;

	try
	{
		// Handle behind the scene
		hr = HandleCommand(dwCmd, pvParam, pbContinueExecute);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyBrowser::OnEvent(DWORD dwEvent, const VARIANT *pvParam)
{
	HRESULT hr = S_OK;

	try
	{
		if (VSI_STUDY_MGR_EVENT_UPDATE == dwEvent)
		{
			VSI_VERIFY(VT_UNKNOWN == V_VT(pvParam), VSI_LOG_ERROR, NULL);
			
			CComVariant *pvUpdate = new CComVariant(V_UNKNOWN(pvParam));
			PostMessage(WM_VSI_SB_UPDATE_ITEMS, (WPARAM)pvUpdate);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyBrowser::OnParamUpdate(DWORD *pdwParamIds, DWORD dwCount)
{
	HRESULT hr(S_OK);

	if (0 == m_dwSink)
		return hr;

	try
	{
		for (DWORD i = 0; i < dwCount; ++i, ++pdwParamIds)
		{
			if (*pdwParamIds == m_dwEventImagePropUpdate)
			{
				int iSel = ListView_GetNextItem(m_wndImageList, -1, LVNI_SELECTED);
				if (iSel >= 0)
				{
					LVITEMW lvi = { LVIF_PARAM };
					lvi.iItem = iSel;

					ListView_GetItem(m_wndImageList, &lvi);
					if (NULL != lvi.lParam)
					{
						IVsiImage *pImage = (IVsiImage*)lvi.lParam;

						CComVariant vName;
						pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vName);

						SYSTEMTIME stDate;
						CComVariant vDate;
						pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vDate);

						COleDateTime date(vDate);
						date.GetAsSystemTime(stDate);

						// The time is stored as UTC. Convert it to local time for display.
						SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

						WCHAR szText[500], szDate[50], szTime[50];
						GetDateFormat(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &stDate, NULL, szDate, _countof(szDate));
						GetTimeFormat(LOCALE_USER_DEFAULT, 0, &stDate, NULL, szTime, _countof(szTime));

						if (VT_EMPTY != V_VT(&vName) && (lstrlen(V_BSTR(&vName)) != 0))
							swprintf_s(szText, _countof(szText), L"%s\r\n%s %s", V_BSTR(&vName), szTime, szDate);
						else
							swprintf_s(szText, _countof(szText),  L"%s %s", szTime, szDate);

						lvi.mask = LVIF_TEXT;
						lvi.pszText = szText;

						ListView_SetItem(m_wndImageList, &lvi);
					}
				}
			}
			else if (*pdwParamIds == m_dwStudiesPathUpdate)
			{
				// Load study list
				SetStudyInfo(NULL, false);

				hr = FillStudyList();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else if (*pdwParamIds == m_dwStudiesAccessFilter)
			{
				// Update image
				UpdateStudyAccessFilterImage();

				// Load study list
				SetStudyInfo(NULL, false);

				hr = FillStudyList();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else if (*pdwParamIds == m_dwOperatorListUpdate)
			{
				BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
					VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);

				if (bSecureMode)
				{
					// Load study list
					SetStudyInfo(NULL, false);

					hr = FillStudyList();
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}
			else if (*pdwParamIds == m_dwSecureModeChanged)
			{
				// Refresh UI layout
				UpdateLogoutState();

				// Load study list
				SetStudyInfo(NULL, false);

				hr = FillStudyList();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}


LRESULT CVsiStudyBrowser::OnInitDialog(UINT, WPARAM, LPARAM, BOOL& bHandled)
{
	HRESULT hr = S_OK;
	bHandled = FALSE;

	try
	{
		// Fix the template
		{
			CRect rc;
			GetDlgItem(IDC_SB_NEW).GetWindowRect(&rc);
			ScreenToClient(&rc);

			GetDlgItem(IDC_SB_CLOSE).MoveWindow(&rc);
		}

		m_hExit.Attach(CreateEvent(NULL, FALSE, FALSE, L"VsiEventStudyListExit"));
		VSI_WIN32_VERIFY(NULL != m_hExit, VSI_LOG_ERROR, NULL);

		m_hReload.Attach(CreateEvent(NULL, FALSE, FALSE, L"VsiEventStudyListReload"));
		VSI_WIN32_VERIFY(NULL != m_hReload, VSI_LOG_ERROR, NULL);

		m_hSearch.Attach(CreateEvent(NULL, FALSE, FALSE, L"VsiEventStudyListSearch"));
		VSI_WIN32_VERIFY(NULL != m_hSearch, VSI_LOG_ERROR, NULL);

		DWORD dwState = VsiGetBitfieldValue<DWORD>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_STATE, m_pPdm);

		// Must follow the UI order from left to right
		static const DWORD pdwCtls[] = {
			IDC_SB_NEW,
			IDC_SB_CLOSE,
			IDC_SB_ANALYSIS,
			IDC_SB_SONO,
			IDC_SB_STRAIN,
			IDC_SB_STUDY_SERIES_INFO,
			IDC_SB_DELETE,
			IDC_SB_EXPORT,
			IDC_SB_COPY_TO,
			IDC_SB_COPY_FROM,
			IDCANCEL,
		};

		// More tools
		{
			CVsiToolsMore::Initialize(GetDlgItem(IDC_SB_TOOLS_MORE));

			for (int i = 0; i < _countof(pdwCtls); ++i)
			{
				if (0 == (VSI_SYS_STATE_HARDWARE & dwState))
				{
					if (IDC_SB_NEW == pdwCtls[i] || IDC_SB_CLOSE == pdwCtls[i])
					{
						continue;
					}
				}

				CVsiToolsMore::AddTool(GetDlgItem(pdwCtls[i]));
			}
		}

		// Search
		{
			CWindow wndDummy(GetDlgItem(IDC_SB_SEARCH));

			CRect rc;
			wndDummy.GetWindowRect(&rc);

			ScreenToClient(&rc);

			m_wndSearch.Create(
				*this,
				rc,
				L"Search",
				WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_TABSTOP,
				WS_EX_CLIENTEDGE,
				IDC_SB_SEARCH);

			// Set Z-order
			m_wndSearch.SetWindowPos(wndDummy, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);

			wndDummy.DestroyWindow();
		}

		// Access filter
		{
			CVsiWindow wndAccessFilter(GetDlgItem(IDC_SB_ACCESS_FILTER));
			wndAccessFilter.ModifyStyle(BS_TYPEMASK, BS_SPLITBUTTON);

			UpdateStudyAccessFilterImage();
		}

		// Help
		{
			m_ilHelp.CreateFromImage(
				_U_STRINGorID(IDB_HELP),
				20,
				1,
				CLR_DEFAULT,
				IMAGE_BITMAP,
				LR_CREATEDIBSECTION);

			BUTTON_IMAGELIST bil = { m_ilHelp, { 0, 0, 0, 0 }, BUTTON_IMAGELIST_ALIGN_CENTER };

			CWindow wndHelp(GetDlgItem(IDC_SB_HELP));

			wndHelp.SendMessage(BCM_SETIMAGELIST, 0 , reinterpret_cast<LPARAM>(&bil));
			wndHelp.SendMessage(
				RegisterWindowMessage(VSI_THEME_COMMAND),
				VSI_BT_CMD_SETTIP,
				(LPARAM)L"Help");
		}

		// Preferences
		{
			m_ilPreferences.CreateFromImage(
				_U_STRINGorID(IDB_PREFERENCES),
				20,
				1,
				CLR_DEFAULT,
				IMAGE_BITMAP,
				LR_CREATEDIBSECTION);

			BUTTON_IMAGELIST bil = { m_ilPreferences, { 0, 0, 0, 0 }, BUTTON_IMAGELIST_ALIGN_CENTER };

			CWindow wndPrefs(GetDlgItem(IDC_SB_PREFS));

			wndPrefs.SendMessage(BCM_SETIMAGELIST, 0 , reinterpret_cast<LPARAM>(&bil));
			wndPrefs.SendMessage(
				RegisterWindowMessage(VSI_THEME_COMMAND),
				VSI_BT_CMD_SETTIP,
				(LPARAM)L"Preferences");
		}

		// Shutdown
		{
			m_ilShutdown.CreateFromImage(
				_U_STRINGorID(IDB_LOGOUT_SHUTDOWN),
				21,
				1,
				CLR_DEFAULT,
				IMAGE_BITMAP,
				LR_CREATEDIBSECTION);

			BUTTON_IMAGELIST bil = { m_ilShutdown, { 0, 0, 0, 0 }, BUTTON_IMAGELIST_ALIGN_CENTER };

			CWindow wndShutdown(GetDlgItem(IDC_SB_SHUTDOWN));

			wndShutdown.SendMessage(BCM_SETIMAGELIST, 0 , reinterpret_cast<LPARAM>(&bil));
			wndShutdown.SendMessage(
				RegisterWindowMessage(VSI_THEME_COMMAND),
				VSI_BT_CMD_SETTIP,
				(LPARAM)L"Shut Down");
		}

		// Shutdown drop menu - system only
		if (0 < (VSI_SYS_STATE_HARDWARE & dwState))
		{
			CVsiWindow wndLogout(GetDlgItem(IDC_SB_LOGOUT));
			wndLogout.ModifyStyle(BS_TYPEMASK, BS_SPLITBUTTON);
		}

		// Set font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
		VsiThemeRecurSetFont(*this, hFont);

		// Layout
		hr = m_pLayout.CoCreateInstance(__uuidof(VsiWindowLayout));
		if (SUCCEEDED(hr))
		{
			hr = m_pLayout->Initialize(m_hWnd, VSI_WL_FLAG_AUTO_RELEASE);
			if (SUCCEEDED(hr))
			{
				for (int i = 0; i < _countof(pdwCtls); ++i)
				{
					m_pLayout->AddControl(0, pdwCtls[i], VSI_WL_MOVE_X);
				}

				m_pLayout->AddControl(0, IDC_SB_HELP, VSI_WL_MOVE_X);
				m_pLayout->AddControl(0, IDC_SB_PREFS, VSI_WL_MOVE_X);
				m_pLayout->AddControl(0, IDC_SB_SHUTDOWN, VSI_WL_MOVE_X);
				m_pLayout->AddControl(0, IDC_SB_LOGOUT, VSI_WL_MOVE_X);
			}
		}

		{
			CRect rc;
			GetDlgItem(IDC_SB_TRACK_BAR).GetWindowRect(&rc);
			m_iTracker = rc.Width();

			GetDlgItem(IDC_SB_PREVIEW).GetWindowRect(&rc);
			m_iPreview = rc.Width();

			CRect rcClient;
			GetClientRect(&rcClient);
			ScreenToClient(&rc);
			m_iBorder = rcClient.right - rc.right;
		}
		
		// Setup window branding
		SetWindowText(CString(MAKEINTRESOURCE(IDS_STYBRW_STUDY_BROWSER)).GetBuffer());			
		
		// Init data list
		{
			hr = m_pDataList.CoCreateInstance(__uuidof(VsiDataListWnd));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating Data List");

			CWindow wndStudyList(GetDlgItem(IDC_SB_STUDY_LIST));

			VSI_LVCOLUMN *plvcolumns(NULL);
			int iColumns(0);

			plvcolumns = g_plvcStudyBrowserListView;
			iColumns =  _countof(g_plvcStudyBrowserListView);

			// Restore column width - not critical
			CString strWidths = VsiGetRangeValue<CString>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_STUDY_BROWSER_COLUMNS_WIDTH, m_pPdm);

			if (strWidths.IsEmpty())
			{
				CComPtr<IVsiParameter> pParam;

				hr = m_pPdm->GetParameter(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_UI_STUDY_BROWSER_COLUMNS_WIDTH, &pParam);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				pParam->ResetValue();

				strWidths = VsiGetRangeValue<CString>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_UI_STUDY_BROWSER_COLUMNS_WIDTH, m_pPdm);
			}

			if (!strWidths.IsEmpty())
			{
				int iPos(0);
				CString strToken = strWidths.Tokenize(L";", iPos);
				while (!strToken.IsEmpty())
				{
					int iSubItem;
					int iWidth;
					int iRet = swscanf_s(strToken, L"%d,%d", &iSubItem, &iWidth);
					if (2 == iRet)
					{
						if ((iWidth > 0) && (iWidth < 1000))
						{
							for (int i = 0; i < iColumns; ++i)
							{
								if (iSubItem == plvcolumns[i].iSubItem)
								{
									plvcolumns[i].cx = iWidth;
									break;
								}
							}
						}
					}

					strToken = strWidths.Tokenize(L";", iPos);
				}
			}

			// Restore column order - not critical
			CString strOrders = VsiGetRangeValue<CString>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_STUDY_BROWSER_COLUMNS_ORDER, m_pPdm);

			if (strOrders.IsEmpty())
			{
				CComPtr<IVsiParameter> pParam;

				hr = m_pPdm->GetParameter(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_UI_STUDY_BROWSER_COLUMNS_ORDER, &pParam);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				pParam->ResetValue();

				strOrders = VsiGetRangeValue<CString>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_UI_STUDY_BROWSER_COLUMNS_ORDER, m_pPdm);
			}

			if (!strOrders.IsEmpty())
			{
				int iPos(0);
				CString strToken = strOrders.Tokenize(L";", iPos);
				while (!strToken.IsEmpty())
				{
					int iSubItem;
					int iOrder;
					int iRet = swscanf_s(strToken, L"%d,%d", &iSubItem, &iOrder);
					if (2 == iRet)
					{
						for (int i = 0; i < iColumns; ++i)
						{
							if (iSubItem == plvcolumns[i].iSubItem)
							{
								plvcolumns[i].iOrder = iOrder;
								plvcolumns[i].bVisible = 0 <= iOrder;
								break;
							}
						}
					}

					strToken = strOrders.Tokenize(L";", iPos);
				}

				// Make sure column 0 is valid (just in case)
				plvcolumns[0].bVisible = TRUE;
				if (0 > plvcolumns[0].iOrder)
				{
					plvcolumns[0].iOrder = 0;
				}
			}

			// Finally, initialize the data list
			hr = m_pDataList->Initialize(
				wndStudyList,
				VSI_DATA_LIST_FLAG_ITEM_STATUS_CALLBACK,
				VSI_DATA_LIST_STUDY,
				m_pApp,
				plvcolumns,
				iColumns);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pDataList->SetEmptyMessage(
				CString(MAKEINTRESOURCE(IDS_STYBRW_STUDY_EMPTYMSG)).GetBuffer());			
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			HFONT hFontNormal, hFontReviewed;
			VsiThemeGetFont(VSI_THEME_FONT_M, &hFontNormal);
			VsiThemeGetFont(VSI_THEME_FONT_M_T, &hFontReviewed);
			hr = m_pDataList->SetFont(hFontNormal, hFontReviewed);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pDataList->SetTextColor(
				VSI_DATA_LIST_ITEM_STATUS_ACTIVE_SINGLE,
				VSI_COLOR_ITEM_ACTIVE);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pDataList->SetTextColor(
				VSI_DATA_LIST_ITEM_STATUS_ACTIVE_LEFT,
				VSI_COLOR_ITEM_ACTIVE);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pDataList->SetTextColor(
				VSI_DATA_LIST_ITEM_STATUS_ACTIVE_RIGHT,
				VSI_COLOR_ITEM_ACTIVE);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pDataList->SetTextColor(
				VSI_DATA_LIST_ITEM_STATUS_ACTIVE,
				VSI_COLOR_ITEM_ACTIVE);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pDataList->SetTextColor(
				VSI_DATA_LIST_ITEM_STATUS_INACTIVE,
				VSI_COLOR_GRAY);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Sorting
			int iCol = VsiGetDiscreteValue<long>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_SORT_STUDY_ITEM_STUDYBROWSER, m_pPdm);
			int iOrder = VsiGetDiscreteValue<long>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_SORT_STUDY_ORDER_STUDYBROWSER, m_pPdm);

			m_pDataList->SetSortSettings(
				VSI_DATA_LIST_STUDY,
				(VSI_DATA_LIST_COL)iCol,
				VSI_SYS_SORT_ORDER_DESCENDING == iOrder);

			iCol = VsiGetDiscreteValue<long>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_SORT_SERIES_ITEM_STUDYBROWSER, m_pPdm);
			iOrder = VsiGetDiscreteValue<long>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_SORT_SERIES_ORDER_STUDYBROWSER, m_pPdm);

			m_pDataList->SetSortSettings(
				VSI_DATA_LIST_SERIES,
				(VSI_DATA_LIST_COL)iCol,
				VSI_SYS_SORT_ORDER_DESCENDING == iOrder);

			iCol = VsiGetDiscreteValue<long>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_SORT_IMAGE_ITEM_STUDYBROWSER, m_pPdm);
			iOrder = VsiGetDiscreteValue<long>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_SORT_IMAGE_ORDER_STUDYBROWSER, m_pPdm);

			m_pDataList->SetSortSettings(
				VSI_DATA_LIST_IMAGE,
				(VSI_DATA_LIST_COL)iCol,
				VSI_SYS_SORT_ORDER_DESCENDING == iOrder);
		}

		// Init study overview control
		m_studyOverview.SubclassWindow(GetDlgItem(IDC_SB_INFO));
		m_studyOverview.Initialize();

		// Checks active series from last instance (if in hardware mode)
		if (0 < (VSI_SYS_STATE_HARDWARE & dwState))
		{
			CComHeapPtr<OLECHAR> pszSeriesNs;
			hr = m_pSession->CheckActiveSeries(FALSE, NULL, &pszSeriesNs);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
			if (pszSeriesNs != NULL)
			{
				m_strReactivateSeriesNs = pszSeriesNs;
				// Report on fill
			}
		}

		m_wndImageList = GetDlgItem(IDC_SB_PREVIEW);
		ListView_SetExtendedListViewStyleEx(
			m_wndImageList,
			LVS_EX_TRACKSELECT,
			LVS_EX_TRACKSELECT);

		UpdateUI();
		UpdateLogoutState();

		// Listen for study manager events
		{
			CComQIPtr<IVsiEventSource> pEventSource(m_pStudyMgr);
			VSI_CHECK_INTERFACE(pEventSource, VSI_LOG_ERROR, NULL);
		
			hr = pEventSource->AdviseEventSink(static_cast<IVsiEventSink*>(this));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Preview width
		{
			m_iPreview = VsiGetRangeValue<int>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_STUDY_BROWSER_PREVIEW_WIDTH, m_pPdm);
		}

		RunThread(L"StudyList", THREAD_PRIORITY_NORMAL);

		// Sets focus
		PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)GetDlgItem(IDC_SB_STUDY_LIST), TRUE);

		// Delay the fill list thread a bit for faster UI
		PostMessage(WM_VSI_SB_FILL_LIST);
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
			VSI_ERROR_LOG(err);

		DestroyWindow();
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		StopPreview();

		m_bStopSearch = TRUE;
		CVsiThread::StopThread(m_hExit, INFINITE);

		if (NULL != m_pStudyMgr)
		{
			CComQIPtr<IVsiEventSource> pEventSource(m_pStudyMgr);
			if (NULL != pEventSource)
			{
				pEventSource->UnadviseEventSink(static_cast<IVsiEventSink*>(this));
			}
		}

		if (NULL != m_pLayout)
		{
			// Auto release
			m_pLayout.Detach();
		}

		m_studyOverview.UnsubclassWindow();
		m_studyOverview.Uninitialize();

		m_pDataList->Uninitialize();

		m_hExit.Close();
		m_hReload.Close();
		m_hSearch.Close();

		m_ilStudyAccess.Destroy();
		m_ilHelp.Destroy();
		m_ilPreferences.Destroy();
		m_ilShutdown.Destroy();
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnSettingChange(UINT, WPARAM, LPARAM, BOOL &bHandled)
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

	PostMessage(WM_VSI_SB_UPDATE_UI);

	return 0;
}

LRESULT CVsiStudyBrowser::OnSize(UINT, WPARAM, LPARAM lp, BOOL &bHandled)
{
	bHandled = FALSE;

	// Is initialized?
	if (0 < m_iTracker)
	{
		CRect rcClient;
		GetClientRect(&rcClient);

		CRect rcAccessFilter;
		CWindow wndAccessFilter(GetDlgItem(IDC_SB_ACCESS_FILTER));
		wndAccessFilter.GetWindowRect(&rcAccessFilter);
		ScreenToClient(&rcAccessFilter);

		CRect rcSearch;
		CWindow wndSearch(GetDlgItem(IDC_SB_SEARCH));
		wndSearch.GetWindowRect(&rcSearch);
		ScreenToClient(&rcSearch);

		int iGap = rcSearch.left - rcAccessFilter.right;

		CRect rcHelp;
		CWindow wndHelp(GetDlgItem(IDC_SB_HELP));
		wndHelp.GetWindowRect(&rcHelp);
		ScreenToClient(&rcHelp);

		CRect rcList;
		CWindow wndList(GetDlgItem(IDC_SB_STUDY_LIST));
		wndList.GetWindowRect(&rcList);
		ScreenToClient(&rcList);

		CRect rcPreview;
		m_wndImageList.GetWindowRect(&rcPreview);
		ScreenToClient(&rcPreview);

		CRect rcTrackbar;
		CWindow wndTrackbar(GetDlgItem(IDC_SB_TRACK_BAR));
		wndTrackbar.GetWindowRect(&rcTrackbar);
		ScreenToClient(&rcTrackbar);

		bool bVertical = rcClient.Height() > rcClient.Width();

		if (bVertical != m_bVertical)
		{
			wndTrackbar.ModifyStyle(bVertical ? 0 : VSI_TRACK_BAR_HORZ, bVertical ? VSI_TRACK_BAR_HORZ : 0);
			m_wndImageList.ModifyStyle(bVertical ? 0 : LVS_ALIGNLEFT, bVertical ? LVS_ALIGNLEFT : 0);

			m_bVertical = bVertical;

			// Delay redraw since the window is moving soon
			m_studyOverview.SetLayout(!m_bVertical, false);
		}

		if (bVertical)
		{
			rcList.right = rcClient.right - m_iBorder;
			if (rcList.right < rcList.left)
			{
				rcList.right = rcList.left;
			}
			rcList.bottom = rcClient.bottom - m_iBorder - m_iPreview - m_iTracker;
			if (rcList.bottom < rcList.top + 32)
			{
				m_iPreview -= rcList.top + 32 - rcList.bottom;
				if (m_iPreview < 0)
				{
					m_iPreview = 0;
				}
				rcList.bottom = rcList.top + 32;
			}

			rcPreview.left = rcList.left;
			rcPreview.top = rcList.bottom + m_iTracker;
			rcPreview.right = rcList.right;
			rcPreview.bottom = rcClient.bottom - m_iBorder;
			if (rcPreview.bottom < rcPreview.top)
			{
				rcPreview.bottom = rcPreview.top;
			}

			rcTrackbar.left = rcList.left;
			rcTrackbar.top = rcList.bottom;
			rcTrackbar.right = rcList.right;
			rcTrackbar.bottom = rcPreview.top;
		}
		else
		{
			rcList.right = rcClient.right - m_iBorder - m_iPreview - m_iTracker;
			if (rcList.right < rcList.left + 32)
			{
				m_iPreview -= rcList.left + 32 - rcList.right;
				if (m_iPreview < 0)
				{
					m_iPreview = 0;
				}
				rcList.right = rcList.left + 32;
			}
			rcList.bottom = rcClient.bottom - m_iBorder;
			if (rcList.bottom < rcList.top)
			{
				rcList.bottom = rcList.top;
			}

			rcPreview.left = rcList.right + m_iTracker;
			rcPreview.top = rcList.top;
			rcPreview.right = rcClient.right - m_iBorder;
			if (rcPreview.right < rcPreview.left)
			{
				rcPreview.right = rcPreview.left;
			}
			rcPreview.bottom = rcList.bottom;

			rcTrackbar.left = rcList.right;
			rcTrackbar.top = rcList.top;
			rcTrackbar.right = rcPreview.left;
			rcTrackbar.bottom = rcList.bottom;
		}

		rcSearch.OffsetRect(rcList.right - rcSearch.right, 0);
		if (rcSearch.right > rcHelp.left - iGap)
		{
			rcSearch.OffsetRect(rcHelp.left - rcSearch.right - iGap, 0);
		}
		else
		{
			if (!m_bVertical)
			{
				// For small screen, align search to the right
				if (rcSearch.Width() > (rcList.Width() / 2))
				{
					rcSearch.OffsetRect(rcHelp.left - rcSearch.right - iGap, 0);
				}
			}
			else
			{
				rcSearch.OffsetRect(rcHelp.left - rcSearch.right - iGap, 0);
			}
		}

		rcAccessFilter.OffsetRect(
			rcSearch.left - rcAccessFilter.Width() - rcAccessFilter.left - iGap, 
			rcSearch.top - rcAccessFilter.top);

		HDWP hdwp = BeginDeferWindowPos(6);

		hdwp = wndAccessFilter.DeferWindowPos(
			hdwp, NULL,
			rcAccessFilter.left, rcAccessFilter.top,
			rcAccessFilter.Width(), rcAccessFilter.Height(),
			SWP_NOZORDER | SWP_NOSIZE);
		hdwp = wndSearch.DeferWindowPos(
			hdwp, NULL,
			rcSearch.left, rcSearch.top,
			rcSearch.Width(), rcSearch.Height(),
			SWP_NOZORDER);
		hdwp = wndList.DeferWindowPos(
			hdwp, NULL,
			rcList.left, rcList.top,
			rcList.Width(), rcList.Height(),
			SWP_NOZORDER | SWP_NOMOVE);
		hdwp = m_wndImageList.DeferWindowPos(
			hdwp, NULL,
			rcPreview.left, rcPreview.top,
			rcPreview.Width(), rcPreview.Height(),
			SWP_NOZORDER);
		hdwp = GetDlgItem(IDC_SB_INFO).DeferWindowPos(
			hdwp, NULL,
			rcPreview.left, rcPreview.top,
			rcPreview.Width(), rcPreview.Height(),
			SWP_NOZORDER);
		hdwp = wndTrackbar.DeferWindowPos(
			hdwp, NULL,
			rcTrackbar.left, rcTrackbar.top,
			rcTrackbar.Width(), rcTrackbar.Height(),
			SWP_NOZORDER);

		EndDeferWindowPos(hdwp);

		int iMin(0);
		int iMax(0);

		int iPreviewMin = VsiGetMinValue<int>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_STUDY_BROWSER_PREVIEW_WIDTH, m_pPdm);
		int iPreviewMax = VsiGetMaxValue<int>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_STUDY_BROWSER_PREVIEW_WIDTH, m_pPdm);

		if (bVertical)
		{
			int iH = HIWORD(lp);
			iMin = iH - iPreviewMax;
			if (iMin < rcList.top)
				iMin = rcList.top;
			iMax = iH - iPreviewMin;
			if (iMax < iMin)
				iMax = iMin;
			if (iMax > rcPreview.bottom - rcTrackbar.Height())
				iMax = rcPreview.bottom - rcTrackbar.Height();
		}
		else
		{
			int iW = LOWORD(lp);
			iMin = iW - iPreviewMax;
			if (iMin < rcList.left)
				iMin = rcList.left;
			iMax = iW - iPreviewMin;
			if (iMax < iMin)
				iMax = iMin;
			if (iMax > rcPreview.right - rcTrackbar.Width())
				iMax = rcPreview.right - rcTrackbar.Width();
		}

		GetDlgItem(IDC_SB_TRACK_BAR).SendMessage(
			WM_VSI_TRACKBAR_SETMINMAX,
			MAKEWPARAM(iMin, iMax));

		if (bVertical)
		{
			if (rcTrackbar.top < rcList.top || rcTrackbar.bottom > rcPreview.bottom)
			{
				int iDefault = VsiGetRangeValueDefault<int>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_UI_STUDY_BROWSER_PREVIEW_WIDTH, m_pPdm);
				SetPreviewSize(iDefault);

				GetDlgItem(IDC_SB_PREVIEW).SendMessage(LVM_REDRAWITEMS);
			}
		}
		else
		{
			if (rcTrackbar.left < rcList.left || rcTrackbar.right > rcPreview.right)
			{
				int iDefault = VsiGetRangeValueDefault<int>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_UI_STUDY_BROWSER_PREVIEW_WIDTH, m_pPdm);
				SetPreviewSize(iDefault);

				GetDlgItem(IDC_SB_PREVIEW).SendMessage(LVM_REDRAWITEMS);
			}
		}

		CVsiToolsMore::Update();

		RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_ALLCHILDREN);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnWindowPosChanged(UINT, WPARAM, LPARAM lp, BOOL &bHandled)
{
	WINDOWPOS *pPos = (WINDOWPOS*)lp;

	// Updates preview list when turn visible
	HRESULT hr = S_OK;

	try
	{
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

			// Delay the update as long as possible
			MSG msg;
			PeekMessage(&msg, *this, WM_VSI_SB_UPDATE_UI, WM_VSI_SB_UPDATE_UI, TRUE);
			PostMessage(WM_VSI_SB_UPDATE_UI);

			// Sets focus to study list
			PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)GetDlgItem(IDC_SB_STUDY_LIST), TRUE);
			
			hr = m_pCursorMgr->SetState(VSI_CURSOR_STATE_ON, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (SWP_HIDEWINDOW & pPos->flags)
		{
			CancelSearchIfRunning();

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

LRESULT CVsiStudyBrowser::OnContextMenu(UINT, WPARAM wp, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	HMENU hMenu(NULL);

	try
	{
		CWindow wndStudyList(GetDlgItem(IDC_SB_STUDY_LIST));
	
		if (wndStudyList == reinterpret_cast<HWND>(wp))
		{
			DWORD dwPos = GetMessagePos();
			POINT pt = { GET_X_LPARAM(dwPos), GET_Y_LPARAM(dwPos) };

			wndStudyList.ScreenToClient(&pt);
			CWindow wndTest(RealChildWindowFromPoint(wndStudyList, pt));
			if (wndTest == ListView_GetHeader(wndStudyList))
			{
				hr = UpdateColumnsRecord();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hMenu = CreatePopupMenu();
				VSI_WIN32_VERIFY(NULL != hMenu, VSI_LOG_ERROR, NULL);

				VSI_LVCOLUMN *plvcolumns(NULL);
				int iColumns(0);

				plvcolumns = g_plvcStudyBrowserListView;
				iColumns =  _countof(g_plvcStudyBrowserListView);

				int iVisible(0);

				MENUITEMINFO mii = { sizeof(MENUITEMINFO), MIIM_ID | MIIM_STATE | MIIM_STRING };
				CString strLabel;

				// Go through visible first
				for (int i = 0; i < iColumns; ++i)
				{
					if (plvcolumns[i].bVisible)
					{
						++iVisible;
					}

					for (int j = 0; j < iColumns; ++j)
					{
						if (plvcolumns[j].bVisible && plvcolumns[j].iOrder == i)
						{
							if (VSI_DATA_LIST_COL_NAME == plvcolumns[j].iSubItem)
							{
								// Cannot remove name
								mii.fState = MFS_CHECKED | MFS_GRAYED;
							}
							else
							{
								mii.fState = MFS_CHECKED | MFS_ENABLED;
							}

							mii.wID = plvcolumns[j].iSubItem;

							strLabel.LoadString(plvcolumns[j].dwTextId);

							mii.dwTypeData = const_cast<LPWSTR>(strLabel.GetString());

							BOOL bInsert = InsertMenuItem(hMenu, static_cast<UINT>(-1), TRUE, &mii);
							VSI_WIN32_VERIFY(bInsert, VSI_LOG_ERROR, NULL);
						}
					}
				}

				// Go through invisible
				for (int i = 0; i < iColumns; ++i)
				{
					if (!plvcolumns[i].bVisible)
					{
						mii.fState = MFS_UNCHECKED | MFS_ENABLED;
						mii.wID = plvcolumns[i].iSubItem;

						strLabel.LoadString(plvcolumns[i].dwTextId);

						mii.dwTypeData = const_cast<LPWSTR>(strLabel.GetString());

						BOOL bInsert = InsertMenuItem(hMenu, iColumns, TRUE, &mii);
						VSI_WIN32_VERIFY(bInsert, VSI_LOG_ERROR, NULL);
					}
				}

				POINT ptMenu;
				GetCursorPos(&ptMenu);

				int iCmd = TrackPopupMenu(
					hMenu, TPM_RIGHTBUTTON | TPM_NONOTIFY | TPM_RETURNCMD,
					ptMenu.x, ptMenu.y, 0, GetTopLevelParent(), NULL);
				if (iCmd > 0)
				{
					int iRemoved(-1);

					for (int i = 0; i < iColumns; ++i)
					{
						if (plvcolumns[i].iSubItem == iCmd)
						{
							plvcolumns[i].bVisible = !plvcolumns[i].bVisible;
							if (plvcolumns[i].bVisible)
							{
								// Show new column at the end
								plvcolumns[i].iOrder = iVisible;
							}
							else
							{
								iRemoved = plvcolumns[i].iOrder;
								plvcolumns[i].iOrder = -1;
							}
							break;
						}
					}

					// When hiding column, repack order values
					if (0 <= iRemoved)
					{
						for (int i = iRemoved + 1; i < iVisible; ++i)
						{
							for (int j = 0; j < iColumns; ++j)
							{
								if (i == plvcolumns[j].iOrder)
								{
									--plvcolumns[j].iOrder;
									break;
								}
							}
						}
					}

					hr = m_pDataList->SetColumns(plvcolumns, iColumns);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				// Save column order
				CString strPair;
				CString strOrders;
				for (int i = 0; i < iColumns; ++i)
				{
					strPair.Format((0 == i) ? L"%d,%d" : L";%d,%d", plvcolumns[i].iSubItem, plvcolumns[i].iOrder);

					strOrders += strPair;
				}

				VsiSetRangeValue<LPCWSTR>(
					strOrders,
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_UI_STUDY_BROWSER_COLUMNS_ORDER, m_pPdm);
			}
		}
	}
	VSI_CATCH(hr);

	if (NULL != hMenu)
	{
		DestroyMenu(hMenu);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnFillList(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		hr = FillStudyList();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnFillListDone(UINT, WPARAM wp, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			// Sets selection
			UpdateSelection();

			CString strMessage;
			switch (wp)
			{
			case VSI_LIST_RELOAD:
				{
					strMessage.LoadString(IDS_STYBRW_NO_STUDY_DATA_AVAILABLE);
				}
				break;
			case VSI_LIST_SEARCH:
				{
					m_wndSearch.SetSearchAnimation(false);
					if (m_strSearch.IsEmpty())
					{
						strMessage.LoadString(IDS_STYBRW_NO_STUDY_DATA_AVAILABLE);
					}
					else
					{
						strMessage.Format(IDS_STYBRW_SEARCH_NO_MATCH, m_strSearch);
					}

					m_studyOverview.SetTextHighlight(m_strSearch);
				}
				break;
			default:
				_ASSERT(0);
			}

			// Change empty message
			hr = m_pDataList->SetEmptyMessage(strMessage);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Recover active series from last session
			if (!m_strReactivateSeriesNs.IsEmpty())
			{
				CComQIPtr<IVsiSeries> pSeries;
				CComPtr<IVsiStudy> pStudy;
				VSI_SERIES_ERROR dwErrorSeries(VSI_SERIES_ERROR_NONE);
				VSI_STUDY_ERROR dwErrorStudy(VSI_STUDY_ERROR_NONE);

				CVsiLockData lock(m_pDataList);
				if (lock.LockData())
				{
					// Gets series
					hr = m_pDataList->SetSelection(m_strReactivateSeriesNs, TRUE);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						CComPtr<IUnknown> pUnkItem;
						VSI_DATA_LIST_TYPE type;
						hr = m_pDataList->GetLastSelectedItem(
							&type, (IUnknown**)&pUnkItem);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						pSeries = pUnkItem;

						if (NULL != pSeries)
						{
							hr = pSeries->GetErrorCode(&dwErrorSeries);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							hr = pSeries->GetStudy(&pStudy);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					
							if (NULL != pStudy)
							{
								hr = pStudy->GetErrorCode(&dwErrorStudy);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							}
						}
					}
				}
			
				if (NULL != pStudy && (VSI_STUDY_ERROR_NONE == dwErrorStudy || VSI_STUDY_ERROR_LOAD_SERIES == dwErrorStudy) && 
					NULL != pSeries && (VSI_SERIES_ERROR_NONE == dwErrorSeries || VSI_SERIES_ERROR_LOAD_IMAGES == dwErrorSeries))
				{
					BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
						VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);

					int iRet = IDYES;

					CComVariant vAcqBy;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vAcqBy);
					if ((S_OK != hr) || (0 == *V_BSTR(&vAcqBy)))
					{
						VsiMessageBox(
							GetTopLevelParent(),
							CString(MAKEINTRESOURCE(bSecureMode ? IDS_STYBRW_SERIES_DID_NOT_CLOSE : IDS_STYBRW_RECENT_OP_DID_NOT_CLOSE)),
							VsiGetRangeValue<CString>(
								VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
								VSI_PARAMETER_SYS_PRODUCT_NAME,
								m_pPdm),
							MB_OK | MB_ICONINFORMATION);
					}
					else
					{
						iRet = VsiMessageBox(
							GetTopLevelParent(),
							CString(MAKEINTRESOURCE(bSecureMode ? IDS_STYBRW_SERIES_DID_NOT_CLOSE2 : IDS_STYBRW_RECENT_OP_DID_NOT_CLOSE2)),
							VsiGetRangeValue<CString>(
								VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
								VSI_PARAMETER_SYS_PRODUCT_NAME,
								m_pPdm),
							MB_YESNO | MB_ICONQUESTION);
					}

					if (IDYES == iRet)
					{
						// Setup session
						hr = m_pSession->SetStudy(0, pStudy, TRUE);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						hr = m_pSession->SetPrimaryStudy(pStudy, TRUE);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						hr = m_pSession->SetSeries(0, pSeries, TRUE);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						hr = m_pSession->SetPrimarySeries(pSeries, TRUE);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						// Sets acquisition session bit
						CVsiBitfield<ULONG> pState;
						hr = m_pPdm->GetParameter(
							VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
							VSI_PARAMETER_SYS_STATE, &pState);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						pState.SetBits(VSI_SYS_STATE_ACQ_SESSION, VSI_SYS_STATE_ACQ_SESSION);

						// Refresh study and series state
						hr = m_pDataList->UpdateItem(NULL, TRUE);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					
						BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
							VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
							VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
						if (!bSecureMode)
						{
							// Select acquired by as the current operator
							CComVariant vOperator;
							hr = pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vOperator);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						
							if ((S_OK == hr) && (0 != *V_BSTR(&vOperator)))
							{
								CComPtr<IVsiOperator> pOperator;
								hr = m_pOperatorMgr->GetOperator(
									VSI_OPERATOR_PROP_NAME,
									V_BSTR(&vOperator),
									&pOperator);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								if (S_OK == hr)
								{
									CComHeapPtr<OLECHAR> pszOperId;
									hr = pOperator->GetId(&pszOperId);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									hr = m_pOperatorMgr->SetCurrentOperator(pszOperId);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								}
							}
						}
					}
					else
					{
						// Remove active series record
						hr = m_pSession->CheckActiveSeries(TRUE, NULL, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}
				else
				{
					// The NS don't exists at all - nothing we can do 
					// Remove active series record
					hr = m_pSession->CheckActiveSeries(TRUE, NULL, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				// Done
				m_strReactivateSeriesNs.Empty();
			}

			GetDlgItem(IDC_SB_ACCESS_FILTER).EnableWindow(TRUE);
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnUpdateUI(UINT, WPARAM, LPARAM, BOOL&)
{
	CVsiLockData lock(m_pDataList);

	if (lock.LockData())
	{
		UpdateUI();
	}
	else
	{
		PostMessage(WM_VSI_SB_UPDATE_UI);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnUpdateItems(UINT, WPARAM wp, LPARAM lp, BOOL&)
{
	CVsiLockData lock(m_pDataList);
	if (lock.LockData())
	{
		CCritSecLock csl(m_csUpdateItems);

		if (0 != wp)
		{
			HRESULT hr(S_OK);
			CComVariant *pvParam = (CComVariant*)wp;

			try
			{
				// Clear preview list
				if (m_pItemPreview != NULL)
				{
					if (VSI_DATA_LIST_STUDY != m_typeSel)
					{
						StopPreview(VSI_PREVIEW_LIST_CLEAR);
					}

					m_pItemPreview.Release();
					m_typeSel = VSI_DATA_LIST_NONE;
				}

				SetStudyInfo(NULL);

				// Update list
				CComQIPtr<IXMLDOMDocument> pXmlDoc(V_UNKNOWN(pvParam));
				VSI_CHECK_INTERFACE(pXmlDoc, VSI_LOG_ERROR, L"Failure getting Update XML document");

				hr = m_pDataList->UpdateList(pXmlDoc);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			VSI_CATCH(hr);

			delete pvParam;
		}
		else
		{
			CVsiListUpdateItemsIter iter;

			for (iter = m_listUpdateItems.begin();
				iter != m_listUpdateItems.end();
				++iter)
			{
				const CString &itemNs = *iter;
				if (itemNs.IsEmpty())
					m_pDataList->UpdateItem(NULL, TRUE);  // All
				else
					m_pDataList->UpdateItem((LPCWSTR)itemNs, FALSE);
			}

			m_listUpdateItems.clear();
		}
	}
	else
	{
		PostMessage(WM_VSI_SB_UPDATE_ITEMS, wp, lp);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedMore(WORD, WORD, HWND, BOOL&)
{
	CVsiToolsMore::ShowDropdown();

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Cancel");

	switch (m_State)
	{
	case VSI_SB_STATE_CANCEL:
		m_pCmdMgr->InvokeCommand(ID_CMD_BROWSE, NULL);
		break;
	default:
		_ASSERT(0);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedHelp(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Help");

	m_pCmdMgr->InvokeCommand(ID_CMD_HELP, NULL);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedPrefs(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Preference");
	HRESULT hr = S_OK;

	CancelSearchIfRunning();

	try
	{
		hr = m_pCmdMgr->InvokeCommand(ID_CMD_PREFERENCES, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedClose(WORD, WORD, HWND, BOOL&)
{
	// Check button state - this command can come from the hot key
	if (!GetDlgItem(IDC_SB_CLOSE).IsWindowEnabled())
		return 0;

	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Close");
	HRESULT hr = S_OK;

	CancelSearchIfRunning();

	try
	{
		hr = m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_SERIES, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedInfo(WORD, WORD, HWND, BOOL&)
{
	// Check button state - this command can come from the hot key
	if (!GetDlgItem(IDC_SB_STUDY_SERIES_INFO).IsWindowEnabled())
		return 0;

	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Info");
	HRESULT hr(S_OK);

	CancelSearchIfRunning();

	try
	{
		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			// Gets focused item
			int iFocused(-1);
			hr = m_pDataList->GetSelectedItem(&iFocused);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (iFocused >= 0)
			{
				CComPtr<IUnknown> pUnkItem;
				VSI_DATA_LIST_TYPE type;
				hr = m_pDataList->GetItem(iFocused, &type, &pUnkItem);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				switch (type)
				{
				case VSI_DATA_LIST_STUDY:
					{
						CComVariant vParam((IUnknown*)pUnkItem);
						m_pCmdMgr->InvokeCommand(ID_CMD_STUDY_SERIES_INFO, &vParam);
					}
					break;
				case VSI_DATA_LIST_SERIES:
					{
						CComVariant vParam((IUnknown*)pUnkItem);
						m_pCmdMgr->InvokeCommand(ID_CMD_STUDY_SERIES_INFO, &vParam);
					}
					break;
				case VSI_DATA_LIST_IMAGE:
					{
						CComQIPtr<IVsiImage> pImage(pUnkItem);

						CComPtr<IVsiSeries> pSeries;
	
						hr = pImage->GetSeries(&pSeries);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComVariant vParam((IUnknown*)pSeries);
						m_pCmdMgr->InvokeCommand(ID_CMD_STUDY_SERIES_INFO, &vParam);
					}
					break;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

STDMETHODIMP CVsiStudyBrowser::DeleteImage(IVsiImage *pImage)
{
	HRESULT hr = S_OK;

	try
	{
		CComPtr<IVsiStudy> pStudy;
		hr = pImage->GetStudy(&pStudy);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vLocked(false);
		hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK == hr && VARIANT_FALSE != V_BOOL(&vLocked))
		{
			VsiMessageBox(
				GetTopLevelParent(),
				CString(MAKEINTRESOURCE(IDS_STYBRW_IMAGE_LOCKED_CANNOT_DELETE)),
				VsiGetRangeValue<CString>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_SYS_PRODUCT_NAME,
					m_pPdm),
				MB_OK | MB_ICONINFORMATION);

			return S_FALSE;
		}

		hr = m_pAppController->SetState(VSI_APP_STATE_DELETE_CONFIRMATION, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CVsiDeleteConfirmation dlg(NULL, 0, NULL, 0, (IUnknown**)&pImage, (int)1);
		INT_PTR rtn = dlg.DoModal(GetTopLevelParent());
		if (IDYES == rtn)
		{
			hr = m_pAppController->RemoveState(VSI_APP_STATE_DELETE_CONFIRMATION, TRUE);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// get browse direction here before image is closed
			int iSlot = 0;
			hr = m_pSession->GetActiveSlot(&iSlot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IUnknown> pUnkMode;
			HRESULT hrMode = m_pSession->GetMode(iSlot, &pUnkMode);
			VSI_CHECK_HR(hrMode, VSI_LOG_ERROR, NULL);

			CComVariant vParam(pImage);
			hr = m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_IMAGE, &vParam);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					
			// Updates cancel button state
			{
				int iImage(0);
				for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
				{
					CComPtr<IUnknown> pUnkView;
					hr = m_pSession->GetModeView(i, &pUnkView);
					if (S_OK == hr)
						++iImage;
				}
						
				if (0 == iImage)
				{
					m_State = VSI_SB_STATE_START;
					GetDlgItem(IDCANCEL).EnableWindow(FALSE);
				}
			}
				
			if (m_pItemPreview != NULL)
			{
				StopPreview(VSI_PREVIEW_LIST_CLEAR);
			}

			// XML for UI update
			CComPtr<IXMLDOMDocument> pXmlDoc;
			hr = VsiCreateDOMDocument(&pXmlDoc);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating DOMDocument");

			// Root element
			CComPtr<IXMLDOMElement> pElmUpdate;
			hr = pXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_UPDATE), &pElmUpdate);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure creating '%s' element", VSI_UPDATE_XML_ELM_UPDATE));

			hr = pXmlDoc->appendChild(pElmUpdate, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IXMLDOMElement> pElmRemove;
			hr = pXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_REMOVE), &pElmRemove);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure creating '%s' element", VSI_UPDATE_XML_ELM_REMOVE));

			// Set number of studies
			CComVariant vNum(0);
			hr = pElmRemove->setAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_STUDIES), vNum);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failed to set attribute (%s)", VSI_REMOVE_XML_ATT_NUMBER_STUDIES));

			// Set number of series
			vNum = 0;
			hr = pElmRemove->setAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_SERIES), vNum);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failed to set attribute (%s)", VSI_REMOVE_XML_ATT_NUMBER_SERIES));

			// Set number of images
			vNum = 1;
			hr = pElmRemove->setAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_IMAGES), vNum);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failed to set attribute (%s)", VSI_REMOVE_XML_ATT_NUMBER_IMAGES));

			hr = pElmUpdate->appendChild(pElmRemove, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
			hr = Delete(pImage, pXmlDoc, pElmRemove);
			if (S_OK != hr)
			{
				VSI_LOG_MSG(VSI_LOG_WARNING, L"Errors occurred deleting image items");

				CString strMsg;

				DWORD dwState = VsiGetBitfieldValue<DWORD>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_SYS_STATE, m_pPdm);
				if (0 < (VSI_SYS_STATE_HARDWARE & dwState))
				{
					// Hardware
					strMsg.LoadString(IDS_STYBRW_ERROR_WHILE_DELETE_STUDY1);
				}
				else
				{
					// Workstation
					strMsg.LoadString(IDS_STYBRW_ERROR_WHILE_DELETE_STUDY2);
				}

				WCHAR szSupport[500];
				VsiGetTechSupportInfo(szSupport, _countof(szSupport));

				strMsg += szSupport;

				VsiMessageBox(
					GetTopLevelParent(),
					strMsg,
					VsiGetRangeValue<CString>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
						VSI_PARAMETER_SYS_PRODUCT_NAME,
						m_pPdm),
					MB_OK | MB_ICONEXCLAMATION);
			}

			UpdateDataList(pXmlDoc);
		}
		else
		{
			hr = m_pAppController->RemoveState(VSI_APP_STATE_DELETE_CONFIRMATION, TRUE);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Cancel
			hr = S_FALSE;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

LRESULT CVsiStudyBrowser::OnBnClickedDelete(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Delete");
	HRESULT hr = S_OK;

	// Check button state - this command can come from the hot key
	if (!GetDlgItem(IDC_SB_DELETE).IsWindowEnabled())
		return 0;

	CancelSearchIfRunning();

	try
	{
		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			// Make sure the selection record is up-to-date
			int iStudyLoaded = 0;
			int iSeriesLoaded = 0;
			int iImageLoaded = 0;
			int iStudyCount = 0;
			int iSeriesCount = 0;
			int iImageCount = 0;
			int iImageCountAll = 0;
			CComPtr<IVsiEnumStudy> pEnumStudy;
			CComPtr<IVsiEnumSeries> pEnumSeries;
			CComPtr<IVsiEnumImage> pEnumImage;
			CComPtr<IVsiEnumImage> pEnumImageAll;
			VSI_DATA_LIST_COLLECTION listSel;

			CComPtr<IVsiSeries> pSeriesActive;
			hr = m_pSession->GetPrimarySeries(&pSeriesActive);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			memset(&listSel, 0, sizeof(listSel));
			listSel.dwFlags = VSI_DATA_LIST_FILTER_SELECT_FOR_DELETE;
			listSel.ppEnumStudy = &pEnumStudy;
			listSel.piStudyCount = &iStudyCount;
			listSel.ppEnumSeries = &pEnumSeries;
			listSel.piSeriesCount = &iSeriesCount;
			listSel.ppEnumImage = &pEnumImage;
			listSel.piImageCount = &iImageCount;
			listSel.pSeriesActive = pSeriesActive;
			hr = m_pDataList->GetSelection(&listSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			memset(&listSel, 0, sizeof(listSel));
			listSel.dwFlags = VSI_DATA_LIST_FILTER_MOVE_TO_CHILDREN;
			listSel.ppEnumImage = &pEnumImageAll;
			listSel.piImageCount = &iImageCountAll;
			hr = m_pDataList->GetSelection(&listSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			typedef std::vector<IUnknown*> CVsiVecDeleteItems;
			typedef CVsiVecDeleteItems::iterator CVsiVecDeleteItemsIter;
			CVsiVecDeleteItems vecStudyUnlocked;
			CVsiVecDeleteItems vecSeriesUnlocked;
			CVsiVecDeleteItems vecImageUnlocked;
			CVsiVecDeleteItems vecItemsLocked;

			// Process all images
			if (pEnumImageAll != NULL)
			{
				CComPtr<IVsiImage> pImage;
				while (S_OK == pEnumImageAll->Next(1, &pImage, NULL))
				{
					CComPtr<IVsiStudy> pStudy;
					hr = pImage->GetStudy(&pStudy);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vLocked(false);
					hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK != hr || VARIANT_FALSE == V_BOOL(&vLocked))
					{
						// Loaded?
						CComVariant vNs;
						hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						HRESULT hrLoaded = m_pSession->GetIsItemLoaded(V_BSTR(&vNs), NULL);
						if (S_OK == hrLoaded)
						{
							++iImageLoaded;
						}
					}

					pImage.Release();
				}
			}

			// Process images
			if (pEnumImage != NULL)
			{
				CComPtr<IVsiImage> pImage;
				while (S_OK == pEnumImage->Next(1, &pImage, NULL))
				{
					CComPtr<IVsiStudy> pStudy;
					hr = pImage->GetStudy(&pStudy);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vLocked(false);
					hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK != hr || VARIANT_FALSE == V_BOOL(&vLocked))
					{
						vecImageUnlocked.push_back(pImage.p);
					}
					else
					{
						vecItemsLocked.push_back(pImage.p);
					}

					pImage.Release();
				}
			}

			// Process series
			CComPtr<IVsiSeries> pActiveSeries;
			hr = m_pSession->GetPrimarySeries(&pActiveSeries);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			BOOL bDeleteActiveSeires(FALSE);

			if (pEnumSeries != NULL)
			{
				CComPtr<IVsiSeries> pSeries;
				while (S_OK == pEnumSeries->Next(1, &pSeries, NULL))
				{
					CComPtr<IVsiStudy> pStudy;
					hr = pSeries->GetStudy(&pStudy);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vLocked(false);
					hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK != hr || VARIANT_FALSE == V_BOOL(&vLocked))
					{
						vecSeriesUnlocked.push_back(pSeries.p);

						// Check if active series will be deleted
						if (pActiveSeries.IsEqualObject(pSeries))
						{
							bDeleteActiveSeires = TRUE;
						}

						// Loaded?
						CComVariant vNs;
						hr = pSeries->GetProperty(VSI_PROP_SERIES_NS, &vNs);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						HRESULT hrLoaded = m_pSession->GetIsItemLoaded(V_BSTR(&vNs), NULL);
						if (S_OK == hrLoaded)
						{
							CCritSecLock csl(m_csUpdateItems);
							m_listUpdateItems.push_back(CString(V_BSTR(&vNs)));
							++iSeriesLoaded;
						}
					}
					else
					{
						vecItemsLocked.push_back(pSeries.p);
					}

					pSeries.Release();
				}
			}

			// Process study
			if (pEnumStudy != NULL)
			{
				CComPtr<IVsiStudy> pStudy;
				while (S_OK == pEnumStudy->Next(1, &pStudy, NULL))
				{
					CComVariant vLocked(false);
					hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK != hr || VARIANT_FALSE == V_BOOL(&vLocked))
					{
						vecStudyUnlocked.push_back(pStudy.p);

						// Loaded?
						CComVariant vNs;
						hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vNs);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						HRESULT hrLoaded = m_pSession->GetIsItemLoaded(V_BSTR(&vNs), NULL);
						if (S_OK == hrLoaded)
						{
							CCritSecLock csl(m_csUpdateItems);
							m_listUpdateItems.push_back(CString(V_BSTR(&vNs)));
							++iStudyLoaded;
						}
					}
					else
					{
						vecItemsLocked.push_back(pStudy.p);
					}

					pStudy.Release();
				}
			}
		
			if (vecImageUnlocked.size() > 0 ||
				vecSeriesUnlocked.size() > 0 ||
				vecStudyUnlocked.size() > 0)
			{
				IUnknown **ppStudy = NULL;
				IUnknown **ppSeries = NULL;
				IUnknown **ppImage = NULL;
		
				if (vecStudyUnlocked.size() > 0)
					ppStudy = &(*vecStudyUnlocked.begin());
				if (vecSeriesUnlocked.size() > 0)
					ppSeries = &(*vecSeriesUnlocked.begin());
				if (vecImageUnlocked.size() > 0)
					ppImage = &(*vecImageUnlocked.begin());
		
				hr = m_pAppController->SetState(VSI_APP_STATE_DELETE_CONFIRMATION, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CVsiDeleteConfirmation dlg(
					ppStudy,
					(int)vecStudyUnlocked.size(),
					ppSeries,
					(int)vecSeriesUnlocked.size(),
					ppImage,
					(int)vecImageUnlocked.size());
				INT_PTR rtn = dlg.DoModal();
				if (IDYES == rtn)
				{
					hr = m_pAppController->RemoveState(VSI_APP_STATE_DELETE_CONFIRMATION, TRUE);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (0 < iStudyLoaded || 0 < iSeriesLoaded)
					{
						// Close acquisition session
						hr = m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_SESSION, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						// Refresh all
						{
							CCritSecLock csl(m_csUpdateItems);
							m_listUpdateItems.push_back(L"");
						
							PostMessage(WM_VSI_SB_UPDATE_ITEMS);
						}
					}

					// If images are opened - close them automatically now
					if (0 < iImageLoaded)
					{
						CComQIPtr<IVsiEnumImage> pEnumImageToClose(*(listSel.ppEnumImage));
						VSI_CHECK_INTERFACE(pEnumImageToClose, VSI_LOG_ERROR, NULL);
					
						hr = pEnumImageToClose->Reset();
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComVariant vNs;
						CComPtr<IVsiImage> pImage;
						while (S_OK == pEnumImageToClose->Next(1, &pImage, NULL))
						{
							vNs.Clear();
							hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							int iSlot;
							HRESULT hrLoaded = m_pSession->GetIsItemLoaded(V_BSTR(&vNs), &iSlot);
							if (S_OK == hrLoaded)
							{
								CComVariant vParam(pImage.p);
								hr = m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_IMAGE, &vParam);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							}

							pImage.Release();
						}
					
						// Updates cancel button state
						{
							int iImage(0);
							for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
							{
								CComPtr<IUnknown> pUnkView;
								hr = m_pSession->GetModeView(i, &pUnkView);
								if (S_OK == hr)
									++iImage;
							}
						
							if (0 == iImage)
							{
								m_State = VSI_SB_STATE_START;
								GetDlgItem(IDCANCEL).EnableWindow(FALSE);
							}
						}
					}
				
					if (m_pItemPreview != NULL)
					{
						StopPreview(VSI_PREVIEW_LIST_CLEAR);
					}

					// XML for UI update
					CComPtr<IXMLDOMDocument> pXmlDoc;
					hr = VsiCreateDOMDocument(&pXmlDoc);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating DOMDocument");

					// Root element
					CComPtr<IXMLDOMElement> pElmUpdate;
					hr = pXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_UPDATE), &pElmUpdate);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure creating '%s' element", VSI_UPDATE_XML_ELM_UPDATE));

					hr = pXmlDoc->appendChild(pElmUpdate, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IXMLDOMElement> pElmRemove;
					hr = pXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_REMOVE), &pElmRemove);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure creating '%s' element", VSI_UPDATE_XML_ELM_REMOVE));

					// Set number of studies
					CComVariant vNum(iStudyCount);
					hr = pElmRemove->setAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_STUDIES), vNum);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failed to set attribute (%s)", VSI_REMOVE_XML_ATT_NUMBER_STUDIES));

					// Set number of series
					vNum = iSeriesCount;
					hr = pElmRemove->setAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_SERIES), vNum);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failed to set attribute (%s)", VSI_REMOVE_XML_ATT_NUMBER_SERIES));

					// Set number of images
					vNum = iImageCount;
					hr = pElmRemove->setAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_IMAGES), vNum);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failed to set attribute (%s)", VSI_REMOVE_XML_ATT_NUMBER_IMAGES));

					hr = pElmUpdate->appendChild(pElmRemove, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
					BOOL bContinue(TRUE);
					CVsiVecDeleteItemsIter iter;

					if (bContinue)
					{
						for (iter = vecImageUnlocked.begin();
							iter != vecImageUnlocked.end();
							++iter)
						{
							CComQIPtr<IVsiImage> pImage(*iter);
							hr = Delete(pImage, pXmlDoc, pElmRemove);
							if (S_OK != hr)
							{
								VSI_LOG_MSG(VSI_LOG_WARNING, L"Errors occurred deleting image items");
								bContinue = FALSE;
							}
						}
					}
				
					if (bContinue)
					{
						for (iter = vecSeriesUnlocked.begin();
							iter != vecSeriesUnlocked.end();
							++iter)
						{
							CComQIPtr<IVsiSeries> pSeries(*iter);
							hr = Delete(pSeries, pXmlDoc, pElmRemove);
							if (S_OK != hr)
							{
								VSI_LOG_MSG(VSI_LOG_WARNING, L"Errors occurred deleting series items");
								bContinue = FALSE;
							}
						}
					}
				
					if (bContinue)
					{
						for (iter = vecStudyUnlocked.begin();
							iter != vecStudyUnlocked.end();
							++iter)
						{
							CComQIPtr<IVsiStudy> pStudy(*iter);
							hr = Delete(pStudy, pXmlDoc, pElmRemove);
							if (S_OK != hr)
							{
								VSI_LOG_MSG(VSI_LOG_WARNING, L"Errors occurred deleting study items");
								bContinue = FALSE;
							}
						}
					}

					if (!bContinue)
					{
						CString strMsg;

						DWORD dwState = VsiGetBitfieldValue<DWORD>(
							VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
							VSI_PARAMETER_SYS_STATE, m_pPdm);
						if (0 < (VSI_SYS_STATE_HARDWARE & dwState))
						{
							// Hardware
							strMsg.LoadString(IDS_STYBRW_ERROR_WHILE_DELETE_STUDY1);
						}
						else
						{
							// Workstation
							strMsg.LoadString(IDS_STYBRW_ERROR_WHILE_DELETE_STUDY2);
						}

						WCHAR szSupport[500];
						VsiGetTechSupportInfo(szSupport, _countof(szSupport));

						strMsg += szSupport;

						VsiMessageBox(
							GetTopLevelParent(),
							strMsg,
							VsiGetRangeValue<CString>(
								VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
								VSI_PARAMETER_SYS_PRODUCT_NAME,
								m_pPdm),
							MB_OK | MB_ICONEXCLAMATION);
					}
					else
					{
						if (vecItemsLocked.size() > 0)
						{
							VsiMessageBox(
								GetTopLevelParent(),
								CString(MAKEINTRESOURCE(IDS_STYBRW_LOCKED_TO_DELETE_UNLOCK)),
								VsiGetRangeValue<CString>(
									VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
									VSI_PARAMETER_SYS_PRODUCT_NAME,
									m_pPdm),
								MB_OK | MB_ICONINFORMATION);
						}
					}

					UpdateDataList(pXmlDoc);

					// Delete active series file
					if (bDeleteActiveSeires)
					{
						hr = m_pSession->CheckActiveSeries(TRUE, NULL, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}
				else
				{
					hr = m_pAppController->RemoveState(VSI_APP_STATE_DELETE_CONFIRMATION, TRUE);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}
			else  // No unlocked items
			{
				if (1 < vecItemsLocked.size())
				{
					VsiMessageBox(
						GetTopLevelParent(),
						CString(MAKEINTRESOURCE(IDS_STYBRW_ITEMS_LOCKED_CANNOT_DELETE)),
						VsiGetRangeValue<CString>(
							VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
							VSI_PARAMETER_SYS_PRODUCT_NAME,
							m_pPdm),
						MB_OK | MB_ICONINFORMATION);
				}
				else if (1 == vecItemsLocked.size())
				{
					VsiMessageBox(
						GetTopLevelParent(),
						CString(MAKEINTRESOURCE(IDS_STYBRW_ITEM_LOCKED_CANNOT_DELETE)),
						VsiGetRangeValue<CString>(
							VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
							VSI_PARAMETER_SYS_PRODUCT_NAME,
							m_pPdm),
						MB_OK | MB_ICONINFORMATION);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedExport(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Export");
	HRESULT hr(S_OK);

	// Check button state - this command can come from the hot key
	if (!GetDlgItem(IDC_SB_EXPORT).IsWindowEnabled())
		return 0;

	CancelSearchIfRunning();

	try
	{
		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			CComPtr<IVsiPropertyBag> pProperties;
			hr = pProperties.CoCreateInstance(__uuidof(VsiPropertyBag));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IVsiEnumImage> pEnumImage;
			int iImages = 0;
			VSI_DATA_LIST_COLLECTION selection;

			memset(&selection, 0, sizeof(VSI_DATA_LIST_COLLECTION));
			selection.dwFlags = VSI_DATA_LIST_FILTER_MOVE_TO_CHILDREN;
			selection.ppEnumImage = &pEnumImage;
			selection.piImageCount = &iImages;

			hr = m_pDataList->GetSelection(&selection);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (pEnumImage != NULL)
			{
				CComVariant vImages(pEnumImage);
				hr = pProperties->Write(L"images", &vImages);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			CComVariant vTable((ULONGLONG)(INT_PTR)GetDlgItem(IDC_SB_STUDY_LIST).m_hWnd);
			hr = pProperties->Write(L"table", &vTable);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			lock.UnlockData();

			CComVariant v(pProperties.p);
			m_pCmdMgr->InvokeCommand(ID_CMD_EXPORT, &v);
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedCopyTo(WORD, WORD, HWND, BOOL&)
{
	// Check button state - this command can come from the hot key
	if (!GetDlgItem(IDC_SB_COPY_TO).IsWindowEnabled())
		return 0;

	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Copy To");

	CancelSearchIfRunning();

	HRESULT hr = S_OK;
	CComPtr<IXMLDOMDocument> pXmlDoc;

	try
	{
		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			// Create XML DOM document with job list
			INT64 iTotalSize = 0;

			// Make sure the selection record is up-to-date
			int iStudy = 0, iSeries = 0, iImage = 0;
			CComPtr<IVsiEnumStudy> pEnumStudy;
			CComPtr<IVsiEnumSeries> pEnumSeries;
			CComPtr<IVsiEnumImage> pEnumImage;
			VSI_DATA_LIST_COLLECTION listSel;

			memset(&listSel, 0, sizeof(listSel));
			listSel.dwFlags = VSI_DATA_LIST_FILTER_SELECT_PARENT;
			listSel.ppEnumStudy = &pEnumStudy;
			listSel.piStudyCount = &iStudy;
			listSel.ppEnumSeries = &pEnumSeries;
			listSel.piSeriesCount = &iSeries;
			listSel.ppEnumImage = &pEnumImage;
			listSel.piImageCount = &iImage;
			hr = m_pDataList->GetSelection(&listSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			lock.UnlockData();
		
			if (0 != (iStudy + iSeries + iImage))
			{
				// Make sure all studies has owner assigned
				{
					// Loops selected studies
					if (pEnumStudy != NULL)
					{
						CComPtr<IVsiStudy> pStudy;
						while (pEnumStudy->Next(1, &pStudy, NULL) == S_OK)
						{
							CComVariant vOwner;
							hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if (S_OK != hr || (0 == *V_BSTR(&vOwner)))
							{
								VsiMessageBox(
									GetTopLevelParent(),
									CString(MAKEINTRESOURCE(IDS_STYBRW_COPY_TO_NO_OWNER)),
									CString(MAKEINTRESOURCE(IDS_STYBRW_COPY_TO_TITLE)),
									MB_OK | MB_ICONINFORMATION);
								return 0;
							}

							pStudy.Release();
						}

						pEnumStudy->Reset();
					}

					// Loops selected series
					if (pEnumSeries != NULL)
					{
						CComPtr<IVsiSeries> pSeries;
						while (pEnumSeries->Next(1, &pSeries, NULL) == S_OK)
						{
							CComPtr<IVsiStudy> pStudy;
							hr = pSeries->GetStudy(&pStudy);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							CComVariant vOwner;
							hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if (S_OK != hr || (0 == *V_BSTR(&vOwner)))
							{
								VsiMessageBox(
									GetTopLevelParent(),
									CString(MAKEINTRESOURCE(IDS_STYBRW_COPY_TO_NO_OWNER)),
									CString(MAKEINTRESOURCE(IDS_STYBRW_COPY_TO_TITLE)),
									MB_OK | MB_ICONINFORMATION);
								return 0;
							}

							pSeries.Release();
						}

						pEnumSeries->Reset();
					}

					// Loops selected images
					if (pEnumImage != NULL)
					{
						CComPtr<IVsiImage> pImage;
						while (pEnumImage->Next(1, &pImage, NULL) == S_OK)
						{
							CComPtr<IVsiStudy> pStudy;
							hr = pImage->GetStudy(&pStudy);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							CComVariant vOwner;
							hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if (S_OK != hr || (0 == *V_BSTR(&vOwner)))
							{
								VsiMessageBox(
									GetTopLevelParent(),
									CString(MAKEINTRESOURCE(IDS_STYBRW_COPY_TO_NO_OWNER)),
									CString(MAKEINTRESOURCE(IDS_STYBRW_COPY_TO_TITLE)),
									MB_OK | MB_ICONINFORMATION);
								return 0;
							}

							pImage.Release();
						}

						pEnumImage->Reset();
					}
				}

				if (iImage > 0)
				{
					int iRet = VsiMessageBox(
						GetTopLevelParent(),
						CString(MAKEINTRESOURCE(IDS_STYBRW_ONLYSUBSET_SELECTED)),
						CString(MAKEINTRESOURCE(IDS_STYBRW_COPY_TO_TITLE)),
						MB_YESNO | MB_ICONWARNING);
					if (IDYES != iRet)
					{
						return 0;
					}
				}

				// Sort everything
				typedef std::map<CString, CComPtr<IXMLDOMElement>> CVsiMapNsToElm;
				typedef CVsiMapNsToElm::iterator CVsiMapNsToElmIter;
		
				CVsiMapNsToElm m_mapNsToElm;

				CComPtr<IXMLDOMElement> pElmRoot;
				WCHAR szPrevStudyId[128];
				WCHAR szPrevSeriesId[128];

				szPrevStudyId[0] = NULL;
				szPrevSeriesId[0] = NULL;

				// Creates XML doc
				hr = VsiCreateDOMDocument(&pXmlDoc);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating DOMDocument");

				// Root element
				hr = pXmlDoc->createElement(CComBSTR(VSI_STUDY_XML_ELM_STUDIES), &pElmRoot);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure creating '%s' element", VSI_STUDY_XML_ELM_STUDIES));

				hr = pXmlDoc->appendChild(pElmRoot, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vExport(VSI_JOB_STUDY_EXPORT);
				hr = pElmRoot->setAttribute(VSI_STUDIES_XML_ATT_JOB, vExport);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failed to set attribute (%s)", VSI_STUDIES_XML_ATT_JOB));

				CComVariant var;

				// Loops selected studies
				CComPtr<IVsiStudy> pStudy;
				if (pEnumStudy != NULL)
				{
					while (pEnumStudy->Next(1, &pStudy, NULL) == S_OK)
					{
						VSI_STUDY_ERROR dwError;
						hr = pStudy->GetErrorCode(&dwError);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (VSI_STUDY_ERROR_LOAD_XML == dwError)
						{
							WCHAR szSupport[500];
							VsiGetTechSupportInfo(szSupport, _countof(szSupport));

							CString strMsg(MAKEINTRESOURCE(IDS_STUDY_OPEN_ERROR));
							strMsg += szSupport;

							VsiMessageBox(
								GetTopLevelParent(),
								strMsg,
								VsiGetRangeValue<CString>(
									VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
									VSI_PARAMETER_SYS_PRODUCT_NAME,
									m_pPdm),
								MB_OK | MB_ICONEXCLAMATION);

							return 0;
						}

						CComPtr<IXMLDOMElement> pStudyElm;

						hr = VsiCreateStudyElement(pXmlDoc, pElmRoot, pStudy, &pStudyElm, NULL, FALSE, TRUE);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,
							VsiFormatMsg(L"Failure creating '%s' element", VSI_STUDY_XML_ELM_STUDY));

						var.Clear();
						hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &var);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, L"Failure getting study NS");
	
						m_mapNsToElm.insert(CVsiMapNsToElm::value_type(V_BSTR(&var), pStudyElm));

						// Add all series
						CComPtr<IVsiEnumSeries> pEnum;
						hr = pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_CREATED_DATE, FALSE, &pEnum, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						// Make sure we have series under the study (bug 15944)
						if (S_OK == hr)
						{
							CComPtr<IVsiSeries> pSeries;
							while (pEnum->Next(1, &pSeries, NULL) == S_OK)
							{
								hr = VsiCreateSeriesElement(pXmlDoc, pStudyElm, pSeries, NULL, NULL);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR,
									VsiFormatMsg(L"Failure creating '%s' element", VSI_SERIES_XML_ELM_SERIES));

								pSeries.Release();
							}
						}

						pStudy.Release();
					}
				}

				// Loops selected series
				CComPtr<IVsiSeries> pSeries;
				CComPtr<IXMLDOMElement> pStudyElement;
				if (pEnumSeries != NULL)
				{
					while (pEnumSeries->Next(1, &pSeries, NULL) == S_OK)
					{
						// Creates study element if necessary
						hr = pSeries->GetStudy(&pStudy);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						VSI_STUDY_ERROR dwError;
						hr = pStudy->GetErrorCode(&dwError);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (VSI_STUDY_ERROR_LOAD_XML == dwError)
						{
							WCHAR szSupport[500];
							VsiGetTechSupportInfo(szSupport, _countof(szSupport));

							CString strMsg(MAKEINTRESOURCE(IDS_STUDY_OPEN_ERROR));
							strMsg += szSupport;

							VsiMessageBox(
								GetTopLevelParent(),
								strMsg,
								VsiGetRangeValue<CString>(
									VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
									VSI_PARAMETER_SYS_PRODUCT_NAME,
									m_pPdm),
								MB_OK | MB_ICONEXCLAMATION);

							return 0;
						}

						var.Clear();
						hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &var);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, L"Failure getting study ID");
				
						CVsiMapNsToElmIter iter = m_mapNsToElm.find(CString(V_BSTR(&var)));
						if (iter != m_mapNsToElm.end())
						{
							pStudyElement = iter->second;
						}
						else
						{
							// New study found

							pStudyElement.Release();
							hr = VsiCreateStudyElement(pXmlDoc, pElmRoot, pStudy, &pStudyElement, NULL, FALSE, TRUE);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,
								VsiFormatMsg(L"Failure creating '%s' element", VSI_STUDY_XML_ELM_STUDY));

							var.Clear();
							hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &var);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, L"Failure getting study NS");

							m_mapNsToElm.insert(CVsiMapNsToElm::value_type(V_BSTR(&var), pStudyElement));

							// 0 size at the beginning
							CComVariant vSize(0);
							hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), vSize);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}

						_ASSERT(pStudyElement != NULL);

						CComPtr<IXMLDOMElement> pSeriesElement;
						hr = VsiCreateSeriesElement(pXmlDoc, pStudyElement, pSeries, &pSeriesElement, &iTotalSize);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,
							VsiFormatMsg(L"Failure creating '%s' element", VSI_SERIES_XML_ELM_SERIES));

						var.Clear();
						hr = pSeries->GetProperty(VSI_PROP_SERIES_NS, &var);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, L"Failure getting series NS");

						m_mapNsToElm.insert(CVsiMapNsToElm::value_type(V_BSTR(&var), pSeriesElement));

						// Update size
						CComVariant vSize(0);
						hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), &vSize);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						vSize.ChangeType(VT_I8);
						V_I8(&vSize) += iTotalSize;

						hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), vSize);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						pStudy.Release();
						pSeries.Release();
					}
				}

				// Loops selected images
				CComPtr<IVsiImage> pImage;
				if (pEnumImage != NULL)
				{
					while (pEnumImage->Next(1, &pImage, NULL) == S_OK)
					{
						// Creates study element if necessary
						hr = pImage->GetSeries(&pSeries);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						VSI_SERIES_ERROR dwSeriesError;
						hr = pSeries->GetErrorCode(&dwSeriesError);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (VSI_SERIES_ERROR_LOAD_XML == dwSeriesError)
						{
							WCHAR szSupport[500];
							VsiGetTechSupportInfo(szSupport, _countof(szSupport));

							CString strMsg(MAKEINTRESOURCE(IDS_STYBRW_ERROR_WHILE_OPEN_SERIES));
							strMsg += szSupport;

							VsiMessageBox(
								GetTopLevelParent(),
								strMsg,
								VsiGetRangeValue<CString>(
									VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
									VSI_PARAMETER_SYS_PRODUCT_NAME,
									m_pPdm),
								MB_OK | MB_ICONEXCLAMATION);

							return 0;
						}

						hr = pSeries->GetStudy(&pStudy);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						VSI_STUDY_ERROR dwStudyError;
						hr = pStudy->GetErrorCode(&dwStudyError);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (VSI_STUDY_ERROR_LOAD_XML == dwStudyError)
						{
							WCHAR szSupport[500];
							VsiGetTechSupportInfo(szSupport, _countof(szSupport));

							CString strMsg(MAKEINTRESOURCE(IDS_STUDY_OPEN_ERROR));
							strMsg += szSupport;

							VsiMessageBox(
								GetTopLevelParent(),
								strMsg,
								VsiGetRangeValue<CString>(
									VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
									VSI_PARAMETER_SYS_PRODUCT_NAME,
									m_pPdm),
								MB_OK | MB_ICONEXCLAMATION);

							return 0;
						}

						CComPtr<IXMLDOMElement> pSeriesElement;

						var.Clear();
						hr = pSeries->GetProperty(VSI_PROP_SERIES_NS, &var);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, L"Failure getting series ID");
				
						CVsiMapNsToElmIter iter = m_mapNsToElm.find(CString(V_BSTR(&var)));
						if (iter != m_mapNsToElm.end())
						{
							pSeriesElement = iter->second;
						}
						else
						{
							var.Clear();
							hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &var);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, L"Failure getting study ID");

							iter = m_mapNsToElm.find(CString(V_BSTR(&var)));
					
							if (iter != m_mapNsToElm.end())
							{
								pStudyElement = iter->second;
							}
							else
							{
								pStudyElement.Release();
								hr = VsiCreateStudyElement(pXmlDoc, pElmRoot, pStudy, &pStudyElement, NULL, FALSE, TRUE);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR,
									VsiFormatMsg(L"Failure creating '%s' element", VSI_STUDY_XML_ELM_STUDY));

								var.Clear();
								hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &var);
								VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, L"Failure getting study NS");

								m_mapNsToElm.insert(CVsiMapNsToElm::value_type(V_BSTR(&var), pStudyElement));

								// 0 size at the beginning
								CComVariant vSize(0);
								hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), vSize);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR,
									VsiFormatMsg(L"Failed to set attribute (%s)", VSI_STUDY_XML_ATT_SIZE));
							}
							_ASSERT(pStudyElement != NULL);

							hr = VsiCreateSeriesElement(pXmlDoc, pStudyElement, pSeries, &pSeriesElement, NULL);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,
								VsiFormatMsg(L"Failure creating '%s' element", VSI_SERIES_XML_ELM_SERIES));

							var.Clear();
							hr = pSeries->GetProperty(VSI_PROP_SERIES_NS, &var);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, L"Failure getting study NS");

							m_mapNsToElm.insert(CVsiMapNsToElm::value_type(V_BSTR(&var), pSeriesElement));

							// Add Series.xml and Measurement.xml size
							{
								CComHeapPtr<OLECHAR> pszSeriesPath;
								hr = pSeries->GetDataPath(&pszSeriesPath);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						
								LARGE_INTEGER liSizeSeriesXml;
								hr = VsiGetFileSize(pszSeriesPath, &liSizeSeriesXml);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						
								WCHAR szMsmntPath[MAX_PATH];
								wcscpy_s(szMsmntPath, _countof(szMsmntPath), pszSeriesPath);
								PathRemoveFileSpec(szMsmntPath);
								wcscat_s(szMsmntPath, _countof(szMsmntPath), L"\\" VSI_FILE_MEASUREMENT);

								LARGE_INTEGER liSizeMsmntXml = { 0 };
								if (0 == _waccess(szMsmntPath, 0))
								{
									hr = VsiGetFileSize(szMsmntPath, &liSizeMsmntXml);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								}
						
								// Update size
								CComVariant vSize(0);
								hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), &vSize);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR,
									VsiFormatMsg(L"Failed to get attribute (%s)", VSI_STUDY_XML_ATT_SIZE));

								vSize.ChangeType(VT_I8);
								V_I8(&vSize) += liSizeSeriesXml.QuadPart;
								V_I8(&vSize) += liSizeMsmntXml.QuadPart;

								hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), vSize);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							}
						}

						_ASSERT(pSeriesElement != NULL);

						hr = VsiCreateImageElement(pXmlDoc, pSeriesElement, pImage, NULL, &iTotalSize);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,
							VsiFormatMsg(L"Failure creating '%s' element", VSI_SERIES_XML_ELM_SERIES));

						// Update size
						CComVariant vSize(0);
						hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), &vSize);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,
							VsiFormatMsg(L"Failed to get attribute (%s)", VSI_STUDY_XML_ATT_SIZE));

						vSize.ChangeType(VT_I8);
						V_I8(&vSize) += iTotalSize;

						hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), vSize);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						pStudy.Release();
						pSeries.Release();
						pImage.Release();
					}
				}

#ifdef _DEBUG
{
				WCHAR szPath[MAX_PATH];
				GetModuleFileName(NULL, szPath, _countof(szPath));
				LPWSTR pszSpt = wcsrchr(szPath, L'\\');
				lstrcpy(pszSpt + 1, VSI_FOLDER_DEBUG);
				lstrcat(szPath, L"\\copyto.xml");
				pXmlDoc->save(CComVariant(szPath));
}
#endif

				// Passes IXMLDOMDocument
				CComVariant v(pXmlDoc);
				m_pCmdMgr->InvokeCommand(ID_CMD_OPEN_COPY_STUDY_TO, &v);
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedCopyFrom(WORD, WORD, HWND, BOOL&)
{
	// Check button state - this command can come from the hot key
	if (!GetDlgItem(IDC_SB_COPY_FROM).IsWindowEnabled())
		return 0;

	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Copy From");

	CancelSearchIfRunning();

	m_pCmdMgr->InvokeCommand(ID_CMD_OPEN_COPY_STUDY_FROM, NULL);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedNew(WORD, WORD, HWND, BOOL&)
{
	// Check button state - this command can come from the hot key
	if (!GetDlgItem(IDC_SB_NEW).IsWindowEnabled())
		return 0;

	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"New");

	CancelSearchIfRunning();

	m_pCmdMgr->InvokeCommand(ID_CMD_NEW_STUDY_SELECT, NULL);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedAnalysis(WORD, WORD, HWND, BOOL&)
{
	// Check button state - this command can come from the hot key
	if (!GetDlgItem(IDC_SB_ANALYSIS).IsWindowEnabled())
		return 0;

	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Analysis");
	HRESULT hr = S_OK;

	CancelSearchIfRunning();

	try
	{
		CWaitCursor wait;

		hr = GenerateFullDisclosureMsmntReport();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedStrain(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Vevo Strain");
	HRESULT hr = S_OK;

	try
	{
		if (!VsiLicenseUtils::CheckFeatureAvailable(VSI_LICENSE_FEATURE_VEVO_STRAIN))
		{
			m_pCmdMgr->InvokeCommand(ID_CMD_FEATURE_NOT_AVAILABLE, NULL);
			return 0;
		}

		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			// Gets focused item
			int iFocused(-1);
			hr = m_pDataList->GetSelectedItem(&iFocused);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (iFocused >= 0)
			{
				VSI_DATA_LIST_TYPE type(VSI_DATA_LIST_NONE);
				CComPtr<IUnknown> pUnkItem;

				hr = m_pDataList->GetItem(iFocused, &type, &pUnkItem);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
				if (VSI_DATA_LIST_IMAGE == type)
				{
					CComVariant vImage(pUnkItem);
					m_pCmdMgr->InvokeCommand(ID_CMD_VEVO_STRAIN, &vImage);

					CComQIPtr<IVsiImage> pImage(pUnkItem);
					if (NULL != pImage)
					{
						// Refresh strain status
						CComQIPtr<IVsiPersistImage> pPersistImage(pImage);
						VSI_CHECK_INTERFACE(pPersistImage, VSI_LOG_ERROR, NULL);

						hr = pPersistImage->LoadImageData(NULL, NULL, NULL, 0);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						// Update UI
						CComVariant vNs;
						hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						{
							CCritSecLock csl(m_csUpdateItems);
							m_listUpdateItems.push_back(CString(V_BSTR(&vNs)));
						}

						PostMessage(WM_VSI_SB_UPDATE_ITEMS);
					}

					MSG msg;
					while (PeekMessage(&msg, NULL, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE))
					{
						;
					}
				}
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedVevoCQ(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Vevo CQ");
	HRESULT hr = S_OK;

	try
	{
		if (!VsiLicenseUtils::CheckFeatureAvailable(VSI_LICENSE_FEATURE_VEVO_CQ))
		{
			m_pCmdMgr->InvokeCommand(ID_CMD_FEATURE_NOT_AVAILABLE, NULL);
			return 0;
		}

		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			// Count how many valid sono vevo images we have
			int iNumValidImages(0);
			hr = CountSonoVevoImages(&iNumValidImages);		
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (iNumValidImages > 5)
			{
				VsiMessageBox(
					GetTopLevelParent(),
					L"Please select up to 5 images for analysis in VevoCQ.",
					L"VevoCQ",
					MB_OK | MB_ICONEXCLAMATION);

				return 0;
			}

			if(iNumValidImages > 0)
			{
				CComPtr<IVsiPropertyBag> pProperties;
				hr = pProperties.CoCreateInstance(__uuidof(VsiPropertyBag));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Get number of images selected
				CComPtr<IVsiEnumImage> pEnumImage;
				int iImages = 0;
				VSI_DATA_LIST_COLLECTION selection;

				memset(&selection, 0, sizeof(VSI_DATA_LIST_COLLECTION));
				selection.dwFlags = VSI_DATA_LIST_FILTER_MOVE_TO_CHILDREN;
				selection.ppEnumImage = &pEnumImage;
				selection.piImageCount = &iImages;

				hr = m_pDataList->GetSelection(&selection);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
				if (pEnumImage != NULL)
				{
					CComVariant vImages(pEnumImage);
					hr = pProperties->Write(L"images", &vImages);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}			

				if (iImages > 0)
				{			
					CComVariant v(pProperties.p);
					m_pCmdMgr->InvokeCommand(ID_CMD_VEVOCQ, &v);
				
					pEnumImage->Reset();

					CComPtr<IVsiImage> pImage;
					while (S_OK == pEnumImage->Next(1, &pImage, NULL))
					{
						// Check for null
						if (pImage == NULL)
							continue;

						// Refresh strain status
						CComQIPtr<IVsiPersistImage> pPersistImage(pImage);
						VSI_CHECK_INTERFACE(pPersistImage, VSI_LOG_ERROR, NULL);

						hr = pPersistImage->LoadImageData(NULL, NULL, NULL, 0);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						// Update UI
						CComVariant vNs;
						hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						{
							CCritSecLock csl(m_csUpdateItems);
							m_listUpdateItems.push_back(CString(V_BSTR(&vNs)));
						}					
					
						pImage.Release();
					}

					PostMessage(WM_VSI_SB_UPDATE_ITEMS);

					MSG msg;
					while (PeekMessage(&msg, NULL, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE))
					{
						;
					}
				}	
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedLogout(WORD, WORD, HWND, BOOL&)
{
	CancelSearchIfRunning();

	CComVariant v(VSI_CLOSE_SERIES_FLAG_LOG_OUT);
	m_pCmdMgr->InvokeCommand(ID_CMD_LOG_OUT, &v);

	return 0;
}

LRESULT CVsiStudyBrowser::OnBnClickedShutdown(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Shut Down");

	int iRet = VsiMessageBox(
		GetTopLevelParent(),
		CString(MAKEINTRESOURCE(IDS_STYBRW_SHUTDOWN_SYSTEM)),
		VsiGetRangeValue<CString>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_PRODUCT_NAME,
			m_pPdm),
		MB_YESNO | MB_ICONQUESTION);

	if (IDYES == iRet)
	{
		CancelSearchIfRunning();
	
		CComVariant vShutdown(true);
		m_pCmdMgr->InvokeCommandAsync(ID_CMD_EXIT_BEGIN, &vShutdown, FALSE);
	}

	return 0;
}


LRESULT CVsiStudyBrowser::OnBnClickedStudyAccessFilter(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		CancelSearchIfRunning();

		VSI_UI_STUDY_ACCESS_FILTER filterType =
			(VSI_UI_STUDY_ACCESS_FILTER)VsiGetDiscreteValue<DWORD>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_STUDY_ACCESS_FILTER, m_pPdm);

		VSI_UI_STUDY_ACCESS_FILTER filterTypeNew = filterType;

		switch (filterType)
		{
		case VSI_UI_STUDY_ACCESS_FILTER_PRIVATE:
			{
				filterTypeNew = VSI_UI_STUDY_ACCESS_FILTER_GROUP;
			}
			break;
		case VSI_UI_STUDY_ACCESS_FILTER_GROUP:
			{
				filterTypeNew = VSI_UI_STUDY_ACCESS_FILTER_ALL;
			}
			break;
		case VSI_UI_STUDY_ACCESS_FILTER_ALL:
			{
				filterTypeNew = VSI_UI_STUDY_ACCESS_FILTER_PRIVATE;
			}
			break;
		}

		if (filterTypeNew != filterType)
		{
			VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Study Access Filter");

			VsiSetDiscreteValue<DWORD>(
				static_cast<DWORD>(filterTypeNew),
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_STUDY_ACCESS_FILTER, m_pPdm);
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnEnChangedSearch(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		bool bSearch(false);

		CString strSearch;
		m_wndSearch.GetWindowText(strSearch);
		strSearch.Trim();

		{
			CCritSecLock cs(m_csSearch);

			if (strSearch != m_strSearch)
			{
				m_strSearch = strSearch;
				bSearch = true;

				m_wndSearch.SetSearchAnimation(true);
			}
		}

		if (bSearch)
		{
			StopPreview(VSI_PREVIEW_LIST_CLEAR);
			SetStudyInfo(NULL, false);
			DisableSelectionUI();

			// Stop current search
			m_bStopSearch = TRUE;

			// Start a new search - must after m_bStopSearch = TRUE
			SetEvent(m_hSearch);

			// Change empty message
			hr = m_pDataList->SetEmptyMessage(CString(MAKEINTRESOURCE(IDS_STYBRW_SEARCHING)));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnListItemChanged(int, LPNMHDR, BOOL& bHandled)
{
	if (IsWindowVisible())
	{
		// Delay the update as long as possible
		MSG msg;
		PeekMessage(&msg, *this, WM_VSI_SB_UPDATE_UI, WM_VSI_SB_UPDATE_UI, TRUE);
		PostMessage(WM_VSI_SB_UPDATE_UI);
	}

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiStudyBrowser::OnListItemActivate(int, LPNMHDR pnmh, BOOL& bHandled)
{
	bHandled = FALSE;

	NMITEMACTIVATE *pnmia = (NMITEMACTIVATE*)pnmh;
	if (pnmia->iItem >= 0)
	{
		HRESULT hr = S_OK;
		CComPtr<IUnknown> pUnk;
		VSI_DATA_LIST_TYPE type(VSI_DATA_LIST_NONE);

		try
		{
			CVsiLockData lock(m_pDataList);
			if (lock.LockData())
			{
				hr = m_pDataList->GetItem(pnmia->iItem, &type, &pUnk);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				switch (type)
				{
				case VSI_DATA_LIST_STUDY:
					{
						VSI_STUDY_ERROR dwError;
						CComQIPtr<IVsiStudy> pStudy(pUnk);
						if (pStudy != NULL)
						{
							pStudy->GetErrorCode(&dwError);

							if (VSI_STUDY_ERROR_VERSION_INCOMPATIBLE & dwError ||
								VSI_STUDY_ERROR_LOAD_XML & dwError)
							{
								// Open will shows incompatible message
								Open(pStudy);
								bHandled = TRUE;
							}
							else
							{
								// Continue - back to the control to expand or collapse
							}
						}
					}
					break;

				case VSI_DATA_LIST_SERIES:
					{
						VSI_SERIES_ERROR dwError;
						CComQIPtr<IVsiSeries> pSeries(pUnk);
						if (pSeries != NULL)
						{
							pSeries->GetErrorCode(&dwError);

							if (VSI_SERIES_ERROR_VERSION_INCOMPATIBLE & dwError ||
								VSI_SERIES_ERROR_LOAD_XML & dwError)
							{
								// Open will shows incompatible message
								Open(pSeries);
								bHandled = TRUE;
							}
							else
							{
								// Continue - back to the control to expand or collapse
							}
						}
					}
					break;

				case VSI_DATA_LIST_IMAGE:
					{
						CComQIPtr<IVsiImage> pImage(pUnk);
						if (pImage != NULL)
						{
							Open(pImage);
							bHandled = TRUE;
						}
					}
					break;
				}
			}
		}
		VSI_CATCH(hr);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnListItemClick(int, LPNMHDR pnmh, BOOL& bHandled)
{
	bHandled = FALSE;

	LPNMITEMACTIVATE pnmia = (LPNMITEMACTIVATE)pnmh;

	if (pnmia->iItem < 0)  // Clicks nowhere
	{
		return 0;
	}

	HRESULT hr = S_OK;

	try
	{
		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			bool bSecureMode = VsiGetBooleanValue<bool>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
				VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);

			// Checks whether clicking on the lock button
			CComPtr<IUnknown> pUnk;
			VSI_DATA_LIST_TYPE type = VSI_DATA_LIST_NONE;
			hr = m_pDataList->GetItem(pnmia->iItem, &type, &pUnk);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (VSI_DATA_LIST_COL_LOCKED != g_plvcStudyBrowserListView[pnmia->iSubItem].iSubItem ||
				VSI_DATA_LIST_STUDY != type)
			{
				return 0;
			}

			// Select the item
			ListView_SetItemState(
				pnmh->hwndFrom, pnmia->iItem,
				LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);

			BOOL bLock(FALSE);

			//
			// Collect all good studies
			std::vector<IVsiStudy*> vecStudies;
		
			CComQIPtr<IVsiStudy> pStudy(pUnk);
			VSI_CHECK_INTERFACE(pStudy, VSI_LOG_ERROR, NULL);

			VSI_STUDY_ERROR dwError(VSI_STUDY_ERROR_NONE);
			pStudy->GetErrorCode(&dwError);

			if (VSI_STUDY_ERROR_LOAD_XML == dwError || VSI_STUDY_ERROR_VERSION_INCOMPATIBLE == dwError)
			{
				// Even the one clicked on is not good
				return 0;  // Skip all
			}
		
			vecStudies.push_back(pStudy.p);

			CComVariant vLock;
			hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLock);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			bLock = (S_OK == hr && VARIANT_FALSE != V_BOOL(&vLock));
			if (S_OK != hr)
			{
				vLock = bLock ? true : false;  // Not defined yet
			}
		
			CComPtr<IVsiEnumStudy> pEnumStudy;  // vecStudies is using weak ref
			if (0 < (pnmia->uKeyFlags & (LVKF_CONTROL | LVKF_SHIFT)))
			{
				// Multiple lock and unlock
				int iStudies = 0;
				VSI_DATA_LIST_COLLECTION selection;

				memset((void*)&selection, 0, sizeof(selection));
				selection.dwFlags = VSI_DATA_LIST_FILTER_NONE;
				selection.ppEnumStudy = &pEnumStudy;
				selection.piStudyCount = &iStudies;

				hr = m_pDataList->GetSelection(&selection);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (NULL != pEnumStudy)
				{
					CComVariant vIdFocused;
					hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vIdFocused);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiStudy> pStudyTest;
					while (S_OK == pEnumStudy->Next(1, &pStudyTest, NULL))
					{
						CComVariant vIdTest;
						hr = pStudyTest->GetProperty(VSI_PROP_STUDY_ID, &vIdTest);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						// Ignores the focused item - already handled by this function
						if (vIdTest != vIdFocused)
						{
							VSI_STUDY_ERROR dwErrorTest(VSI_STUDY_ERROR_NONE);
							pStudyTest->GetErrorCode(&dwErrorTest);

							// Ignores bad study
							if (VSI_STUDY_ERROR_NONE == dwErrorTest)
							{
								CComVariant vLockTest;
								hr = pStudyTest->GetProperty(VSI_PROP_STUDY_LOCKED, &vLockTest);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							
								if (S_FALSE == hr)
								{
									vLockTest = false;
								}

								if (vLock == vLockTest)
								{
									vecStudies.push_back(pStudyTest.p);
								}
							}
						}
					
						pStudyTest.Release();
					}
				}
			}
		
			// Permission check
			BOOL bAdminAuthenticated(FALSE);

			hr = m_pOperatorMgr->GetIsAdminAuthenticated(&bAdminAuthenticated);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
			if (!bAdminAuthenticated)
			{
				std::set<CString> setAdmin;
				std::set<CString> setStandard;
			
				std::vector<IVsiStudy*>::iterator iter;
			
				for (iter = vecStudies.begin(); iter != vecStudies.end(); ++iter)
				{
					IVsiStudy *pStudyTest = *iter;
				
					CComVariant vOwner;
					hr = pStudyTest->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiOperator> pOperatorOwner;


					hr = m_pOperatorMgr->GetOperator(
						VSI_OPERATOR_PROP_NAME,
						V_BSTR(&vOwner),
						&pOperatorOwner);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						BOOL bAuthenticated(FALSE);

						if (S_OK == pOperatorOwner->HasPassword())
						{
							CComHeapPtr<OLECHAR> pszIdOwner;

							hr = pOperatorOwner->GetId(&pszIdOwner);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							hr = m_pOperatorMgr->GetIsAuthenticated(
								pszIdOwner,
								&bAuthenticated);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
						else
						{
							// In this case, no password is authenticated
							bAuthenticated = TRUE;
						}

						if (!bAuthenticated)
						{
							if (!bSecureMode)
							{
								VSI_OPERATOR_TYPE dwType;
								hr = pOperatorOwner->GetType(&dwType);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								// Admin can do anything
								if (VSI_OPERATOR_TYPE_ADMIN == dwType)
								{
									setAdmin.insert(V_BSTR(&vOwner));
								}
								else
								{
									setStandard.insert(V_BSTR(&vOwner));
								}
							}
							else
							{
								CString strMsg;
								if (1 < vecStudies.size())
								{
									// for more than one
									strMsg = VsiFormatMsg(L"The study owner(s) or an administrator must log in to %s the studies.", bLock ? L"unlock" : L"lock");
								}
								else
								{
									VSI_OPERATOR_TYPE dwType;
									hr = pOperatorOwner->GetType(&dwType);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
									if (VSI_OPERATOR_TYPE_ADMIN == dwType)
									{
										strMsg = VsiFormatMsg(L"%s, must log in to %s the study.", V_BSTR(&vOwner), bLock ? L"unlock" : L"lock");
									}
									else
									{
										strMsg = VsiFormatMsg(L"%s, or an administrator, must log in to %s the study.", V_BSTR(&vOwner), bLock ? L"unlock" : L"lock");
									}
								}

								VsiMessageBox(
									GetTopLevelParent(),
									strMsg,
									VsiGetRangeValue<CString>(
										VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
										VSI_PARAMETER_SYS_PRODUCT_NAME,
										m_pPdm),
									MB_OK | MB_ICONINFORMATION);

								return 0;
							}
						}
					}
					else
					{
						// Owner gone
						setAdmin.insert(V_BSTR(&vOwner));
					}
				}

				// Check whether there are any admin with password
				DWORD dwAdminWithPassowrd = 0;
				hr = m_pOperatorMgr->GetOperatorCount(
					VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_PASSWORD,
					&dwAdminWithPassowrd);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
				std::set<CString>::const_iterator iterOper;
				for (iterOper = setAdmin.begin();
					iterOper != setAdmin.end();
					++iterOper)
				{
					// Only using the 1st in the set
					const CString &strOper = *iterOper;

					CComPtr<IVsiOperator> pOperator;

					hr = m_pOperatorMgr->GetOperator(
						VSI_OPERATOR_PROP_NAME,
						strOper,
						&pOperator);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
					if (NULL == pOperator)
					{
						// Owner gone
						if (bSecureMode)
						{
							// Check whether there are any admin with password
							if (0 == dwAdminWithPassowrd)
							{
								// Owner gone, no admin with password, just let the user do it
								continue;
							}
						}
						else
						{
							// Owner gone, standard mode, just let the user do it (bug# 16922)
							continue;
						}
					}

					CString strTitle;
					CString strMsg;

					if (bLock)
					{
						// Unlock study
						strTitle.LoadString(IDS_STYBRW_UNLOCK_STUDY);
					}
					else
					{
						// Lock study
						strTitle.LoadString(IDS_STYBRW_LOCK_STUDY);
					}

					if (1 < setStandard.size() + setAdmin.size())
					{
						strMsg.Format(L"Enter study owner or administrator password to %s the study. Use administrator password to %s multiple studies at once.", bLock ? L"unlock" : L"lock",  bLock ? L"unlock" : L"lock");
					}
					else
					{
						strMsg.Format(L"Enter study owner or administrator password to %s the study.", bLock ? L"unlock" : L"lock");
					}

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
						&(pOperator.p));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						bAdminAuthenticated = TRUE;
						break;
					}
					else
					{
						// Cancel
						return 0;
					}
				}
			
				if (!bAdminAuthenticated)
				{
					CString strTitle;
					CString strMsg;

					if (bLock)
					{
						// Unlock study
						strTitle.LoadString(IDS_STYBRW_UNLOCK_STUDY);
					}
					else
					{
						// Lock study
						strTitle.LoadStringW(IDS_STYBRW_LOCK_STUDY);
					}
					if (1 < setStandard.size() + setAdmin.size())
					{
						strMsg.Format(L"Enter study owner or administrator password to %s the study. Use administrator password to %s multiple studies at once.", bLock ? L"lock" : L"unlock",  bLock ? L"lock" : L"unlock");
					}
					else
					{
						strMsg.Format(L"Enter study owner or administrator password to %s the study.", bLock ? L"unlock" : L"lock");
					}

					for (iterOper = setStandard.begin();
						iterOper != setStandard.end();
						++iterOper)
					{
						const CString &strOper = *iterOper;

						CComPtr<IVsiOperator> pOperator;

						hr = m_pOperatorMgr->GetOperator(
							VSI_OPERATOR_PROP_NAME,
							strOper,
							&pOperator);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);		

						if (NULL == pOperator)
						{
							// Owner gone

							// Check whether there are any admin with password
							if (0 == dwAdminWithPassowrd)
							{
								// Owner gone, no admin with password, just let the user do it
								continue;
							}
						}

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
							&(pOperator.p));
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					
						if (S_OK == hr)
						{
							VSI_OPERATOR_TYPE dwType;
							hr = pOperator->GetType(&dwType);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if (VSI_OPERATOR_TYPE_ADMIN == dwType)
							{
								break;
							}
						}
						else
						{
							// Cancel
							return 0;
						}
					}
				}
			}

			// Finally
			{
				vLock = (FALSE == bLock);
				CComVariant vLockOld = (FALSE != bLock);

				std::vector<IVsiStudy*>::iterator iter;
				for (iter = vecStudies.begin(); iter != vecStudies.end(); ++iter)
				{
					IVsiStudy *pStudyTest = *iter;

					CComHeapPtr<OLECHAR> pszOperatorId;
					CComPtr<IVsiStudy> pActiveStudy;
					CComPtr<IVsiSeries> pActiveSeries;
					CComPtr<IVsiSeries> pSeries;
					hr = m_pSession->GetPrimarySeries(&pActiveSeries);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);


					if (pActiveSeries)
					{
						hr = pActiveSeries->GetStudy(&pActiveStudy);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						if (S_OK == hr)
						{
							CComVariant vActiveId;
							hr = pActiveStudy->GetProperty(VSI_PROP_STUDY_ID, &vActiveId);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							if (S_OK == hr)
							{
								CComVariant vId;
								hr = pStudyTest->GetProperty(VSI_PROP_STUDY_ID, &vId);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								if (S_OK == hr)
								{
									if (vActiveId == vId)
									{
										CComVariant vStudyName;
										hr = pStudyTest->GetProperty(VSI_PROP_STUDY_NAME, &vStudyName);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

										VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(
											L"Study %s was not locked for containing open series", V_BSTR(&vStudyName)));

										VsiMessageBox(
											GetTopLevelParent(),
											VsiFormatMsg(L"%s contains an open series. \n\nPlease close the series before locking the study", V_BSTR(&vStudyName)),
											VsiGetRangeValue<CString>(
												VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
												VSI_PARAMETER_SYS_PRODUCT_NAME,
												m_pPdm),
											MB_OK | MB_ICONINFORMATION);

										continue;
									}
								}
							}
						}
					}

					hr = pStudyTest->SetProperty(VSI_PROP_STUDY_LOCKED, &vLock);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComQIPtr<IVsiPersistStudy> pStudyPersist(pStudyTest);
					hr = pStudyPersist->SaveStudyData(NULL, 0);
					if (FAILED(hr))
					{
						hr = pStudyTest->SetProperty(VSI_PROP_STUDY_LOCKED, &vLockOld);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}

				::InvalidateRect(pnmh->hwndFrom, NULL, FALSE);

				// Signals event
				VsiSetParamEvent(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
					VSI_PARAMETER_EVENTS_GENERAL_STUDY_LOCK_UPDATE, m_pPdm);
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnListItemRClick(int, LPNMHDR pnmh, BOOL& bHandled)
{
	bHandled = FALSE;

	LPNMITEMACTIVATE pnmia = (LPNMITEMACTIVATE)pnmh;

	if (pnmia->iItem < 0)  // Clicks nowhere
	{
		return 0;
	}

	HRESULT hr = S_OK;
	HMENU hMenuTop(NULL);

	try
	{
		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			ULONG ulSysState = VsiGetBitfieldValue<ULONG>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_SYS_STATE, m_pPdm);

			hMenuTop = LoadMenu(
				_AtlBaseModule.GetResourceInstance(),
				MAKEINTRESOURCE(IDR_MENU_STUDY_ITEM));
			VSI_WIN32_VERIFY(NULL != hMenuTop, VSI_LOG_ERROR, NULL);

			HMENU hMenu = GetSubMenu(hMenuTop, 0);
			VSI_VERIFY(NULL != hMenu, VSI_LOG_ERROR, NULL);
		
			SetMenuDefaultItem(hMenu, ID_SB_OPEN, MF_BYCOMMAND);

			if (0 == (ulSysState & VSI_SYS_STATE_ENGINEER_MODE))
			{
				DeleteMenu(hMenu, ID_SB_EXPLORE, MF_BYCOMMAND);
				DeleteMenu(hMenu, ID_SB_DELETE_DCM, MF_BYCOMMAND);
				DeleteMenu(hMenu, ID_SB_FIX_SAX_LV_CALC, MF_BYCOMMAND);
			}

			int iStudy(0), iSeries(0), iImage(0);
			VSI_DATA_LIST_COLLECTION listSel;
			CComPtr<IVsiEnumImage> pEnumImage;
			CComPtr<IUnknown> pUnkItem;
			VSI_DATA_LIST_TYPE type(VSI_DATA_LIST_NONE);

			memset(&listSel, 0, sizeof(listSel));
			listSel.dwFlags = VSI_DATA_LIST_FILTER_SELECT_CHILDREN;
			listSel.piStudyCount = &iStudy;
			listSel.piSeriesCount = &iSeries;
			listSel.piImageCount = &iImage;
			listSel.ppEnumImage = &pEnumImage;
			hr = m_pDataList->GetSelection(&listSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Gets focused item
			int iFocused(-1);
			hr = m_pDataList->GetSelectedItem(&iFocused);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (iFocused >= 0)
			{
				hr = m_pDataList->GetItem(iFocused, &type, &pUnkItem);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			if (0 < iStudy || 0 < iSeries || 1 != iImage)
			{
				iImage = 0;
			}
			else
			{
				// Single
				if (VSI_DATA_LIST_IMAGE != type)
				{
					iImage = 0;
				}
			}

			if (0 == iImage)
			{
				DeleteMenu(hMenu, ID_SB_OPEN, MF_BYCOMMAND);
				DeleteMenu(hMenu, ID_SB_IMAGELABEL, MF_BYCOMMAND);
				DeleteMenu(hMenu, ID_SB_DELETE_DCM, MF_BYCOMMAND);
			}
			else
			{
				// Don't allow image label if image is not supported
				CComQIPtr<IVsiImage> pImage(pUnkItem);
				if (NULL != pImage)
				{
					VSI_IMAGE_ERROR dwError;
					hr = pImage->GetErrorCode(&dwError);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (VSI_IMAGE_ERROR_NONE != dwError)
					{
						DeleteMenu(hMenu, ID_SB_IMAGELABEL, MF_BYCOMMAND);
					}
				}
			}

			// Get lock state
			BOOL bLockAll = VsiGetBooleanValue<BOOL>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
				VSI_PARAMETER_SYS_LOCK_ALL, m_pPdm);

			// TODO: remove ID_SB_DELETE_DCM if not VevoStrain or VevoCQ

			// Move series
			listSel.dwFlags = VSI_DATA_LIST_FILTER_NONE;
			listSel.ppEnumImage = NULL;
			hr = m_pDataList->GetSelection(&listSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (1 < iSeries || type != VSI_DATA_LIST_SERIES)
			{
				DeleteMenu(hMenu, ID_SB_MOVE_SERIES, MF_BYCOMMAND);
			}

			// Switch measurement package
			int iAllImages(0);
			CComPtr<IVsiEnumImage> pEnumImageAll;

			memset(&listSel, 0, sizeof(listSel));
			listSel.dwFlags = VSI_DATA_LIST_FILTER_SKIP_ERRORS_AND_WARNINGS | VSI_DATA_LIST_FILTER_MOVE_TO_CHILDREN;
			listSel.ppEnumImage = &pEnumImageAll;
			listSel.piImageCount = &iAllImages;
			hr = m_pDataList->GetSelection(&listSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (0 == iAllImages)
			{
				DeleteMenu(hMenu, ID_SB_CHANGE_MSMNT_PACKAGE, MF_BYCOMMAND);
			}

			iStudy = 0;
			CComPtr<IVsiEnumStudy> pEnumStudy;
			memset(&listSel, 0, sizeof(listSel));
			listSel.dwFlags = VSI_DATA_LIST_FILTER_SKIP_ERRORS_AND_WARNINGS | VSI_DATA_LIST_FILTER_SELECT_PARENT;
			listSel.piStudyCount = &iStudy;
			listSel.ppEnumStudy = &pEnumStudy;
			hr = m_pDataList->GetSelection(&listSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
				VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);

			BOOL bAdminAuthenticated(FALSE);
			if (bSecureMode)
			{
				hr = m_pOperatorMgr->GetIsAdminAuthenticated(&bAdminAuthenticated);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			BOOL bAllStudiesLocked(TRUE);
			if (pEnumStudy != NULL)
			{
				CComPtr<IVsiStudy> pStudy;
				while (S_OK == pEnumStudy->Next(1, &pStudy, NULL))
				{
					CComVariant vLocked(false);
					hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK != hr || VARIANT_FALSE == V_BOOL(&vLocked))
					{
						bAllStudiesLocked = FALSE;
						break;
					}

					pStudy.Release();
				}

				hr = pEnumStudy->Reset();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			BOOL bAllImagesLocked(TRUE);
			if (pEnumImageAll != NULL)
			{
				CComPtr<IVsiImage> pImage;
				while (S_OK == pEnumImageAll->Next(1, &pImage, NULL))
				{
					CComPtr<IVsiStudy> pStudy;
					hr = pImage->GetStudy(&pStudy);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vLocked(false);
					hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK != hr || VARIANT_FALSE == V_BOOL(&vLocked))
					{
						bAllImagesLocked = FALSE;
						break;
					}

					pImage.Release();
				}

				hr = pEnumImageAll->Reset();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			// Batch change study information - available to User Management Mode
			if (bSecureMode)
			{				
				if (0 < iStudy)
				{
					// Gets focused item
					iFocused = -1;
					type = VSI_DATA_LIST_NONE;
					hr = m_pDataList->GetSelectedItem(&iFocused);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (iFocused >= 0)
					{
						pUnkItem = NULL;
						hr = m_pDataList->GetItem(iFocused, &type, &pUnkItem);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (VSI_DATA_LIST_STUDY != type)
						{
							DeleteMenu(hMenu, ID_SB_CHANGE_STUDIES_INFO, MF_BYCOMMAND);
						}
						else
						{
							BOOL bAnySameOwner(FALSE);
							CComHeapPtr<OLECHAR> pszName;

							// Admin can always change sharing
							if (!bAdminAuthenticated)
							{
								// For standard operator, allows changing of sharing if the current operator
								// owned any one of the selected studies
								CComPtr<IVsiOperator> pOperatorSel;
								hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorSel);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								if (S_OK == hr)
								{
									hr = pOperatorSel->GetName(&pszName);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								}

								CComPtr<IVsiStudy> pStudy;
								while (S_OK == pEnumStudy->Next(1, &pStudy, NULL))
								{
									if (!bAnySameOwner)
									{
										CComVariant vOwner(L"");
										hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

										if (S_OK == hr && (0 != *V_BSTR(&vOwner)))
										{
											if (0 == wcscmp(pszName, V_BSTR(&vOwner)))
											{
												bAnySameOwner = TRUE;
												break;
											}
										}
									}

									pStudy.Release();
								}

								pEnumStudy->Reset();
							}
		
							if (!bAdminAuthenticated && !bAnySameOwner)
							{
								// Disable command
								EnableMenuItem(hMenu, ID_SB_CHANGE_STUDIES_INFO, MF_BYCOMMAND | MF_GRAYED);
							}
						}
					}
				}
				else
				{
					// Remove command
					DeleteMenu(hMenu, ID_SB_CHANGE_STUDIES_INFO, MF_BYCOMMAND);
				}
			}
			else
			{
				// Remove command
				DeleteMenu(hMenu, ID_SB_CHANGE_STUDIES_INFO, MF_BYCOMMAND);
			}


			BOOL bAnyEkvFilesExist(FALSE);
			BOOL bEkvStudyLocked(FALSE);
			if (pEnumImageAll != NULL)
			{
				CComPtr<IVsiImage> pImage;
				while (S_OK == pEnumImageAll->Next(1, &pImage, NULL))
				{
					if (VsiEkvUtils::CheckEkvStudyFilesExist(pImage))
					{
						bAnyEkvFilesExist = TRUE;

						hr = pImage->GetProperty(VSI_PROP_IMAGE_ID, &m_vIdSelected);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IVsiSeries> pSeries;
						hr = pImage->GetSeries(&pSeries);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IVsiStudy> pStudy;
						hr = pSeries->GetStudy(&pStudy);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComVariant vLocked(false);
						hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
						if (VARIANT_FALSE != V_BOOL(&vLocked))
						{
							bEkvStudyLocked = TRUE;
						}

						break;
					}
					pImage.Release();
				}

				hr = pEnumImageAll->Reset();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			if (!bAnyEkvFilesExist)
			{
				DeleteMenu(hMenu, ID_SB_DELETE_EKV_RAW, MF_BYCOMMAND);
			}
			VsiCleanMenu(hMenu);

			POINT ptMenu;
			GetCursorPos(&ptMenu);

			int iCmd = TrackPopupMenu(
				hMenu, TPM_RIGHTBUTTON | TPM_NONOTIFY | TPM_RETURNCMD,
				ptMenu.x, ptMenu.y, 0, GetTopLevelParent(), NULL);

			switch (iCmd)
			{
			case ID_SB_OPEN:
				{
					StopPreview();

					CComPtr<IVsiImage> pImage;
					if (pEnumImage->Next(1, &pImage, NULL) == S_OK)
					{
						Open(pImage);
					}
				}
				break;
			case ID_SB_EXPLORE:
				{
					CComHeapPtr<OLECHAR> pszPath;

					switch (type)
					{
					case VSI_DATA_LIST_STUDY:
						{
							CComQIPtr<IVsiStudy> pStudy(pUnkItem);
							VSI_CHECK_INTERFACE(pStudy, VSI_LOG_ERROR, NULL);
						
							hr = pStudy->GetDataPath(&pszPath);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
						break;
					case VSI_DATA_LIST_SERIES:
						{
							CComQIPtr<IVsiSeries> pSeries(pUnkItem);
							VSI_CHECK_INTERFACE(pSeries, VSI_LOG_ERROR, NULL);

							hr = pSeries->GetDataPath(&pszPath);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
						break;
					case VSI_DATA_LIST_IMAGE:
						{
							CComQIPtr<IVsiImage> pImage(pUnkItem);
							VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

							hr = pImage->GetDataPath(&pszPath);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
						break;
					}
				
					if (pszPath != NULL)
					{
						CString strCmd;
						strCmd.Format(L"explorer /select,%s", pszPath);
					
						STARTUPINFO si = {sizeof(STARTUPINFO)};
						PROCESS_INFORMATION pi = { 0 };

						BOOL bRet = CreateProcess(
							NULL, 
							(LPWSTR)(LPCWSTR)strCmd, 
							NULL, NULL, FALSE, 
							NORMAL_PRIORITY_CLASS, 
							NULL, 
							NULL,
							&si,
							&pi
						);

						if (bRet)
						{
							CloseHandle(pi.hThread);
							CloseHandle(pi.hProcess);
						}
					}
				}
				break;
			case ID_SB_IMAGELABEL:
				{
					m_pCmdMgr->InvokeCommand(ID_CMD_IMAGE_LABEL, NULL);
				}
				break;
			case ID_SB_DELETE:
				{
					SendMessage(WM_COMMAND, IDC_SB_DELETE);
				}
				break;
			case ID_SB_DELETE_EKV_RAW:
				{
					if (bEkvStudyLocked)
					{
						VsiMessageBox(
							GetTopLevelParent(),
							CString(MAKEINTRESOURCE(IDS_STYBRW_STUDY_LOCKED_CANNOT_MODIFY)),
							VsiGetRangeValue<CString>(
								VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
								VSI_PARAMETER_SYS_PRODUCT_NAME,
								m_pPdm),
							MB_OK | MB_ICONINFORMATION);
					}
					else
					{
						hr = DeleteEkvRaw(pEnumImageAll);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}
				break;
			case ID_SB_DELETE_DCM:
				{
					BOOL bStudyLocked = FALSE;
					// Check study lock state
					if (bLockAll)
					{
						int iFocused(-1);
						hr = m_pDataList->GetSelectedItem(&iFocused);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (iFocused >= 0)
						{
							VSI_DATA_LIST_TYPE type;
							CComPtr<IUnknown> pUnkItem;
							hr = m_pDataList->GetItem(iFocused, &type, &pUnkItem);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						
							if (VSI_DATA_LIST_IMAGE == type)
							{
								CComQIPtr<IVsiImage> pImage(pUnkItem);
								VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

								pImage->GetProperty(VSI_PROP_IMAGE_ID, &m_vIdSelected);

								CComPtr<IVsiSeries> pSeries;
								hr = pImage->GetSeries(&pSeries);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								CComPtr<IVsiStudy> pStudy;
								hr = pSeries->GetStudy(&pStudy);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								CComVariant vLocked(false);
								hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
								if (VARIANT_TRUE == V_BOOL(&vLocked))
								{
									bStudyLocked = TRUE;
								}
							}
						}
					}

					if (bStudyLocked)
					{
						VsiMessageBox(
							GetTopLevelParent(),
							CString(MAKEINTRESOURCE(IDS_STYBRW_IMAGE_LOCKED_CANNOT_MODIFY)),
							VsiGetRangeValue<CString>(
								VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
								VSI_PARAMETER_SYS_PRODUCT_NAME,
								m_pPdm),
							MB_OK | MB_ICONINFORMATION);
					}
					else
					{
						if (VSI_DATA_LIST_IMAGE == type)
						{
							HANDLE hFind = INVALID_HANDLE_VALUE;
							WCHAR szSearch[MAX_PATH], szFile[MAX_PATH];
							WIN32_FIND_DATA fd;

							CComQIPtr<IVsiImage> pImage(pUnkItem);
							VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

							CComHeapPtr<OLECHAR> pszPath;
							hr = pImage->GetDataPath(&pszPath);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							lstrcpy(szSearch, pszPath);
							PathRemoveFileSpec(pszPath);
							PathRemoveExtension(szSearch);

							CComVariant v(false);

							// Check for VevoStrain
							CComVariant vVevoStrain;
							hr = pImage->GetProperty(VSI_PROP_IMAGE_VEVOSTRAIN, &vVevoStrain);
							if (hr == S_OK)
							{
								if (V_BOOL(&vVevoStrain) == VARIANT_TRUE)
								{
									lstrcat(szSearch, L".dcm.*");
									pImage->SetProperty(VSI_PROP_IMAGE_VEVOSTRAIN, &v);
								}
							}
							else
							{
								// Check for SonoVevo
								CComVariant vSonoVevo;
								hr = pImage->GetProperty(VSI_PROP_IMAGE_VEVOCQ, &vSonoVevo);
								if (hr == S_OK)
								{
									if (V_BOOL(&vSonoVevo) == VARIANT_TRUE)
									{
										lstrcat(szSearch, L".sono*");
										pImage->SetProperty(VSI_PROP_IMAGE_VEVOCQ, &v);
									}
								}
								// Don't know what we have here, check for both
								else
								{
									lstrcat(szSearch, L"*.dcm.*");
									pImage->SetProperty(VSI_PROP_IMAGE_VEVOSTRAIN, &v);
									pImage->SetProperty(VSI_PROP_IMAGE_VEVOCQ, &v);
								}
							}

							hFind = FindFirstFile(szSearch, &fd);

							while (INVALID_HANDLE_VALUE != hFind)
							{
								if (fd.cFileName[0] == L'.')
								{
									NULL;		// skip the dot.
								}
								else if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
								{
									NULL;		// skip folder
								}
								else
								{
									lstrcpy(szFile, pszPath);
									lstrcat(szFile, L"\\");
									lstrcat(szFile, fd.cFileName);
									DeleteFile(szFile);
								}

								if (!FindNextFile(hFind, &fd))
								{
									// Done with the current directory. Close the handle and return.
									FindClose(hFind);
									hFind = INVALID_HANDLE_VALUE;
									break;
								}
							}

							if (INVALID_HANDLE_VALUE != hFind)
								FindClose(hFind);
						}
					}
				}
				break;
			case ID_SB_MOVE_SERIES:
				{
					// Check if the study is locked remove "Move" option
					CComQIPtr<IVsiSeries> pSeries(pUnkItem);
					VSI_CHECK_INTERFACE(pSeries, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiStudy> pStudy;
					hr = pSeries->GetStudy(&pStudy);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vLocked(false);
					hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
					if (VARIANT_TRUE == V_BOOL(&vLocked))
					{
						// Display a message
						m_pAppView->SetPopupText(
							CString(MAKEINTRESOURCE(IDS_STYBRW_UNLOCK_BEFORE_MOVING_SERIES)),
							1500);
					}
					else
					{
						CComVariant v(pUnkItem);
						m_pCmdMgr->InvokeCommand(ID_CMD_OPEN_MOVE_SERIES, &v);
					}
				}
				break;
			case ID_SB_CHANGE_MSMNT_PACKAGE:
				{
					if (bLockAll && bAllImagesLocked)
					{
						VsiMessageBox(
							GetTopLevelParent(),
							L"The selected images are locked and cannot be modified.",
							L"Change Measurement Package",
							MB_OK | MB_ICONEXCLAMATION);
					}
					else
					{
						CVsiChangeMsmntPackageDlg dlg(m_pApp, iAllImages, pEnumImageAll);
						dlg.DoModal(*this);
					}
				}
				break;
			case ID_SB_CHANGE_STUDIES_INFO:
				{
					if (bLockAll && bAllStudiesLocked)
					{
						VsiMessageBox(
							GetTopLevelParent(),
							L"The selected studies are locked and cannot be modified.",
							L"Change Study Access",
							MB_OK | MB_ICONEXCLAMATION);
					}
					else
					{
						CVsiChangeStudyInfoDlg dlg(
							m_pPdm,
							m_pOperatorMgr,
							m_pStudyMgr,
							m_pSession,
							pEnumStudy,
							bAdminAuthenticated);
						dlg.DoModal(*this);
					}
				}
				break;
			case ID_SB_FIX_SAX_LV_CALC:
				{
					hr = FixSaxCalc(pEnumImageAll);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
				break;
			}
		}
	}
	VSI_CATCH(hr);

	if (NULL != hMenuTop)
	{
		DestroyMenu(hMenuTop);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnListSortChanged(int, LPNMHDR, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		// Save sorting settings
		VSI_DATA_LIST_COL col;
		BOOL bDecending;

		// Study
		m_pDataList->GetSortSettings(VSI_DATA_LIST_STUDY, &col, &bDecending);

		VsiSetDiscreteValue<long>(
			(long)col,
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_STUDY_ITEM_STUDYBROWSER, m_pPdm);

		VsiSetDiscreteValue<long>(
			bDecending ? VSI_SYS_SORT_ORDER_DESCENDING : VSI_SYS_SORT_ORDER_ASCENDING,
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_STUDY_ORDER_STUDYBROWSER, m_pPdm);

		// Series
		int iSeriesColOld = VsiGetDiscreteValue<long>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_SERIES_ITEM_STUDYBROWSER, m_pPdm);
		int iSeriesOrderOld = VsiGetDiscreteValue<long>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_SERIES_ORDER_STUDYBROWSER, m_pPdm);

		m_pDataList->GetSortSettings(VSI_DATA_LIST_SERIES, &col, &bDecending);

		int iSeriesColNew = (int)col;
		int iSeriesOrderNew = bDecending ? VSI_SYS_SORT_ORDER_DESCENDING : VSI_SYS_SORT_ORDER_ASCENDING;

		VsiSetDiscreteValue<long>(
			(long)col,
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_SERIES_ITEM_STUDYBROWSER, m_pPdm);

		VsiSetDiscreteValue<long>(
			bDecending ? VSI_SYS_SORT_ORDER_DESCENDING : VSI_SYS_SORT_ORDER_ASCENDING,
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_SERIES_ORDER_STUDYBROWSER, m_pPdm);

		// Image
		int iImageColOld = VsiGetDiscreteValue<long>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_IMAGE_ITEM_STUDYBROWSER, m_pPdm);
		int iImageOrderOld = VsiGetDiscreteValue<long>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_IMAGE_ORDER_STUDYBROWSER, m_pPdm);

		m_pDataList->GetSortSettings(VSI_DATA_LIST_IMAGE, &col, &bDecending);

		int iImageColNew = (int)col;
		int iImageOrderNew = bDecending ? VSI_SYS_SORT_ORDER_DESCENDING : VSI_SYS_SORT_ORDER_ASCENDING;

		VsiSetDiscreteValue<long>(
			(long)col,
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_IMAGE_ITEM_STUDYBROWSER, m_pPdm);

		VsiSetDiscreteValue<long>(
			bDecending ? VSI_SYS_SORT_ORDER_DESCENDING : VSI_SYS_SORT_ORDER_ASCENDING,
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_UI_SORT_IMAGE_ORDER_STUDYBROWSER, m_pPdm);

		// Refresh preview
		if (VSI_DATA_LIST_STUDY == m_typeSel)
		{
			if (iSeriesColNew != iSeriesColOld || iSeriesOrderNew != iSeriesOrderOld)
			{
				CComQIPtr<IVsiStudy> pStudy(m_pItemPreview);
				VSI_CHECK_INTERFACE(pStudy, VSI_LOG_ERROR, NULL);

				m_studyOverview.SetStudy(
					pStudy,
					GetSeriesPropIdFromColumnId(iSeriesColNew),
					(VSI_SYS_SORT_ORDER_DESCENDING == iSeriesOrderNew),
					true);
			}
		}
		else if (VSI_DATA_LIST_SERIES == m_typeSel)
		{
			if (iImageColNew != iImageColOld || iImageOrderNew != iImageOrderOld)
			{
				if (NULL != m_pItemPreview)
				{
					StopPreview(VSI_PREVIEW_LIST_CLEAR);
					StartPreview(TRUE);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnListGetItemStatus(int, LPNMHDR pnmh, BOOL&)
{
	HRESULT hr = S_OK;
	LRESULT lRet = VSI_DATA_LIST_ITEM_STATUS_NORMAL;

	try
	{
		NM_DL_ITEM *pItem = (NM_DL_ITEM*)pnmh;
		switch (pItem->type)
		{
		case VSI_DATA_LIST_STUDY:
			{
				CComQIPtr<IVsiStudy> pStudy(pItem->pUnkItem);
				if (NULL != pStudy)  // Error checking
				{
					CComPtr<IVsiStudy> pStudyActive;
					hr = m_pSession->GetPrimaryStudy(&pStudyActive);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (pStudyActive != NULL)
					{
						CComVariant v1, v2;
						pStudy->GetProperty(VSI_PROP_STUDY_ID, &v1);
						pStudyActive->GetProperty(VSI_PROP_STUDY_ID, &v2);
						if (v1 != v2)
						{
							lRet |= VSI_DATA_LIST_ITEM_STATUS_INACTIVE;
						}
					}
				}
			}
			break;
		case VSI_DATA_LIST_SERIES:
			{
				CComQIPtr<IVsiSeries> pSeries(pItem->pUnkItem);
				if (NULL != pSeries)  // Error checking
				{
					CComPtr<IVsiSeries> pSeriesActive;
					hr = m_pSession->GetPrimarySeries(&pSeriesActive);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (pSeriesActive != NULL)
					{
						CComVariant vNs1, vNs2;
						pSeries->GetProperty(VSI_PROP_SERIES_NS, &vNs1);
						pSeriesActive->GetProperty(VSI_PROP_SERIES_NS, &vNs2);
						if (vNs1 != vNs2)
						{
							lRet |= VSI_DATA_LIST_ITEM_STATUS_INACTIVE;
						}
					}
				}
			}
			break;
		case VSI_DATA_LIST_IMAGE:
			{
				CComQIPtr<IVsiImage> pImage(pItem->pUnkItem);
				if (NULL != pImage)  // Error checking
				{
					CComVariant vNs;
					hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
					if (S_OK == hr)
					{
						// Not corrupted

						// Active?
						int iSlot(-1);
						HRESULT hrLoaded = m_pSession->GetIsImageLoaded(V_BSTR(&vNs), &iSlot);
						if (S_OK == hrLoaded)
						{
							LONG lLayout = VsiGetDiscreteValue<LONG>(
								VSI_PDM_ROOT_APP,
								VSI_PDM_GROUP_SETTINGS,
								VSI_PARAMETER_UI_DISPLAY_LAYOUT,
								m_pPdm);
							if (VSI_MODE_LAYOUT_DUAL == lLayout)
							{
								lRet |= (0 == iSlot) ?
									VSI_DATA_LIST_ITEM_STATUS_ACTIVE_LEFT :
									VSI_DATA_LIST_ITEM_STATUS_ACTIVE_RIGHT;
							}
							else
							{
								lRet |= VSI_DATA_LIST_ITEM_STATUS_ACTIVE_SINGLE;
							}
						}
						else
						{
							CComPtr<IVsiSeries> pSeriesActive;
							hr = m_pSession->GetPrimarySeries(&pSeriesActive);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						
							if (pSeriesActive != NULL)
							{
								CComPtr<IVsiSeries> pSeries;
								hr = pImage->GetSeries(&pSeries);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							
								if (!pSeriesActive.IsEqualObject(pSeries))
								{
									lRet |= VSI_DATA_LIST_ITEM_STATUS_INACTIVE;
								}
							}

							HRESULT hrReviewed = m_pSession->GetIsItemReviewed(V_BSTR(&vNs));
							if (S_OK == hrReviewed)
							{
								lRet |= VSI_DATA_LIST_ITEM_STATUS_REVIEWED;
							}
						}
					}
				}
			}
			break;
		default:
			_ASSERT(0);
		}
	}
	VSI_CATCH(hr);

	return lRet;
}

LRESULT CVsiStudyBrowser::OnPreviewListDbClick(int, LPNMHDR pnmh, BOOL&)
{
	NMITEMACTIVATE *pnmia = (NMITEMACTIVATE*)pnmh;
	if (pnmia->iItem >= 0)
	{
		StopPreview();

		LVITEMW lvi = { LVIF_PARAM };
		lvi.iItem = pnmia->iItem;

		ListView_GetItem(pnmh->hwndFrom, &lvi);
		if (lvi.lParam == NULL)
			return 0;

		Open((IVsiImage*)lvi.lParam);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnPreviewListItemChanged(int, LPNMHDR pnmh, BOOL&)
{
	HRESULT hr;

	try
	{
		NMLISTVIEW *pnmlv = (NMLISTVIEW*)pnmh;
		if (LVIS_SELECTED == pnmlv->uNewState && pnmlv->iItem >= 0)
		{
			IVsiImage *pImage = (IVsiImage*)pnmlv->lParam;

			if (GetFocus() == pnmh->hwndFrom)  // By user clicking
			{
				CComVariant vNs;
				hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = SetSelection(V_BSTR(&vNs));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			CComQIPtr<IVsiSeries> pSeries(m_pItemPreview);
			if (pSeries != NULL)
			{
				LONG lCount(0);
				hr = pSeries->GetImageCount(&lCount, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
				CString strText;
				strText.Format(IDS_STYBRW_IMAGE_COUNT_STATUS, lCount - pnmlv->iItem, lCount);

				m_pAppView->SetStatusText(strText, 2000);
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyBrowser::OnPreviewListKeyDown(int, LPNMHDR pnmh, BOOL&)
{
	NMLVKEYDOWN *pnmvkd = (NMLVKEYDOWN*)pnmh;

	if (VK_RETURN == pnmvkd->wVKey)
	{
		int iSel = ListView_GetNextItem(
			pnmh->hwndFrom, -1, LVNI_SELECTED);

		if (iSel >= 0)
		{
			StopPreview();

			LVITEMW lvi = { LVIF_PARAM };
			lvi.iItem = iSel;

			ListView_GetItem(pnmh->hwndFrom, &lvi);
			if (lvi.lParam == NULL)
				return 0;

			Open((IVsiImage*)lvi.lParam);
		}
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnPreviewListDeleteItem(int, LPNMHDR pnmh, BOOL&)
{
	LPNMLISTVIEW plv = (LPNMLISTVIEW)pnmh;

	LVITEMW lvi = { LVIF_PARAM };
	lvi.iItem = plv->iItem;

	ListView_GetItem(pnmh->hwndFrom, &lvi);

	IVsiImage *pImage = (IVsiImage*)lvi.lParam;
	pImage->Release();

	return 0;
}

LRESULT CVsiStudyBrowser::OnTrackBarMove(int, LPNMHDR pnmh, BOOL&)
{
	LPVSI_TRACKBAR_MOVE_NMHDR ptbm = (LPVSI_TRACKBAR_MOVE_NMHDR)pnmh;

	CRect rc;
	GetClientRect(&rc);

	SetPreviewSize((m_bVertical ? rc.bottom : rc.right) - ptbm->iPos - m_iTracker - m_iBorder);
	
	if (ptbm->bFinal)
	{
		GetDlgItem(IDC_SB_PREVIEW).SendMessage(LVM_REDRAWITEMS);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnStudyAccessFilterDropDown(int, LPNMHDR, BOOL&)
{
	HRESULT hr = S_OK;
	HMENU hMenuTop(NULL);
	HBITMAP hbmpSelf(NULL);
	HBITMAP hbmpGroup(NULL);
	HBITMAP hbmpAll(NULL);

	try
	{
		VSI_UI_STUDY_ACCESS_FILTER filterType =
			(VSI_UI_STUDY_ACCESS_FILTER)VsiGetDiscreteValue<DWORD>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_STUDY_ACCESS_FILTER, m_pPdm);

		VSI_UI_STUDY_ACCESS_FILTER filterTypeNew = filterType;

		UINT idDefault = IDB_STUDY_ACCESS_PRIVATE;
		switch (filterType)
		{
		case VSI_UI_STUDY_ACCESS_FILTER_PRIVATE:
			idDefault = ID_STUDY_ACCESS_FILTER_PRIVATE;
			break;
		case VSI_UI_STUDY_ACCESS_FILTER_GROUP:
			idDefault = ID_STUDY_ACCESS_FILTER_GROUP;
			break;
		case VSI_UI_STUDY_ACCESS_FILTER_ALL:
			idDefault = ID_STUDY_ACCESS_FILTER_ALL;
			break;
		default:
			_ASSERT(0);
		}

		hMenuTop = LoadMenu(
			_AtlBaseModule.GetResourceInstance(),
			MAKEINTRESOURCE(IDR_MENU_STUDY_ACCESS_FILTER));
		VSI_WIN32_VERIFY(NULL != hMenuTop, VSI_LOG_ERROR, NULL);

		HMENU hMenu = GetSubMenu(hMenuTop, 0);
		VSI_VERIFY(NULL != hMenu, VSI_LOG_ERROR, NULL);
		
		SetMenuDefaultItem(hMenu, idDefault, MF_BYCOMMAND);

		// Images
		hbmpSelf = (HBITMAP)LoadImage(
			_AtlBaseModule.GetResourceInstance(),
			MAKEINTRESOURCE(IDB_STUDY_ACCESS_PRIVATE),
			IMAGE_BITMAP,
			0, 0,
			LR_CREATEDIBSECTION);
		VSI_WIN32_VERIFY(NULL != hbmpSelf, VSI_LOG_ERROR, NULL);

		hr = VsiBitmapPremultipliedAlpha(hbmpSelf);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hbmpGroup = (HBITMAP)LoadImage(
			_AtlBaseModule.GetResourceInstance(),
			MAKEINTRESOURCE(IDB_STUDY_ACCESS_GROUP),
			IMAGE_BITMAP,
			0, 0,
			LR_CREATEDIBSECTION);
		VSI_WIN32_VERIFY(NULL != hbmpGroup, VSI_LOG_ERROR, NULL);

		hr = VsiBitmapPremultipliedAlpha(hbmpGroup);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hbmpAll = (HBITMAP)LoadImage(
			_AtlBaseModule.GetResourceInstance(),
			MAKEINTRESOURCE(IDB_STUDY_ACCESS_ALL),
			IMAGE_BITMAP,
			0, 0,
			LR_CREATEDIBSECTION);
		VSI_WIN32_VERIFY(NULL != hbmpAll, VSI_LOG_ERROR, NULL);

		hr = VsiBitmapPremultipliedAlpha(hbmpAll);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		MENUITEMINFO mii = { sizeof(mii) };
		mii.fMask = MIIM_BITMAP;

		mii.hbmpItem = hbmpSelf;
		BOOL bRet = SetMenuItemInfo(hMenu, ID_STUDY_ACCESS_FILTER_PRIVATE, FALSE, &mii);
		VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, NULL);

		mii.hbmpItem = hbmpGroup;
		bRet = SetMenuItemInfo(hMenu, ID_STUDY_ACCESS_FILTER_GROUP, FALSE, &mii);
		VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, NULL);

		mii.hbmpItem = hbmpAll;
		bRet = SetMenuItemInfo(hMenu, ID_STUDY_ACCESS_FILTER_ALL, FALSE, &mii);
		VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, NULL);

		CRect rc;
		GetDlgItem(IDC_SB_ACCESS_FILTER).GetWindowRect(&rc);

		POINT ptMenu = { rc.left, rc.bottom };

		int iCmd = TrackPopupMenu(
			hMenu, TPM_RIGHTBUTTON | TPM_NONOTIFY | TPM_RETURNCMD,
			ptMenu.x, ptMenu.y, 0, GetTopLevelParent(), NULL);

		switch (iCmd)
		{
		case ID_STUDY_ACCESS_FILTER_PRIVATE:
			{
				filterTypeNew = VSI_UI_STUDY_ACCESS_FILTER_PRIVATE;
			}
			break;
		case ID_STUDY_ACCESS_FILTER_GROUP:
			{
				filterTypeNew = VSI_UI_STUDY_ACCESS_FILTER_GROUP;
			}
			break;
		case ID_STUDY_ACCESS_FILTER_ALL:
			{
				filterTypeNew = VSI_UI_STUDY_ACCESS_FILTER_ALL;
			}
			break;
		}

		if (filterTypeNew != filterType)
		{
			VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Study Access Filter");

			VsiSetDiscreteValue<DWORD>(
				static_cast<DWORD>(filterTypeNew),
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_STUDY_ACCESS_FILTER, m_pPdm);
		}
	}
	VSI_CATCH(hr);

	if (NULL != hMenuTop)
	{
		DestroyMenu(hMenuTop);
	}

	if (NULL != hbmpSelf)
	{
		DeleteObject(hbmpSelf);
	}
	if (NULL != hbmpGroup)
	{
		DeleteObject(hbmpGroup);
	}
	if (NULL != hbmpAll)
	{
		DeleteObject(hbmpAll);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnLogoutDropDown(int, LPNMHDR, BOOL&)
{
	HRESULT hr = S_OK;
	HMENU hMenuTop(NULL);
	HBITMAP hbmpShutdown(NULL);

	try
	{
		hMenuTop = LoadMenu(
			_AtlBaseModule.GetResourceInstance(),
			MAKEINTRESOURCE(IDR_MENU_LOGOUT));
		VSI_WIN32_VERIFY(NULL != hMenuTop, VSI_LOG_ERROR, NULL);

		HMENU hMenu = GetSubMenu(hMenuTop, 0);
		VSI_VERIFY(NULL != hMenu, VSI_LOG_ERROR, NULL);
		
		// Images
		hbmpShutdown = (HBITMAP)LoadImage(
			_AtlBaseModule.GetResourceInstance(),
			MAKEINTRESOURCE(IDB_LOGOUT_SHUTDOWN),
			IMAGE_BITMAP,
			0, 0,
			LR_CREATEDIBSECTION);
		VSI_WIN32_VERIFY(NULL != hbmpShutdown, VSI_LOG_ERROR, NULL);

		hr = VsiBitmapPremultipliedAlpha(hbmpShutdown);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		MENUITEMINFO mii = { sizeof(mii) };
		mii.fMask = MIIM_BITMAP;

		mii.hbmpItem = hbmpShutdown;
		BOOL bRet = SetMenuItemInfo(hMenu, ID_LOGOUT_SHUTDOWN, FALSE, &mii);
		VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, NULL);

		CRect rc;
		GetDlgItem(IDC_SB_LOGOUT).GetWindowRect(&rc);

		POINT ptMenu = { rc.right, rc.bottom };

		int iCmd = TrackPopupMenu(
			hMenu, TPM_RIGHTALIGN | TPM_RIGHTBUTTON | TPM_NONOTIFY | TPM_RETURNCMD,
			ptMenu.x, ptMenu.y, 0, GetTopLevelParent(), NULL);

		switch (iCmd)
		{
		case ID_LOGOUT_SHUTDOWN:
			{
				VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Shut Down");

				CComVariant vShutdown(true);
				hr = m_pCmdMgr->InvokeCommandAsync(ID_CMD_EXIT_BEGIN, &vShutdown, FALSE);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			break;
		}
	}
	VSI_CATCH(hr);

	if (NULL != hMenuTop)
	{
		DestroyMenu(hMenuTop);
	}

	if (NULL != hbmpShutdown)
	{
		DeleteObject(hbmpShutdown);
	}

	return 0;
}

LRESULT CVsiStudyBrowser::OnHdnEndDrag(int, LPNMHDR, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		hr = UpdateColumnsRecord();
		if (SUCCEEDED(hr))
		{
			VSI_LVCOLUMN *plvcolumns(NULL);
			int iColumns(0);

			plvcolumns = g_plvcStudyBrowserListView;
			iColumns =  _countof(g_plvcStudyBrowserListView);

			// Save column order
			CString strOrders;
			CString strPair;
			for (int i = 0; i < iColumns; ++i)
			{
				strPair.Format((0 == i) ? L"%d,%d" : L";%d,%d", plvcolumns[i].iSubItem, plvcolumns[i].iOrder);

				strOrders += strPair;
			}

			VsiSetRangeValue<LPCWSTR>(
				strOrders,
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_STUDY_BROWSER_COLUMNS_ORDER, m_pPdm);
		}
	}
	VSI_CATCH(hr);


	return 0;
}

LRESULT CVsiStudyBrowser::OnHdnEndTrack(int, LPNMHDR, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		hr = UpdateColumnsRecord();
		if (SUCCEEDED(hr))
		{
			VSI_LVCOLUMN *plvcolumns(NULL);
			int iColumns(0);

			plvcolumns = g_plvcStudyBrowserListView;
			iColumns =  _countof(g_plvcStudyBrowserListView);

			// Save column width
			CString strWidths;
			CString strPair;
			for (int i = 0; i < iColumns; ++i)
			{
				strPair.Format((0 == i) ? L"%d,%d" : L";%d,%d", plvcolumns[i].iSubItem, plvcolumns[i].cx);

				strWidths += strPair;
			}

			VsiSetRangeValue<LPCWSTR>(
				strWidths,
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_STUDY_BROWSER_COLUMNS_WIDTH, m_pPdm);
		}
	}
	VSI_CATCH(hr);


	return 0;
}

HRESULT CVsiStudyBrowser::VSI_STATE_HANDLER_IMP(VSI_APP_STATE_MODE_VIEW_CLOSE)
{
	HRESULT hr = S_OK;

	try
	{
		// Closing of a mode view

		// Refreshes visited and active state
		CComPtr<IVsiImage> pImage;
		hr = m_pSession->GetImage(V_I4(pvParam), &pImage);
		if (S_OK == hr)
		{
			CComVariant vNs;
			hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CCritSecLock csl(m_csUpdateItems);
			m_listUpdateItems.push_back(CString(V_BSTR(&vNs)));
		}

		PostMessage(WM_VSI_SB_UPDATE_ITEMS);
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_STATE_HANDLER_IMP(VSI_APP_STATE_MODE_CLOSE)
{
	UNREFERENCED_PARAMETER(pvParam);

	HRESULT hr = S_OK;

	m_State = VSI_SB_STATE_START;
	GetDlgItem(IDCANCEL).EnableWindow(FALSE);

	UpdateNewCloseButtons();

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_STATE_HANDLER_IMP(VSI_APP_STATE_MODE)
{
	UNREFERENCED_PARAMETER(pvParam);

	HRESULT hr = S_OK;

	try
	{
		// Starting of a mode

		// Refreshes visited and active state
		CComPtr<IVsiImage> pImage;
		hr = m_pSession->GetImage(VSI_SLOT_ACTIVE, &pImage);
		if (S_OK == hr)
		{
			CComVariant vNs;
			hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CCritSecLock csl(m_csUpdateItems);
			m_listUpdateItems.push_back(CString(V_BSTR(&vNs)));
		}

		PostMessage(WM_VSI_SB_UPDATE_ITEMS);

		m_State = VSI_SB_STATE_CANCEL;
		GetDlgItem(IDCANCEL).EnableWindow(TRUE);

		UpdateNewCloseButtons();
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_STATE_HANDLER_IMP(VSI_APP_STATE_SPLIT_SCREEN)
{
	UNREFERENCED_PARAMETER(pvParam);

	HRESULT hr = S_OK;

	try
	{
		// Refreshes active state
		for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
		{
			CComPtr<IVsiImage> pImage;
			hr = m_pSession->GetImage(i, &pImage);
			if (S_OK == hr)
			{
				CComVariant vNs;
				hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CCritSecLock csl(m_csUpdateItems);
				m_listUpdateItems.push_back(CString(V_BSTR(&vNs)));
			}
		}

		PostMessage(WM_VSI_SB_UPDATE_ITEMS);
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_STATE_HANDLER_IMP(VSI_APP_STATE_IMAGE_LABEL)
{
	UNREFERENCED_PARAMETER(pvParam);

	HRESULT hr = S_OK;

	try
	{
		if (VSI_VIEW_STATE_SHOW == m_ViewState)
		{
			// Get lock all state
			BOOL bLockAll = VsiGetBooleanValue<BOOL>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
				VSI_PARAMETER_SYS_LOCK_ALL, m_pPdm);

			BOOL bStudyTotalLocked = FALSE;
			if (bLockAll)
			{
				CVsiLockData lock(m_pDataList);
				if (lock.LockData())
				{
					int iFocused(-1);
					hr = m_pDataList->GetSelectedItem(&iFocused);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (iFocused >= 0)
					{
						VSI_DATA_LIST_TYPE type;
						CComPtr<IUnknown> pUnkItem;
						hr = m_pDataList->GetItem(iFocused, &type, &pUnkItem);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						
						if (VSI_DATA_LIST_IMAGE == type)
						{
							CComQIPtr<IVsiImage> pImage(pUnkItem);
							VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

							pImage->GetProperty(VSI_PROP_IMAGE_ID, &m_vIdSelected);

							CComPtr<IVsiSeries> pSeries;
							hr = pImage->GetSeries(&pSeries);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							CComPtr<IVsiStudy> pStudy;
							hr = pSeries->GetStudy(&pStudy);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							CComVariant vLocked(false);
							hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
							if (VARIANT_TRUE == V_BOOL(&vLocked))
							{
								bStudyTotalLocked = TRUE;
							}
						}
					}
				}
			}

			if (bStudyTotalLocked)
			{
				VsiMessageBox(
					GetTopLevelParent(),
					CString(MAKEINTRESOURCE(IDS_STYBRW_IMAGE_LOCKED_CANNOT_MODIFY)),
					VsiGetRangeValue<CString>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
						VSI_PARAMETER_SYS_PRODUCT_NAME,
						m_pPdm),
					MB_OK | MB_ICONINFORMATION);

				// Cancel state
				CComPtr<IVsiPropertyBag> pProps;
				hr = pProps.CoCreateInstance(__uuidof(VsiPropertyBag));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vCloseCode(TRUE);
				hr = pProps->Write(L"closecode", &vCloseCode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vParams(pProps.p);
				m_pCmdMgr->InvokeCommand(ID_CMD_IMAGE_LABEL, &vParams);  // Close panel
			}
			else
			{
				if (IsWindowVisible())
				{
					CVsiLockData lock(m_pDataList);
					if (lock.LockData())
					{
						// Change image label
						int iImage = 0, iSeries = 0;
						CComPtr<IVsiEnumImage> pEnumImage;

						VSI_DATA_LIST_COLLECTION listSel = { VSI_DATA_LIST_FILTER_SELECT_CHILDREN };
						listSel.piSeriesCount = &iSeries;
						listSel.ppEnumImage = &pEnumImage;
						listSel.piImageCount = &iImage;
						hr = m_pDataList->GetSelection(&listSel);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						lock.UnlockData();

						if (0 == iSeries && 1 == iImage)
						{
							CComPtr<IVsiImage> pImage;
							hr = pEnumImage->Next(1, &pImage, NULL);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

							// Don't allow image label if image is not supported
							VSI_IMAGE_ERROR dwError;
							hr = pImage->GetErrorCode(&dwError);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if (VSI_IMAGE_ERROR_NONE == dwError)
							{
								hr = m_pImageLabel.CoCreateInstance(__uuidof(VsiChangeImageLabel));
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								hr = m_pImageLabel->Activate(m_pApp, GetParent(), pImage);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							}
						}
					}
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_STATE_HANDLER_IMP(VSI_APP_STATE_IMAGE_LABEL_CLOSE)
{
	UNREFERENCED_PARAMETER(pvParam);

	HRESULT hr(S_OK);

	try
	{
		if (NULL != m_pImageLabel)
		{
			hr = m_pImageLabel->Deactivate();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (NULL != pvParam)
			{
				CComQIPtr<IVsiPropertyBag> pProps(V_UNKNOWN(pvParam));
				VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

				CComVariant vCloseCode(S_OK == hr);
				hr = pProps->Write(L"closecode", &vCloseCode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
			VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

			CComPtr<IVsiGarbageCollect> pGc;
			hr = pServiceProvider->QueryService(
				VSI_SERVICE_GC,
				__uuidof(IVsiGarbageCollect),
				(IUnknown**)&pGc);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pGc->Dispose(m_pImageLabel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_pImageLabel = NULL;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_STATE_HANDLER_IMP(VSI_APP_STATE_STUDY_BROWSER)
{
	UNREFERENCED_PARAMETER(pvParam);

	UpdateNewCloseButtons();
	
	CVsiLockData lock(m_pDataList);
	if (lock.LockData())
	{
		m_pDataList->UpdateItemStatus(NULL, TRUE);
	}

	return S_OK;
}

HRESULT CVsiStudyBrowser::VSI_STATE_HANDLER_IMP(VSI_APP_STATE_ANALYSIS_BROWSER_CLOSE)
{
	UNREFERENCED_PARAMETER(pvParam);

	// release AnalysisSet object
	if (NULL != m_pAnalysisSet)
	{
		m_pAnalysisSet->Uninitialize();
		m_pAnalysisSet.Release();
	}

	return S_OK;
}

HRESULT CVsiStudyBrowser::VSI_CMD_HANDLER_IMP(ID_CMD_UPDATE_ANALYSIS_BROWSER)
{
	UNREFERENCED_PARAMETER(pbContinueExecute);

	HRESULT hr = S_OK;

	try
	{
		// Update measurement selection dialog
		if (NULL == pvParam)
		{
			VsiSetBooleanValue(TRUE, VSI_PDM_ROOT_APP, 
				VSI_PDM_GROUP_EVENTS, 
				VSI_PARAMETER_EVENTS_MEASUREMENT_DIALOG_UPDATE,
				m_pPdm);
		}

		// Update Analysis Browser
		if (NULL != m_pAnalysisSet)
		{
			hr = GenerateFullDisclosureMsmntReport(FALSE);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Fire update event to Analysis Browser
			{
				// Get updated report xml from Analysis Set	
				CComPtr<IXMLDOMDocument> pXmlReportDoc;
				hr = m_pAnalysisSet->GetReportXml(&pXmlReportDoc);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
				CComPtr<IVsiParameter> pParam;
				hr = m_pPdm->GetParameter(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
					VSI_PARAMETER_EVENTS_GENERAL_ANALYSIS_BROWSER_UPDATE,
					&pParam);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComQIPtr<IVsiParameterRange> pUpdate(pParam);
				VSI_CHECK_INTERFACE(pUpdate, VSI_LOG_ERROR, NULL);

				CComVariant vUpdateXmlDoc((IUnknown*)pXmlReportDoc);
				hr = pUpdate->SetValue(vUpdateXmlDoc, FALSE);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_CMD_HANDLER_IMP(ID_CMD_STUDY_SERIES_INFO)
{
	UNREFERENCED_PARAMETER(pbContinueExecute);

	HRESULT hr = S_OK;

	try
	{
		if (VSI_VIEW_STATE_SHOW == m_ViewState)
		{
			if (NULL == pvParam)  // From control panel
			{
				if (GetDlgItem(IDC_SB_STUDY_SERIES_INFO).IsWindowEnabled())
				{
					SendMessage(WM_COMMAND, IDC_SB_STUDY_SERIES_INFO);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_CMD_HANDLER_IMP(ID_CMD_DELETE)
{
	UNREFERENCED_PARAMETER(pbContinueExecute);

	HRESULT hr = S_OK;

	try
	{
		if (VSI_VIEW_STATE_SHOW == m_ViewState)
		{
			if (NULL == pvParam)  // From control panel
			{
				if (GetDlgItem(IDC_SB_DELETE).IsWindowEnabled())
				{
					SendMessage(WM_COMMAND, IDC_SB_DELETE);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_CMD_HANDLER_IMP(ID_CMD_EXPORT)
{
	UNREFERENCED_PARAMETER(pbContinueExecute);

	HRESULT hr = S_OK;

	try
	{
		if (VSI_VIEW_STATE_SHOW == m_ViewState)
		{
			if (NULL == pvParam)  // From control panel
			{
				if (GetDlgItem(IDC_SB_EXPORT).IsWindowEnabled())
				{
					SendMessage(WM_COMMAND, IDC_SB_EXPORT);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_CMD_HANDLER_IMP(ID_CMD_REPORT)
{
	UNREFERENCED_PARAMETER(pbContinueExecute);
	UNREFERENCED_PARAMETER(pvParam);

	HRESULT hr = S_OK;

	try
	{
		if (VSI_VIEW_STATE_SHOW == m_ViewState)
		{
			hr = GenerateFullDisclosureMsmntReport();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::VSI_CMD_HANDLER_IMP(ID_CMD_COPY_STUDY_TO)
{
	UNREFERENCED_PARAMETER(pbContinueExecute);

	HRESULT hr = S_OK;

	try
	{
		if (VSI_VIEW_STATE_SHOW == m_ViewState)
		{
			if (NULL == pvParam)  // From control panel
			{
				if (GetDlgItem(IDC_SB_COPY_TO).IsWindowEnabled())
				{
					SendMessage(WM_COMMAND, IDC_SB_COPY_TO);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

DWORD CVsiStudyBrowser::ThreadProc()
{
	HRESULT hr = S_OK;

	try
	{
		hr = CoInitializeEx(0, COINIT_MULTITHREADED);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HANDLE phWaits[] = {
			m_hExit,
			m_hReload,
			m_hSearch
		};

		for (;;)
		{
			DWORD dwWait = WaitForMultipleObjects(_countof(phWaits), phWaits, FALSE, INFINITE);

			if ((WAIT_OBJECT_0 + VSI_LIST_EXIT) == dwWait)
			{
				// Exit
				break;
			}
			else if ((WAIT_OBJECT_0 + VSI_LIST_RELOAD) == dwWait)
			{
				try
				{
					// Refresh study list based on the latest study path

					// Get data path
					CVsiRange<LPCWSTR> pParamPathData;
					hr = m_pPdm->GetParameter(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
						VSI_PARAMETER_SYS_PATH_DATA, &pParamPathData);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComQIPtr<IVsiParameterRange> pRange(pParamPathData.m_pParam);
					VSI_CHECK_INTERFACE(pRange, VSI_LOG_ERROR, NULL);

					CComVariant vPath;
					hr = pRange->GetValue(&vPath);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					VSI_LOAD_STUDY_FILTER filter = { 0 };
					VSI_LOAD_STUDY_FILTER *pFilter(NULL);
					CComHeapPtr<OLECHAR> pszOperatorName;
					CComHeapPtr<OLECHAR> pszGroup;

					BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
						VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
					if (bSecureMode)
					{
						CComPtr<IVsiOperator> pOperator;
						hr = m_pOperatorMgr->GetCurrentOperator(&pOperator);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						hr = pOperator->GetName(&pszOperatorName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						filter.pszOperator = pszOperatorName;

						VSI_UI_STUDY_ACCESS_FILTER filterType =
							(VSI_UI_STUDY_ACCESS_FILTER)VsiGetDiscreteValue<DWORD>(
								VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
								VSI_PARAMETER_UI_STUDY_ACCESS_FILTER, m_pPdm);

						switch (filterType)
						{
						case VSI_UI_STUDY_ACCESS_FILTER_PRIVATE:
							{
								filter.bPublic = FALSE;

								pFilter = &filter;
							}
							break;
						case VSI_UI_STUDY_ACCESS_FILTER_GROUP:
							{
								hr = pOperator->GetGroup(&pszGroup);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								filter.pszGroup = pszGroup;
								filter.bPublic = FALSE;

								pFilter = &filter;
							}
							break;
						case VSI_UI_STUDY_ACCESS_FILTER_ALL:
							{
								VSI_OPERATOR_TYPE operatorType;
								hr = pOperator->GetType(&operatorType);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							
								if (0 == (VSI_OPERATOR_TYPE_ADMIN & operatorType))
								{
									// Not admin, cannot show all, turn on filtering
									hr = pOperator->GetGroup(&pszGroup);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									filter.pszGroup = pszGroup;
									filter.bPublic = TRUE;

									pFilter = &filter;
								}
							}
							break;
						}
					}

					hr = m_pDataList->Clear();
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = m_pStudyMgr->LoadData(V_BSTR(&vPath), VSI_DATA_TYPE_STUDY_LIST, pFilter);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiEnumStudy> pEnum;
					hr = m_pStudyMgr->GetStudyEnumerator(FALSE, &pEnum);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = m_pDataList->Fill(pEnum, 3);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CString strSearch;
					{
						CCritSecLock cs(m_csSearch);
						strSearch = m_strSearch;
						m_strSearch.Empty();
					}

					BOOL bStop = FALSE;			
					m_pDataList->Filter(L"", &bStop);

					if (!strSearch.IsEmpty())
					{
						PostMessage(WM_COMMAND, MAKEWPARAM(IDC_SB_SEARCH, EN_CHANGE), (LPARAM)(HWND)GetDlgItem(IDC_SB_SEARCH));
					}
					else
					{
						PostMessage(WM_VSI_SB_FILL_LIST_DONE, VSI_LIST_RELOAD);
					}
				}
				VSI_CATCH_(err)
				{
					hr = err;
					if (FAILED(hr))
						VSI_ERROR_LOG(err);

					PostMessage(WM_VSI_SB_FILL_LIST_DONE, VSI_LIST_RELOAD);
				}
			}
			else if ((WAIT_OBJECT_0 + VSI_LIST_SEARCH) == dwWait)
			{
				CString strSearch;
				{
					CCritSecLock cs(m_csSearch);
					strSearch = m_strSearch;
				}

				m_bStopSearch = FALSE;
				
				m_pDataList->Filter(strSearch, &m_bStopSearch);

				if (!m_bStopSearch)
				{
					PostMessage(WM_VSI_SB_FILL_LIST_DONE, VSI_LIST_SEARCH);
				}
			}
			else
			{
				VSI_LOG_MSG(VSI_LOG_WARNING, L"WaitForMultipleObjects failed");
				// Keep going
			}
		}
	}
	VSI_CATCH(hr);

	CoUninitialize();

	return 0;
}

void CVsiStudyBrowser::PreviewImages()
{
	HRESULT hr(S_OK);
	BOOL bRet;

	try
	{
		if (NULL == m_pItemPreview)
		{
			StopPreview();
		}
		else
		{
			CVsiLockData lock(m_pDataList);
			if (lock.LockData())
			{
				if (NULL == m_pEnumImage)
				{
					// Gets selection
					m_vIdSelected.Clear();

					int iFocused(-1);
					hr = m_pDataList->GetSelectedItem(&iFocused);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (iFocused >= 0)
					{
						VSI_DATA_LIST_TYPE type;
						CComPtr<IUnknown> pUnkItem;
						hr = m_pDataList->GetItem(iFocused, &type, &pUnkItem);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						
						if (VSI_DATA_LIST_IMAGE == type)
						{
							CComQIPtr<IVsiImage> pImage(pUnkItem);
							VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

							pImage->GetProperty(VSI_PROP_IMAGE_ID, &m_vIdSelected);
						}
					}

					int iCol = VsiGetDiscreteValue<long>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
						VSI_PARAMETER_UI_SORT_IMAGE_ITEM_STUDYBROWSER, m_pPdm);
					int iOrder = VsiGetDiscreteValue<long>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
						VSI_PARAMETER_UI_SORT_IMAGE_ORDER_STUDYBROWSER, m_pPdm);

					CComQIPtr<IVsiSeries> pSeries(m_pItemPreview);
					if (pSeries != NULL)
					{
						hr = pSeries->GetImageEnumerator(
							GetImagePropIdFromColumnId(iCol),
							VSI_SYS_SORT_ORDER_DESCENDING == iOrder,
							&m_pEnumImage, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						LONG lCount;
						hr = pSeries->GetImageCount(&lCount, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						ListView_SetItemCount(m_wndImageList, lCount);
					}
				}

				if (NULL != m_pEnumImage)
				{
					bool bSearch(false);
					{
						CCritSecLock cs(m_csSearch);
						bSearch = !m_strSearch.IsEmpty();
					}

					IVsiImage *ppImage[2];
					ULONG ulCount = 0;
					hr = m_pEnumImage->Next(_countof(ppImage), ppImage, &ulCount);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					for (ULONG i = 0; i < ulCount; ++i)
					{
						if (bSearch)
						{
							CComVariant vNs;
							hr = ppImage[i]->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
							if (FAILED(hr))
							{
								ppImage[i]->Release();
								continue;  // skip
							}

							VSI_DATA_LIST_ITEM_STATUS dwStatus(VSI_DATA_LIST_ITEM_STATUS_NORMAL);
							m_pDataList->GetItemStatus(V_BSTR(&vNs), &dwStatus);

							if (dwStatus & VSI_DATA_LIST_ITEM_STATUS_HIDDEN)
							{
								ppImage[i]->Release();
								continue;  // skip
							}
						}

						CComHeapPtr<OLECHAR> pFile;
						hr = ppImage[i]->GetThumbnailFile(&pFile);
						if (FAILED(hr))
						{
							ppImage[i]->Release();
							continue;  // skip
						}

						CComVariant vName;
						hr = ppImage[i]->GetProperty(VSI_PROP_IMAGE_NAME, &vName);
						if (FAILED(hr))
						{
							ppImage[i]->Release();
							continue;  // skip
						}

						SYSTEMTIME stDate;
						CComVariant vDate;
						hr = ppImage[i]->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vDate);
						if (FAILED(hr))
						{
							ppImage[i]->Release();
							continue;  // skip
						}

						COleDateTime date(vDate);
						date.GetAsSystemTime(stDate);

						// The time is stored as UTC. Convert it to local time for display.
						bRet = SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);
						if (!bRet)
						{
							ppImage[i]->Release();
							continue;  // skip
						}

						WCHAR szText[500], szDate[50], szTime[50];
						GetDateFormat(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &stDate, NULL, szDate, _countof(szDate));
						GetTimeFormat(LOCALE_USER_DEFAULT, 0, &stDate, NULL, szTime, _countof(szTime));

						if (VT_EMPTY != V_VT(&vName) && (lstrlen(V_BSTR(&vName)) != 0))
							swprintf_s(szText, _countof(szText), L"%s\r\n%s %s", V_BSTR(&vName), szTime, szDate);
						else
							swprintf_s(szText, _countof(szText), L"%s %s", szTime, szDate);

						LVITEM lvi = { LVIF_TEXT | LVIF_IMAGE | LVIF_PARAM };
						lvi.pszText = szText;
						lvi.iImage = (int)(INT_PTR)(LPCWSTR)pFile;
						lvi.lParam = (LPARAM)(ppImage[i]);
						ppImage[i]->AddRef();

						// Select?
						if (VT_EMPTY != V_VT(&m_vIdSelected))
						{
							CComVariant vId;
							ppImage[i]->GetProperty(VSI_PROP_IMAGE_ID, &vId);

							if (m_vIdSelected == vId)
							{
								lvi.mask |= LVIF_STATE;
								lvi.state = LVIS_SELECTED;
								lvi.stateMask = LVIS_SELECTED;
							}
						}

						ListView_InsertItem(m_wndImageList, &lvi);

						ppImage[i]->Release();
					}

					if (S_OK != hr)
					{
						StopPreview();
					}
				}
			}
		}
	}
	VSI_CATCH(hr);

	if (FAILED(hr))
	{
		StopPreview();
	}
}

void CVsiStudyBrowser::CancelSearchIfRunning()
{
	// Cancel search if in the middle of a search
	CVsiLockData lock(m_pDataList);
	if (!lock.LockData())
	{
		// Cancel search
		m_wndSearch.SetWindowText(L"");
	}
}

void CVsiStudyBrowser::UpdateSelection() throw(...)
{
	HRESULT hr(S_OK);

	CVsiLockData lock(m_pDataList);
	if (lock.LockData())
	{
		if (!m_strSelection.IsEmpty())
		{
			m_pDataList->SetSelection(m_strSelection, TRUE);
			m_strSelection.Empty();
		}
		else
		{
			// Select item
			int iCount = 0;
			hr = m_pDataList->GetItemCount(&iCount);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (iCount > 0)
			{
				hr = m_pDataList->SetSelectedItem(0, TRUE);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}
	}
}

HRESULT CVsiStudyBrowser::CountSonoVevoImages(int *piImages)
{
	HRESULT hr = S_OK;

	try
	{
		*piImages = 0;

		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			// Get number of images selected
			CComPtr<IVsiEnumImage> pEnumImage;
			int iImages = 0;
			VSI_DATA_LIST_COLLECTION selection;

			memset(&selection, 0, sizeof(VSI_DATA_LIST_COLLECTION));
			selection.dwFlags = VSI_DATA_LIST_FILTER_MOVE_TO_CHILDREN;
			selection.ppEnumImage = &pEnumImage;
			selection.piImageCount = &iImages;

			hr = m_pDataList->GetSelection(&selection);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			int iNumValidImages(0);
			if (iImages > 0)
			{
				pEnumImage->Reset();

				CComPtr<IVsiImage> pImage;
				while (S_OK == pEnumImage->Next(1, &pImage, NULL))
				{
					// Check for null
					if (pImage == NULL)
						continue;
				
					VSI_IMAGE_ERROR dwError;
					hr = pImage->GetErrorCode(&dwError);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (VSI_IMAGE_ERROR_NONE != dwError)
					{
						pImage.Release();
						continue;
					}

					CComQIPtr<IVsiMode> pMode;

					CComVariant vMode;
					hr = pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					CComVariant vSono;
					hr = m_pModeMgr->GetModePropertyFromInternalName(V_BSTR(&vMode), L"sono", &vSono);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (VARIANT_FALSE != V_BOOL(&vSono))
					{
						CComVariant vFrame;				
						hr = pImage->GetProperty(VSI_PROP_IMAGE_FRAME_TYPE, &vFrame);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						CComVariant vLength;				
						hr = pImage->GetProperty(VSI_PROP_IMAGE_LENGTH, &vLength);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						CComVariant vDisplayName;				
						hr = pImage->GetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vDisplayName);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);
					
						BOOL bFrame = VARIANT_TRUE == V_BOOL(&vFrame);
						BOOL b3D = (NULL != wcsstr(V_BSTR(&vDisplayName), L"3D"));

						if (!bFrame && !b3D && 10 <= V_R8(&vLength))
						{
							iNumValidImages++;
						}
					}

					pImage.Release();
				}			
			}	

			*piImages = iNumValidImages;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::UpdateSonoVevoButton()
{
	HRESULT hr = S_OK;

	try
	{		
		int iNumValidImages(0);
		hr = CountSonoVevoImages(&iNumValidImages);		
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
		GetDlgItem(IDC_SB_SONO).EnableWindow(iNumValidImages > 0);
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT	CVsiStudyBrowser::UpdateStudyAccessFilterImage()
{
	HRESULT hr = S_OK;

	try
	{		
		VSI_UI_STUDY_ACCESS_FILTER filterType =
			(VSI_UI_STUDY_ACCESS_FILTER)VsiGetDiscreteValue<DWORD>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_UI_STUDY_ACCESS_FILTER, m_pPdm);

		UINT iImage = IDB_STUDY_ACCESS_PRIVATE;
		UINT iTips = IDS_STYBRW_ACCESS_FILTER_PRIVATE;
		switch (filterType)
		{
		case VSI_UI_STUDY_ACCESS_FILTER_PRIVATE:
			iImage = IDB_STUDY_ACCESS_PRIVATE;
			iTips = IDS_STYBRW_ACCESS_FILTER_PRIVATE;
			break;
		case VSI_UI_STUDY_ACCESS_FILTER_GROUP:
			iImage = IDB_STUDY_ACCESS_GROUP;
			iTips = IDS_STYBRW_ACCESS_FILTER_GROUP;
			break;
		case VSI_UI_STUDY_ACCESS_FILTER_ALL:
			iImage = IDB_STUDY_ACCESS_ALL;
			iTips = IDS_STYBRW_ACCESS_FILTER_ALL;
			break;
		default:
			_ASSERT(0);
		}

		m_ilStudyAccess.Destroy();
		m_ilStudyAccess.CreateFromImage(
			_U_STRINGorID(iImage),
			29,
			1,
			CLR_DEFAULT,
			IMAGE_BITMAP,
			LR_CREATEDIBSECTION);

		BUTTON_IMAGELIST bil = { m_ilStudyAccess, { 0, 0, 0, 0 }, BUTTON_IMAGELIST_ALIGN_CENTER };

		CWindow wndAccess(GetDlgItem(IDC_SB_ACCESS_FILTER));

		wndAccess.SendMessage(BCM_SETIMAGELIST, 0 , reinterpret_cast<LPARAM>(&bil));
		wndAccess.SendMessage(
			RegisterWindowMessage(VSI_THEME_COMMAND),
			VSI_BT_CMD_SETTIP,
			(LPARAM)CString(MAKEINTRESOURCE(iTips)).GetString());
	}
	VSI_CATCH(hr);

	return hr;
}

void CVsiStudyBrowser::UpdateUI()
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

		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
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

			lock.UnlockData();
		}

		int iTotal = iStudy + iSeries + iImage;

		GetDlgItem(IDC_SB_DELETE).EnableWindow(iTotal > 0);
		GetDlgItem(IDC_SB_COPY_TO).EnableWindow(iTotal > 0);
		GetDlgItem(IDC_SB_EXPORT).EnableWindow(iTotal > 0);

		UpdateNewCloseButtons();

		// Gets focused item
		int iFocused(-1);
		VSI_DATA_LIST_TYPE type = VSI_DATA_LIST_NONE;
		CComPtr<IUnknown> pUnkItemFocused;
		if (lock.LockData())
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
				else if (VSI_DATA_LIST_IMAGE == type)
				{
					CComQIPtr<IVsiImage> pImage(pUnkItemFocused);
					VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

					VSI_IMAGE_ERROR dwError(VSI_IMAGE_ERROR_NONE);
					pImage->GetErrorCode(&dwError);

					if (VSI_IMAGE_ERROR_NONE != dwError)
					{
						StopPreview(VSI_PREVIEW_LIST_CLEAR);
						m_pItemPreview = NULL;
						SetStudyInfo(NULL, false);
						iFocused = -1;
					}
				}
			}

			GetDlgItem(IDC_SB_ANALYSIS).EnableWindow(0 < iTotal);

			if ((iFocused >= 0) && (VSI_DATA_LIST_IMAGE == type))
			{
				CComQIPtr<IVsiImage> pImage(pUnkItemFocused);
				VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

				CComVariant vMode;				
				hr = pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				CComVariant vFrame;				
				hr = pImage->GetProperty(VSI_PROP_IMAGE_FRAME_TYPE, &vFrame);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				CComVariant vLength;				
				hr = pImage->GetProperty(VSI_PROP_IMAGE_LENGTH, &vLength);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				CComVariant vDisplayName;				
				hr = pImage->GetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vDisplayName);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				// Display strain button
				CComVariant vStrain;
				hr = m_pModeMgr->GetModePropertyFromInternalName(
					V_BSTR(&vMode), L"strain", &vStrain);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				GetDlgItem(IDC_SB_STRAIN).EnableWindow(
					(VARIANT_FALSE != V_BOOL(&vStrain)) &&				// Mode support strain
					(VARIANT_FALSE == V_BOOL(&vFrame)) &&				// Not frame
					(NULL == wcsstr(V_BSTR(&vDisplayName), L"3D") &&	// Not 3D (e.g. RF B-Mode 3D)
					10.0 <= V_R8(&vLength))		
					);

				// Display sono button
				hr = UpdateSonoVevoButton();				
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else
			{
				GetDlgItem(IDC_SB_STRAIN).EnableWindow(FALSE);
				GetDlgItem(IDC_SB_SONO).EnableWindow(FALSE);
			}
			
			// Enable info button only when the focus is either
			// on the study or series or image
			GetDlgItem(IDC_SB_STUDY_SERIES_INFO).EnableWindow(iFocused >= 0);
			if (1 == iTotal)
			{
				// If total > 1, it is a batch export, the state was updated above
				GetDlgItem(IDC_SB_EXPORT).EnableWindow(iFocused >= 0);
			}

			GetDlgItem(IDC_SB_ANALYSIS).EnableWindow(iFocused >= 0);
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
					VSI_CHECK_INTERFACE(pPersist, VSI_LOG_ERROR, NULL);

					hr = pPersist->LoadSeriesData(NULL, VSI_DATA_TYPE_IMAGE_LIST, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					StartPreview();
				}
			}
			else if (VSI_DATA_LIST_IMAGE == type)
			{
				CComQIPtr<IVsiImage> pImage(pUnkItemFocused);
				VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

				CComPtr<IVsiSeries> pSeries;
				hr = pImage->GetSeries(&pSeries);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

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
				else
				{
					// Selects image in preview
					CComVariant vIdSel;
					pImage->GetProperty(VSI_PROP_IMAGE_ID, &vIdSel);

					LV_ITEM item = { LVIF_PARAM };

					int iCount = ListView_GetItemCount(m_wndImageList);
					for (item.iItem = 0; item.iItem < iCount; ++item.iItem)
					{
						BOOL bRet = ListView_GetItem(m_wndImageList, &item);
						if (!bRet) continue;

						IVsiImage *pImageTest = (IVsiImage*)item.lParam;
						CComVariant vId;
						pImageTest->GetProperty(VSI_PROP_IMAGE_ID, &vId);
						if (vId == vIdSel)
						{
							item.mask = LVIF_STATE;
							item.stateMask = LVIS_SELECTED;
							item.state = LVIS_SELECTED;
							ListView_SetItem(m_wndImageList, &item);
							ListView_EnsureVisible(m_wndImageList, item.iItem, FALSE);

							CComQIPtr<IVsiSeries> pSeriesPreview(m_pItemPreview);
							if (pSeriesPreview != NULL)
							{
								LONG lCount(0);
								hr = pSeriesPreview->GetImageCount(&lCount, NULL);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								
								CString strText;
								strText.Format(IDS_STYBRW_IMAGE_COUNT_STATUS, lCount - item.iItem, lCount);

								m_pAppView->SetStatusText(strText, 2000);
							}

							break;
						}
					}
				}
			}
		}
	}
	VSI_CATCH(hr);
}

void CVsiStudyBrowser::UpdateLogoutState() throw(...)
{
	BOOL bApp = (m_dwAppMode & VSI_APP_MODE_WORKSTATION) == 0;

	CWindow wndHelp(GetDlgItem(IDC_SB_HELP));
	CWindow wndPrefs(GetDlgItem(IDC_SB_PREFS));
	CWindow wndShutdown(GetDlgItem(IDC_SB_SHUTDOWN));
	CWindow wndLogout(GetDlgItem(IDC_SB_LOGOUT));

	CRect rcHelp;
	CRect rcPrefs;
	CRect rcShutdown;
	CRect rcLogout;

	wndHelp.GetWindowRect(&rcHelp);
	ScreenToClient(&rcHelp);

	wndPrefs.GetWindowRect(&rcPrefs);
	ScreenToClient(&rcPrefs);

	wndShutdown.GetWindowRect(&rcShutdown);
	ScreenToClient(&rcShutdown);

	wndLogout.GetWindowRect(&rcLogout);
	ScreenToClient(&rcLogout);

	int iShift;
	if (bApp)
	{
		iShift = rcLogout.Width() - rcShutdown.Width();
	}
	else
	{
		iShift = rcLogout.left - rcPrefs.right + rcLogout.Width();
	}

	BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
		VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
		VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);

	if (bSecureMode)
	{
		wndLogout.EnableWindow(TRUE);

		int iShift = rcLogout.Width() - rcShutdown.Width();

		if (0 == (WS_VISIBLE & wndLogout.GetStyle()))
		{
			m_pLayout->AdjustControl(IDC_SB_HELP, -iShift, VSI_WL_ADJUST_X);
			m_pLayout->AdjustControl(IDC_SB_PREFS, -iShift, VSI_WL_ADJUST_X);
			m_pLayout->AdjustControl(IDC_SB_SHUTDOWN, -iShift, VSI_WL_ADJUST_X);

			wndLogout.ShowWindow(SW_SHOWNA);
			wndShutdown.ShowWindow(SW_HIDE);
			wndShutdown.EnableWindow(FALSE);

			m_pLayout->Refresh();
		}

		wndShutdown.ShowWindow(SW_HIDE);
		wndShutdown.EnableWindow(FALSE);
	}
	else
	{
		wndLogout.EnableWindow(FALSE);

		if (0 != (WS_VISIBLE & wndLogout.GetStyle()))
		{
			m_pLayout->AdjustControl(IDC_SB_HELP, iShift, VSI_WL_ADJUST_X);
			m_pLayout->AdjustControl(IDC_SB_PREFS, iShift, VSI_WL_ADJUST_X);
			m_pLayout->AdjustControl(IDC_SB_SHUTDOWN, iShift, VSI_WL_ADJUST_X);

			wndLogout.ShowWindow(SW_HIDE);
			m_pLayout->Refresh();
		}

		if (bApp)
		{
			wndShutdown.ShowWindow(SW_SHOW);
			wndShutdown.EnableWindow(TRUE);
		}
		else
		{
			wndShutdown.ShowWindow(SW_HIDE);
			wndShutdown.EnableWindow(FALSE);
		}
	}

	CVsiWindow wndAccessFilter(GetDlgItem(IDC_SB_ACCESS_FILTER));
	wndAccessFilter.EnableWindow(bSecureMode);
	wndAccessFilter.ShowWindow(bSecureMode ? SW_SHOWNA : SW_HIDE);
}

void CVsiStudyBrowser::OpenItem()
{
	HRESULT hr = S_OK;

	try
	{
		if (GetFocus() == m_wndImageList)
		{
			int iSel = ListView_GetNextItem(m_wndImageList, -1, LVNI_SELECTED);
			if (iSel >= 0)
			{
				StopPreview();

				LVITEMW lvi = { LVIF_PARAM };
				lvi.iItem = iSel;

				ListView_GetItem(m_wndImageList, &lvi);
				if (lvi.lParam == NULL)
					return;

				Open((IVsiImage*)lvi.lParam);

				return;
			}
		}

		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			// Make sure the selection record is up-to-date
			int iSel(0);
			hr = m_pDataList->GetSelectedItem(&iSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
			if (S_OK == hr)
			{
				StopPreview();

				VSI_DATA_LIST_TYPE type;
				CComPtr<IUnknown> pUnkItem;

				hr = m_pDataList->GetItem(iSel, &type, &pUnkItem);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
				switch (type)
				{
				case VSI_DATA_LIST_STUDY:
					{
						CComQIPtr<IVsiStudy> pStudy(pUnkItem);
						VSI_CHECK_INTERFACE(pStudy, VSI_LOG_ERROR, NULL);
						Open(pStudy);
					}
					break;
				case VSI_DATA_LIST_SERIES:
					{
						CComQIPtr<IVsiSeries> pSeries(pUnkItem);
						VSI_CHECK_INTERFACE(pSeries, VSI_LOG_ERROR, NULL);
						Open(pSeries);
					}
					break;
				case VSI_DATA_LIST_IMAGE:
					{
						CComQIPtr<IVsiImage> pImage(pUnkItem);
						VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);
						Open(pImage);
					}
					break;
				default:
					VSI_LOG_MSG(VSI_LOG_ERROR, NULL);
				}
			}
		}
	}
	VSI_CATCH(hr);
}

void CVsiStudyBrowser::Open(IVsiStudy *pStudy)
{
	VSI_STUDY_ERROR dwError;
	pStudy->GetErrorCode(&dwError);

	if (VSI_STUDY_ERROR_NONE == dwError)
	{
		CComVariant vParam(pStudy);
		m_pCmdMgr->InvokeCommand(ID_CMD_STUDY_SERIES_INFO, &vParam);
	}
	else if (VSI_STUDY_ERROR_VERSION_INCOMPATIBLE & dwError)
	{
		WCHAR szSupport[500];
		VsiGetTechSupportInfo(szSupport, _countof(szSupport));

		CString strMsg(MAKEINTRESOURCE(IDS_STYBRW_STUDY_NOT_SUPPORTED));
		strMsg += szSupport;

		VsiMessageBox(
			GetTopLevelParent(),
			strMsg,
			CString(MAKEINTRESOURCE(IDS_STYBRW_STUDY_VERIFICATION)),
			MB_OK | MB_ICONEXCLAMATION);
	}
	else
	{
		WCHAR szSupport[500];
		VsiGetTechSupportInfo(szSupport, _countof(szSupport));

		CString strMsg(MAKEINTRESOURCE(IDS_STUDY_OPEN_ERROR));
		strMsg += szSupport;

		VsiMessageBox(
			GetTopLevelParent(),
			strMsg,
			VsiGetRangeValue<CString>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_SYS_PRODUCT_NAME,
				m_pPdm),
			MB_OK | MB_ICONEXCLAMATION);
	}
}

void CVsiStudyBrowser::Open(IVsiSeries *pSeries)
{
	VSI_SERIES_ERROR dwError;
	pSeries->GetErrorCode(&dwError);

	if (VSI_SERIES_ERROR_NONE == dwError)
	{
		CComVariant vParam(pSeries);
		m_pCmdMgr->InvokeCommand(ID_CMD_STUDY_SERIES_INFO, &vParam);
	}
	else if (VSI_SERIES_ERROR_VERSION_INCOMPATIBLE & dwError)
	{
		WCHAR szSupport[500];
		VsiGetTechSupportInfo(szSupport, _countof(szSupport));

		CString strMsg(MAKEINTRESOURCE(IDS_STYBRW_SERIES_NOT_SUPPORTED));			
		strMsg += szSupport;

		VsiMessageBox(
			GetTopLevelParent(),
			strMsg,
			CString(MAKEINTRESOURCE(IDS_STYBRW_SERIES_DATA_VERIFICATION)),
			MB_OK | MB_ICONEXCLAMATION);
	}
	else
	{
		WCHAR szSupport[500];
		VsiGetTechSupportInfo(szSupport, _countof(szSupport));

		CString strMsg(MAKEINTRESOURCE(IDS_STYBRW_ERROR_WHILE_OPEN_SERIES));
		strMsg += szSupport;

		VsiMessageBox(
			GetTopLevelParent(),
			strMsg,
			VsiGetRangeValue<CString>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_SYS_PRODUCT_NAME,
				m_pPdm),
			MB_OK | MB_ICONEXCLAMATION);
	}
}

void CVsiStudyBrowser::Open(IVsiImage *pImage)
{
	CWaitCursor wait;

	CComVariant v(pImage);
	m_pCmdMgr->InvokeCommand(ID_CMD_OPEN_IMAGE, &v);
}

HRESULT CVsiStudyBrowser::Delete(
	IVsiStudy *pStudy,
	IXMLDOMDocument *pXmlDoc,
	IXMLDOMElement *pElmRemove)
{
	HRESULT hr(S_OK);

	try
	{
		// 1st - Hide the study from the system by rename to .del extension
		// 2nd - Delete the .del folder

		CComHeapPtr<OLECHAR> pszPath;
		hr = pStudy->GetDataPath(&pszPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		PathRemoveFileSpec(pszPath);

		// Rename folder
		WCHAR szDest[MAX_PATH];
		swprintf_s(szDest, _countof(szDest), L"%s.del", pszPath.operator OLECHAR *());

		BOOL bRet = MoveFile(pszPath, szDest);
		if (bRet)  // OK
		{
			bRet = VsiRemoveDirectory(szDest);
			if (bRet)
			{
				VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Deleted study: %s", szDest));
			}
			else
			{
				VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"Failure deleting study: %s", szDest));
				return S_FALSE;  // Reports errors in caller
			}
		}
		else
		{
			VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"Failure moving study: %s to %s", pszPath, szDest));
			return S_FALSE;  // Reports errors in caller
		}

		CComVariant vId;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		CComPtr<IXMLDOMElement> pElmItem;
		hr = pXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_ITEM), &pElmItem);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"Failure creating '%s' element", VSI_UPDATE_XML_ELM_ITEM));

		hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), vId);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"Failed to set attribute (%s)", VSI_UPDATE_XML_ATT_NS));

		hr = pElmRemove->appendChild(pElmItem, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);
	
	return hr;
}

HRESULT CVsiStudyBrowser::Delete(
	IVsiSeries *pSeries,
	IXMLDOMDocument *pXmlDoc,
	IXMLDOMElement *pElmRemove)
{
	HRESULT hr(S_OK);

	try
	{
		// 1st - Hide the series from the system by rename to .del extension
		// 2nd - Delete the .del folder

		CComHeapPtr<OLECHAR> pszPath;
		hr = pSeries->GetDataPath(&pszPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		PathRemoveFileSpec(pszPath);

		// Rename folder
		WCHAR szDest[MAX_PATH];
		swprintf_s(szDest, _countof(szDest), L"%s.del", pszPath.operator OLECHAR *());

		BOOL bRet = MoveFile(pszPath, szDest);
		if (bRet)  // OK
		{
			bRet = VsiRemoveDirectory(szDest);
			if (bRet)
			{
				VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Deleted series: %s", szDest));
			}
			else
			{
				VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"Failure deleting series: %s", szDest));
				return S_FALSE;  // Reports errors in caller
			}
		}
		else
		{
			VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"Failure moving series: %s to %s", pszPath, szDest));
			return S_FALSE;  // Reports errors in caller
		}

		CComPtr<IXMLDOMElement> pElmItem;
		hr = pXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_ITEM), &pElmItem);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vNs;
		hr = pSeries->GetProperty(VSI_PROP_SERIES_NS, &vNs);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), vNs);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pElmRemove->appendChild(pElmItem, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);
	
	return hr;
}

HRESULT CVsiStudyBrowser::Delete(
	IVsiImage *pImage,
	IXMLDOMDocument *pXmlDoc,
	IXMLDOMElement *pElmRemove)
{
	HRESULT hr(S_OK);

	try
	{
		CComHeapPtr<OLECHAR> pszFile;
		hr = pImage->GetDataPath(&pszFile);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		LPCWSTR pszSpt = wcsrchr(pszFile, L'.');

		WCHAR szPath[MAX_PATH];
		lstrcpyn(szPath, pszFile, (int)(pszSpt - pszFile) + 1);

		HANDLE hFile = CreateFile(
			pszFile,
			DELETE,
			0,
			NULL,
			OPEN_EXISTING,
			FILE_ATTRIBUTE_NORMAL,
			NULL);

		if (INVALID_HANDLE_VALUE == hFile)
		{
			VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"Failure deleting image: %s", pszFile));
			return HRESULT_FROM_WIN32(GetLastError());  // Reports errors in caller
		}

		CloseHandle(hFile);

		// Delete xml file first
		BOOL bRet = VsiRemoveFile(pszFile, FALSE);
		if (bRet)
		{
			VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Deleted image: %s", pszFile));
		}
		else
		{
			VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"Failure deleting image: %s", pszFile));
			return S_FALSE;  // Reports errors in caller
		}

		// Other files to delete
		WCHAR szDest[MAX_PATH];
		lstrcpy(szDest, szPath);
		lstrcat(szDest, L".*");

		bRet = VsiRemoveFile(szDest, FALSE);
		if (bRet)
		{
			VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Deleted image: %s", szDest));
		}
		else
		{
			VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"Failure deleting image: %s", szDest));
			// The xml file is gone already - the rest are not critical, just wasted spaces
			// Keep going
		}

		CComPtr<IXMLDOMElement> pElmItem;
		hr = pXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_ITEM), &pElmItem);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"Failure creating '%s' element", VSI_UPDATE_XML_ELM_ITEM));

		CComVariant vNs;
		hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), vNs);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"Failed to set attribute (%s)", VSI_UPDATE_XML_ATT_NS));

		hr = pElmRemove->appendChild(pElmItem, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);
	
	return hr;
}

void CVsiStudyBrowser::SetStudyInfo(IVsiStudy *pStudy, bool bUpdateUI)
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

VSI_PROP_IMAGE CVsiStudyBrowser::GetImagePropIdFromColumnId(int iCol)
{
	VSI_PROP_IMAGE prop;

	switch (iCol)
	{
	case VSI_SYS_SORT_ITEM_SB_NAME:
		prop = VSI_PROP_IMAGE_NAME;
		break;
	case VSI_SYS_SORT_ITEM_SB_LOCKED:
		prop = (VSI_PROP_IMAGE)VSI_PROP_STUDY_LOCKED;
		break;
	default:
	case VSI_SYS_SORT_ITEM_SB_DATE:
		prop = VSI_PROP_IMAGE_ACQUIRED_DATE;
		break;
	case VSI_SYS_SORT_ITEM_SB_OWNER:
		prop = (VSI_PROP_IMAGE)VSI_PROP_STUDY_OWNER;
		break;
	case VSI_SYS_SORT_ITEM_SB_ACQUIRE_BY:
		prop = (VSI_PROP_IMAGE)VSI_PROP_SERIES_ACQUIRED_BY;
		break;
	case VSI_SYS_SORT_ITEM_SB_MODE:
		prop = VSI_PROP_IMAGE_MODE_DISPLAY;
		break;
	case VSI_SYS_SORT_ITEM_SB_LENGTH:
		prop = VSI_PROP_IMAGE_LENGTH;
		break;
	}

	return prop;
}

VSI_PROP_SERIES CVsiStudyBrowser::GetSeriesPropIdFromColumnId(int iCol)
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

BOOL CVsiStudyBrowser::CheckForOpenItems(
	const VSI_DATA_LIST_COLLECTION &listSel,
	int *piNumStudy,
	int *piNumSeries,
	int *piNumImage)
{
	BOOL bRet = FALSE;
	HRESULT hr = S_OK;
	CComVariant vNs;
	int iSlot(-1);

	try
	{
		CComQIPtr<IVsiEnumStudy> pEnumStudy(*(listSel.ppEnumStudy));
		CComQIPtr<IVsiEnumSeries> pEnumSeries(*(listSel.ppEnumSeries));
		CComQIPtr<IVsiEnumImage> pEnumImage(*(listSel.ppEnumImage));

		if (pEnumImage != NULL)
		{
			CComPtr<IVsiImage> pImage;
			while (S_OK == pEnumImage->Next(1, &pImage, NULL))
			{
				vNs.Clear();
				hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				HRESULT hrLoaded = m_pSession->GetIsItemLoaded(V_BSTR(&vNs), &iSlot);
				if (S_OK == hrLoaded)
					++(*piNumImage);

				pImage.Release();
			}
		}

		if (pEnumSeries != NULL)
		{
			CComPtr<IVsiSeries> pSeries;
			while (S_OK == pEnumSeries->Next(1, &pSeries, NULL))
			{
				vNs.Clear();
				hr = pSeries->GetProperty(VSI_PROP_SERIES_NS, &vNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				HRESULT hrLoaded = m_pSession->GetIsItemLoaded(V_BSTR(&vNs), &iSlot);
				if (S_OK == hrLoaded)
					++(*piNumSeries);

				pSeries.Release();
			}
		}

		if (pEnumStudy != NULL)
		{
			CComPtr<IVsiStudy> pStudy;
			while (S_OK == pEnumStudy->Next(1, &pStudy, NULL))
			{
				vNs.Clear();
				hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				HRESULT hrLoaded = m_pSession->GetIsItemLoaded(V_BSTR(&vNs), &iSlot);
				if (S_OK == hrLoaded)
					++(*piNumStudy);

				pStudy.Release();
			}
		}

		if (pEnumImage != NULL)
			pEnumImage->Reset();
		if (pEnumSeries != NULL)
			pEnumSeries->Reset();
		if (pEnumStudy != NULL)
			pEnumStudy->Reset();
	}
	VSI_CATCH(hr);

	if (S_OK == hr)
	{
		if (*piNumStudy > 0 || *piNumSeries > 0 || *piNumImage > 0)
			bRet = TRUE;
	}

	return bRet;
}

void CVsiStudyBrowser::UpdateNewCloseButtons()
{
	HRESULT hr(S_OK);

	try
	{
		DWORD dwState = VsiGetBitfieldValue<DWORD>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_STATE, m_pPdm);
		if (0 < (VSI_SYS_STATE_HARDWARE & dwState))
		{
			CWindow wndNew(GetDlgItem(IDC_SB_NEW));
			CWindow wndClose(GetDlgItem(IDC_SB_CLOSE));

			if (0 < (VSI_SYS_STATE_ACQ_SESSION & dwState))
			{
				CVsiToolsMore::AddIgnore(wndNew);
				CVsiToolsMore::RemoveIgnore(wndClose);

				if (wndNew.IsWindowEnabled())
				{
					wndNew.EnableWindow(FALSE);
					wndNew.ShowWindow(SW_HIDE);
				}
				if (!wndClose.IsWindowEnabled())
				{
					wndClose.EnableWindow(TRUE);
					wndClose.ShowWindow(SW_SHOWNA);
				}
			}
			else
			{
				CVsiToolsMore::AddIgnore(wndClose);
				CVsiToolsMore::RemoveIgnore(wndNew);

				if (!wndNew.IsWindowEnabled())
				{
					wndNew.EnableWindow(TRUE);
					wndNew.ShowWindow(SW_SHOWNA);
				}
				if (wndClose.IsWindowEnabled())
				{
					wndClose.EnableWindow(FALSE);
					wndClose.ShowWindow(SW_HIDE);
				}
			}

			CVsiToolsMore::Update();
		}
	}
	VSI_CATCH(hr);
}

void CVsiStudyBrowser::UpdateDataList(IXMLDOMDocument *pXmlDoc)
{
	HRESULT hr = S_OK;

	try
	{
#ifdef _DEBUG
{
		WCHAR szPath[MAX_PATH];
		GetModuleFileName(NULL, szPath, _countof(szPath));
		LPWSTR pszSpt = wcsrchr(szPath, L'\\');
		lstrcpy(pszSpt + 1, VSI_FOLDER_DEBUG);
		lstrcat(szPath, L"\\update.xml");
		pXmlDoc->save(CComVariant(szPath));
}
#endif

		hr = m_pStudyMgr->Update(pXmlDoc);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Fire update event
		CComQIPtr<IVsiPdmEvents> pPdmEvents(m_pPdm);
		VSI_CHECK_INTERFACE(pPdmEvents, VSI_LOG_ERROR, NULL);
		
		hr = pPdmEvents->FireUpdate();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	UpdateUI();
}

void CVsiStudyBrowser::SetPreviewSize(int iSize)
{
	m_iPreview = iSize;
		
	VsiSetRangeValue<int>(
		m_iPreview,
		VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
		VSI_PARAMETER_UI_STUDY_BROWSER_PREVIEW_WIDTH, m_pPdm);

	CRect rc;
	GetClientRect(&rc);
	SendMessage(WM_SIZE, 0, MAKELPARAM(rc.Width(), rc.Height()));
}

HRESULT CVsiStudyBrowser::GenerateFullDisclosureMsmntReport(BOOL bActivateReport)
{
	// Gather up the selected studies and series. Pass them on to the report
	// generator in the form of an xml document.
	HRESULT hr = S_OK;
	CComPtr<IXMLDOMDocument> pXmlDoc;

	try
	{
		int iImage = 0;
		// Create XML DOM document with job list		
		// Get selection
		CComPtr<IVsiEnumImage> pEnumImage;
		VSI_DATA_LIST_COLLECTION listSel;

		CVsiLockData lock(m_pDataList);
		if (lock.LockData())
		{
			memset(&listSel, 0, sizeof(listSel));
			listSel.dwFlags = VSI_DATA_LIST_FILTER_SELECT_ALL_IMAGES;
			listSel.ppEnumImage = &pEnumImage;
			listSel.piImageCount = &iImage;
			hr = m_pDataList->GetSelection(&listSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			lock.UnlockData();
		}
		
		if (0 == iImage)
		{
			VsiMessageBox(*this, L"To generate a report, you must select at least one image.", L"Report", MB_OK | MB_ICONWARNING);

			return 0;
		}
	
		WCHAR szPrevId[128];

		szPrevId[0] = NULL;

		// release AnalysisSet object
		if (NULL != m_pAnalysisSet)
		{
			m_pAnalysisSet->Uninitialize();
			m_pAnalysisSet.Release();
		}

		// Creates XML doc
		hr = VsiCreateDOMDocument(&pXmlDoc);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating DOMDocument");

		CComPtr<IXMLDOMProcessingInstruction> pPI;
		hr = pXmlDoc->createProcessingInstruction(CComBSTR(L"xml"), CComBSTR(L"version='1.0'"), &pPI);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pXmlDoc->appendChild(pPI, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Create and attach the root element
		CComPtr<IXMLDOMElement> pAnalysisResults;

		hr = pXmlDoc->createElement(CComBSTR(VSI_MSMNT_XML_ELE_ANAL_RSLTS), &pAnalysisResults);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pAnalysisResults->setAttribute(
			CComBSTR(VSI_MSMNT_XML_ATT_VERSION), CComVariant(VSI_MSMNT_FILE_VERSION));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failed to write version");
		hr = pXmlDoc->appendChild(pAnalysisResults, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failed to append root element to document");

		hr = m_pAnalysisSet.CoCreateInstance(__uuidof(VsiAnalysisSet));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failed to create Analysis Set");

		// Initialize it with an output xml report file
		CComVariant vXmlDoc(pXmlDoc);
		hr = m_pAnalysisSet->Initialize(m_pApp, &vXmlDoc);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failed to initialize Analysis Set");

		// Loops selected images
		CComPtr<IVsiImage> pImage;
		if (pEnumImage != NULL)
		{
			while (pEnumImage->Next(1, &pImage, NULL) == S_OK)
			{
				VSI_IMAGE_ERROR dwError;
				hr = pImage->GetErrorCode(&dwError);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (VSI_STUDY_ERROR_LOAD_XML == dwError)
				{
					WCHAR szSupport[500];
					VsiGetTechSupportInfo(szSupport, _countof(szSupport));

					CString strMsg(MAKEINTRESOURCE(IDS_STYBRW_OPEN_IMAGE_ERROR));
					strMsg += szSupport;

					VsiMessageBox(
						GetTopLevelParent(),
						strMsg,
						VsiGetRangeValue<CString>(
							VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
							VSI_PARAMETER_SYS_PRODUCT_NAME,
							m_pPdm),
						MB_OK | MB_ICONEXCLAMATION);

					return 0;
				}

				// Copy image to analysis report
				hr = m_pAnalysisSet->CopyImageFromAnalysisMsmntInfo(pImage);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure copying image to the analysis file");

				pImage.Release();
			}
		}

#ifdef _DEBUG
{
		WCHAR szPath[MAX_PATH];
		GetModuleFileName(NULL, szPath, _countof(szPath));
		LPWSTR pszSpt = wcsrchr(szPath, L'\\');
		lstrcpy(pszSpt + 1, VSI_FOLDER_DEBUG);
		lstrcat(szPath, L"\\AnalysisReportB4C.xml");
		pXmlDoc->save(CComVariant(szPath));
}
#endif

		// Process calculations for series with individual images selected in the Study Browser
		// Those series are the series with unprocessed calculations
		if (iImage > 0)
		{
			// FALSE parameter means only marked series are going to be processed
			hr = m_pAnalysisSet->ProcessCalculations(FALSE, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure to process outstanding calculations");

#ifdef _DEBUG
{
		WCHAR szPath[MAX_PATH];
		GetModuleFileName(NULL, szPath, _countof(szPath));
		LPWSTR pszSpt = wcsrchr(szPath, L'\\');
		lstrcpy(pszSpt + 1, VSI_FOLDER_DEBUG);
		lstrcat(szPath, L"\\AnalysisReportCal.xml");
		pXmlDoc->save(CComVariant(szPath));
}
#endif
		}

		if (bActivateReport)
		{
			CWaitCursor wait;

			// Passes IXMLDOMDocument
			CComVariant v(pXmlDoc);
			m_pCmdMgr->InvokeCommand(ID_CMD_OPEN_ANALYSIS_BROWSER, &v);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

// Check for existence of data folder, and create if it doesn't exist
// then signal the FillStudyList thread
HRESULT CVsiStudyBrowser::FillStudyList()
{
	HRESULT hr = S_OK;

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

		CComVariant vPath;
		hr = pRange->GetValue(&vPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		BOOL bDataFolderExists = TRUE;
		if (INVALID_FILE_ATTRIBUTES == GetFileAttributes(V_BSTR(&vPath)))
		{
			bDataFolderExists = VsiCreateDirectory(V_BSTR(&vPath), TRUE);
		}

		if (bDataFolderExists)
		{
			DisableSelectionUI();

			m_bStopSearch = TRUE;

			SetEvent(m_hReload);
		}
		else 
		{
			// Can't create data folder, pop message box
			CString strMsg;

			DWORD dwState = VsiGetBitfieldValue<DWORD>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_SYS_STATE, m_pPdm);
			if (0 < (VSI_SYS_STATE_HARDWARE & dwState))
			{
				// Hardware
				strMsg = L"An error occurred while loading the study data.\n";
			}
			else
			{
				// Workstation
				strMsg = L"An error occurred while loading the study data from:\n\n";
				strMsg += V_BSTR(&vPath);
				strMsg += L"\n\nMake sure you have the correct studies location permissions.\n";
			}

			WCHAR szSupport[500];
			VsiGetTechSupportInfo(szSupport, _countof(szSupport));

			strMsg += szSupport;

			VsiMessageBox(
				GetTopLevelParent(),
				strMsg,
				VsiGetRangeValue<CString>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_SYS_PRODUCT_NAME,
					m_pPdm),
				MB_OK | MB_ICONEXCLAMATION);

			SendMessage(WM_VSI_SB_FILL_LIST_DONE, VSI_LIST_RELOAD);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::UpdateColumnsRecord()
{
	HRESULT hr = S_OK;

	try
	{
		CWindow wndStudyList(GetDlgItem(IDC_SB_STUDY_LIST));

		VSI_LVCOLUMN *plvcolumns(NULL);
		int iColumns(0);

		plvcolumns = g_plvcStudyBrowserListView;
		iColumns =  _countof(g_plvcStudyBrowserListView);

		// Get latest settings
		for (int i = 0; i < iColumns; ++i)
		{
			plvcolumns[i].bVisible = FALSE;
			plvcolumns[i].iOrder = -1;
		}

		CWindow wndHeader(ListView_GetHeader(wndStudyList));
		int iListColumns = Header_GetItemCount(wndHeader);

		LVCOLUMN lvc = { LVCF_SUBITEM | LVCF_WIDTH | LVCF_ORDER };
		for (int i = 0; i < iListColumns; ++i)
		{
			ListView_GetColumn(wndStudyList, i, &lvc);

			for (int j = 0; j < iColumns; ++j)
			{
				if (lvc.iSubItem == plvcolumns[j].iSubItem)
				{
					plvcolumns[j].cx = lvc.cx;
					plvcolumns[j].bVisible = TRUE;
					plvcolumns[j].iOrder = lvc.iOrder;
					break;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// Sierra III removed several calculations (e.g LV Vol;d), this function add them back
/// </summary>
HRESULT	CVsiStudyBrowser::FixSaxCalc(IVsiEnumImage *pEnumImages)
{
	HRESULT hr = S_OK;

	try
	{
		CComPtr<IVsiImage> pImage;
		while (S_OK == pEnumImages->Next(1, &pImage, NULL))
		{
			WCHAR szMsmntFile[MAX_PATH];
			*szMsmntFile = 0;

			try
			{
				CComHeapPtr<OLECHAR> pszDataPath;
				hr = pImage->GetDataPath(&pszDataPath);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				wcscpy_s(szMsmntFile, pszDataPath);

				BOOL bRet = PathRenameExtension(szMsmntFile, L"." VSI_FILE_EXT_MEASUREMENT);
				VSI_VERIFY(FALSE != bRet, VSI_LOG_ERROR, NULL);

				CComPtr<IVsiMsmntFile> pMsmntFile;
				hr = pMsmntFile.CoCreateInstance(__uuidof(VsiMsmntFile));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				VSI_MSMNT_FILE_PATCH patch(VSI_MSMNT_FILE_PATCH_SAX_CALC);
				hr = pMsmntFile->Initialize(szMsmntFile, &patch);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (VSI_MSMNT_FILE_PATCH_RESULT_PATCHED == patch)
				{
					hr = pMsmntFile->Resave();
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				hr = pMsmntFile->Uninitialize();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			VSI_CATCH_(err)
			{
				if (0 != *szMsmntFile)
				{
					VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"Fixed Calculation Failed: %s", szMsmntFile));
				}
				else
				{
					VSI_LOG_MSG(VSI_LOG_WARNING, L"Fixed Calculation Failed: Path unknown");
				}
			}

			pImage.Release();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiStudyBrowser::DeleteEkvRaw(IVsiEnumImage *pEnumImageAll)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"DeleteEkvRaw");
	HRESULT hr = S_OK;

	try
	{
		// Process all images
		if (pEnumImageAll != NULL)
		{
			CComPtr<IVsiImage> pImage;
			while (S_OK == pEnumImageAll->Next(1, &pImage, NULL))
			{
				CComPtr<IVsiStudy> pStudy;
				hr = pImage->GetStudy(&pStudy);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vLocked(false);
				hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK != hr || VARIANT_FALSE == V_BOOL(&vLocked))
				{
					VsiEkvUtils::DeleteEkvStudyFilesData(pImage);
				}

				pImage.Release();
			}
		}

		// Set update layout event
		VsiSetParamEvent(
			VSI_PDM_ROOT_APP, 
			VSI_PDM_GROUP_EVENTS, 
			VSI_PARAMETER_EVENTS_LAYOUT_UPDATE,
			m_pPdm);
	}
	VSI_CATCH(hr);

	return hr;
}

void CVsiStudyBrowser::DisableSelectionUI()
{
	static const DWORD pdwCtls[] = {
		IDC_SB_ANALYSIS,
		IDC_SB_SONO,
		IDC_SB_STRAIN,
		IDC_SB_STUDY_SERIES_INFO,
		IDC_SB_DELETE,
		IDC_SB_EXPORT,
		IDC_SB_COPY_TO,
		IDC_SB_ACCESS_FILTER,
	};

	for (int i = 0; i < _countof(pdwCtls); ++i)
	{
		GetDlgItem(pdwCtls[i]).EnableWindow(FALSE);
	}
}
