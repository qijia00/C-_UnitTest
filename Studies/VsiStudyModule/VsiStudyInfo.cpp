/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudyInfo.cpp
**
**	Description:
**		Implementation of CVsiStudyInfo
**
*******************************************************************************/

#include "stdafx.h"
#include <ATLComTime.h>
#include <VsiSaxUtils.h>
#include <VsiWtl.h>
#include <VsiWindow.h>
#include <VsiServiceProvider.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiServiceKey.h>
#include <VsiPdmModule.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiParamMeasurement.h>
#include <VsiStudyXml.h>
#include <VsiBuildNumber.h>
#include "VsiAnalRsltXmlTags.h"
#include "VsiStudyInfo.h"

// Commands that this view listen
#define VSI_CMD(_a) _a,
static const DWORD g_pStudyInfoCmds[] =
{
	VSI_CMDS()
};


//
// Date field length
#define VSI_MAX_DATE_LENGTH		10


static const DWORD g_pdwStudyId[] = {
	VSI_PROP_STUDY_NAME, IDC_SI_STUDY_NAME,
	VSI_PROP_STUDY_NOTES, IDC_SI_STUDY_NOTES,
	VSI_PROP_STUDY_GRANTING_INSTITUTION, IDC_SI_GRANTING_INSTITUTION,
};

static const DWORD g_pdwStudyAccessId[] = 
{
	IDC_SI_STUDY_ACCESS_LABEL,
	IDC_SI_STUDY_ACCESS,
};

static const DWORD g_pdwSeriesId[] = {
	VSI_PROP_SERIES_NAME, IDC_SI_SERIES_NAME,
	VSI_PROP_SERIES_NOTES, IDC_SI_SERIES_NOTES,
	VSI_PROP_SERIES_APPLICATION, IDC_SI_STATIC_APPLICATION,
	VSI_PROP_SERIES_MSMNT_PACKAGE, IDC_SI_STATIC_MSMNT_PACKAGE,
	VSI_PROP_SERIES_ANIMAL_ID, IDC_SI_ANIMAL_ID,
	VSI_PROP_SERIES_ANIMAL_COLOR, IDC_SI_COLOR,
	VSI_PROP_SERIES_STRAIN, IDC_SI_HERITAGE,
	VSI_PROP_SERIES_SOURCE, IDC_SI_SOURCE,
	VSI_PROP_SERIES_WEIGHT, IDC_SI_WEIGHT,
	VSI_PROP_SERIES_TYPE, IDC_SI_TYPE,
	VSI_PROP_SERIES_HEART_RATE, IDC_SI_HEART_RATE,
	VSI_PROP_SERIES_BODY_TEMP, IDC_SI_TEMPERATURE,
	VSI_PROP_SERIES_CUSTOM1, IDC_SI_C1,
	VSI_PROP_SERIES_CUSTOM2, IDC_SI_C2,
	VSI_PROP_SERIES_CUSTOM3, IDC_SI_C3,
	VSI_PROP_SERIES_CUSTOM4, IDC_SI_C4,
	VSI_PROP_SERIES_CUSTOM5, IDC_SI_C5,
	VSI_PROP_SERIES_CUSTOM6, IDC_SI_C6,
	VSI_PROP_SERIES_CUSTOM7, IDC_SI_C7,
	VSI_PROP_SERIES_CUSTOM8, IDC_SI_C8,
	VSI_PROP_SERIES_CUSTOM9, IDC_SI_C9,
	VSI_PROP_SERIES_CUSTOM10, IDC_SI_C10,
	VSI_PROP_SERIES_CUSTOM11, IDC_SI_C11,
	VSI_PROP_SERIES_CUSTOM12, IDC_SI_C12,
	VSI_PROP_SERIES_CUSTOM13, IDC_SI_C13,
	VSI_PROP_SERIES_CUSTOM14, IDC_SI_C14,
};

// Special props
static const DWORD g_pdwSeriesDateId[] = {
	VSI_PROP_SERIES_DATE_OF_BIRTH, IDC_SI_DATE_OF_BIRTH,
	VSI_PROP_SERIES_DATE_MATED, IDC_SI_DATE_MATED,
	VSI_PROP_SERIES_DATE_PLUGGED, IDC_SI_DATE_PLUGGED,
};

static const DWORD g_pdwSeriesLabelId[] = {
	IDC_SI_SERIES_INFO,
	IDC_SI_SERIES_NAME_LABEL,
	IDC_SI_APPLICATION_LABEL,
	IDC_SI_MSMNT_PACKAGE_LABEL,
	IDC_SI_ANIMAL_ID_LABEL,
	IDC_SI_COLOR_LABEL,
	IDC_SI_HERITAGE_LABEL,
	IDC_SI_SOURCE_LABEL,
	IDC_SI_WEIGHT_LABEL,
	IDC_SI_TYPE_LABEL,
	IDC_SI_DATE_OF_BIRTH_LABEL,
	IDC_SI_DATE_OF_BIRTH_PICKER,
	IDC_SI_HEART_RATE_LABEL,
	IDC_SI_HEART_RATE_UNIT,
	IDC_SI_TEMPERATURE_LABEL,
	IDC_SI_SEX_LABEL,
	IDC_SI_MALE,
	IDC_SI_FEMALE,
	IDC_SI_DATE_MATED_LABEL,
	IDC_SI_DATE_MATED_PICKER,
	IDC_SI_DATE_PLUGGED_LABEL,
	IDC_SI_DATE_PLUGGED_PICKER,
	IDC_SI_SERIES_NOTES_LABEL,
	IDC_SI_DATEFORMAT1,
	IDC_SI_DATEFORMAT2,
	IDC_SI_DATEFORMAT3,
};

static const DWORD g_pdwCustomId[] = {
	VSI_PROP_SERIES_CUSTOM1, IDC_SI_C1_LABEL, IDC_SI_C1,
	VSI_PROP_SERIES_CUSTOM2, IDC_SI_C2_LABEL, IDC_SI_C2,
	VSI_PROP_SERIES_CUSTOM3, IDC_SI_C3_LABEL, IDC_SI_C3,
	VSI_PROP_SERIES_CUSTOM4, IDC_SI_C4_LABEL, IDC_SI_C4,
	VSI_PROP_SERIES_CUSTOM5, IDC_SI_C5_LABEL, IDC_SI_C5,
	VSI_PROP_SERIES_CUSTOM6, IDC_SI_C6_LABEL, IDC_SI_C6,
	VSI_PROP_SERIES_CUSTOM7, IDC_SI_C7_LABEL, IDC_SI_C7,
	VSI_PROP_SERIES_CUSTOM8, IDC_SI_C8_LABEL, IDC_SI_C8,
	VSI_PROP_SERIES_CUSTOM9, IDC_SI_C9_LABEL, IDC_SI_C9,
	VSI_PROP_SERIES_CUSTOM10, IDC_SI_C10_LABEL, IDC_SI_C10,
	VSI_PROP_SERIES_CUSTOM11, IDC_SI_C11_LABEL, IDC_SI_C11,
	VSI_PROP_SERIES_CUSTOM12, IDC_SI_C12_LABEL, IDC_SI_C12,
	VSI_PROP_SERIES_CUSTOM13, IDC_SI_C13_LABEL, IDC_SI_C13,
	VSI_PROP_SERIES_CUSTOM14, IDC_SI_C14_LABEL, IDC_SI_C14,
};



CVsiStudyInfo::CVsiStudyInfo():
	m_dwSink(0),
	m_ViewState(VSI_VIEW_STATE_HIDE),
	m_state(VSI_SI_STATE_NONE),
	m_bClosed(FALSE),
	m_bAllowChangingOperatorOwner(TRUE),
	m_bAllowChangingOperatorAcq(TRUE),
	m_dwEventOperatorListUpdate(0),
	m_dwEventSysState(0),
	m_hilDate(NULL),
	m_hbrushMandatory(NULL),
	m_bSecureMode(FALSE),
	m_bAdmin(FALSE),
	m_bStudyChanged(false),
	m_bSeriesChanged(false)
{
}

CVsiStudyInfo::~CVsiStudyInfo()
{
	_ASSERT(NULL == m_hilDate);
	_ASSERT(NULL == m_hbrushMandatory);
}

STDMETHODIMP CVsiStudyInfo::Activate(IVsiApp *pApp, HWND hwndParent)
{
	HRESULT hr = S_OK;

	try
	{
		m_pApp = pApp;

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// PDM
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		ULONG ulSysState = VsiGetBitfieldValue<ULONG>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_STATE, m_pPdm);

		// Get Command Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CMD_MGR,
			__uuidof(IVsiCommandManager),
			(IUnknown**)&m_pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Cursor Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CURSOR_MANAGER,
			__uuidof(IVsiCursorManager),
			(IUnknown**)&m_pCursorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Operator Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_OPERATOR_MGR,
			__uuidof(IVsiOperatorManager),
			(IUnknown**)&m_pOperatorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Study Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_STUDY_MGR,
			__uuidof(IVsiStudyManager),
			(IUnknown**)&m_pStudyMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Gets session
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_SESSION_MGR,
			__uuidof(IVsiSession),
			(IUnknown**)&m_pSession);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Scan Format Manager
		if (VSI_SYS_STATE_HARDWARE & ulSysState)
		{
			hr = pServiceProvider->QueryService(
				VSI_SERVICE_SCANFORMAT_MANAGER,
				__uuidof(IVsiScanFormatManager),
				(IUnknown**)&m_pScanFormatMgr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Query Measurement Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_MEASUREMENT_MANAGER,
			__uuidof(IVsiMsmntManager),
			(IUnknown**)&m_pMsmntMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Operator List update event id
		CComPtr<IVsiParameter> pParamEvent;
		hr = m_pPdm->GetParameter(VSI_PDM_ROOT_APP,	VSI_PDM_GROUP_EVENTS,
			VSI_PARAMETER_EVENTS_GENERAL_OPERATOR_LIST_UPDATE,
			&pParamEvent);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParamEvent->GetId(&m_dwEventOperatorListUpdate);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		pParamEvent.Release();
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_STATE,
			&pParamEvent);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParamEvent->GetId(&m_dwEventSysState);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Advise as a sink for commands
		hr = m_pCmdMgr->AdviseCommandSink(
			g_pStudyInfoCmds,
			_countof(g_pStudyInfoCmds),
			VSI_CMD_SINK_TYPE_EXECUTE,
			this);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Events sink
		CComQIPtr<IVsiPdmEvents> pPdmEvents(m_pPdm);
		hr = pPdmEvents->AdviseParamUpdateEventSink(
			static_cast<IVsiParamUpdateEventSink*>(this), &m_dwSink);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HWND hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
		{
			VSI_ERROR_LOG(err);
		}

		Deactivate();
	}

	return hr;
}

LRESULT CVsiStudyInfo::OnWindowPosChanged(UINT, WPARAM, LPARAM lp, BOOL &bHandled)
{
	WINDOWPOS *pPos = (WINDOWPOS*)lp;

	HRESULT hr = S_OK;

	try
	{
		if (SWP_SHOWWINDOW & pPos->flags)
		{
			UpdateUI();

			hr = m_pCursorMgr->SetState(VSI_CURSOR_STATE_ON, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	bHandled = FALSE;

	return 1;  // Act as if nothing happened
}

STDMETHODIMP CVsiStudyInfo::Deactivate()
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

	if (m_pCmdMgr != NULL)
	{
		// Un-advise as a sink for mode commands
		hr = m_pCmdMgr->UnadviseCommandSink(
			g_pStudyInfoCmds,
			_countof(g_pStudyInfoCmds),
			VSI_CMD_SINK_TYPE_EXECUTE,
			this);
		if (FAILED(hr))
		{
			VSI_LOG_MSG(VSI_LOG_WARNING, L"Failure unadvising command sink");
		}

		m_pCmdMgr.Release();
	}
	m_pSeries.Release();
	m_pStudy.Release();
	m_pStudyMgr.Release();
	m_pOperatorMgr.Release();
	m_pPdm.Release();
	m_pApp.Release();

	return hr;
}

STDMETHODIMP CVsiStudyInfo::GetWindow(HWND *phWnd)
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

STDMETHODIMP CVsiStudyInfo::GetAccelerator(HACCEL *phAccel)
{
	UNREFERENCED_PARAMETER(phAccel);
	return E_NOTIMPL;
}

STDMETHODIMP CVsiStudyInfo::GetIsBusy(
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

STDMETHODIMP CVsiStudyInfo::SetState(VSI_SI_STATE state)
{
	_ASSERT(!IsWindow());  // Must set state before the window is activated

	m_state = (VSI_SI_STATE)(state & VSI_SI_STATE_MASK);
	m_bClosed = ((state & VSI_SI_STATE_CLOSED) == VSI_SI_STATE_CLOSED);

	return S_OK;
}

STDMETHODIMP CVsiStudyInfo::SetStudy(IVsiStudy *pStudy)
{
	_ASSERT(!IsWindow());  // Must assign the study before the window is activated

	m_pStudy = pStudy;

	return S_OK;
}

STDMETHODIMP CVsiStudyInfo::SetSeries(IVsiSeries *pSeries)
{
	_ASSERT(!IsWindow());  // Must assign the series before the window is activated

	m_pSeries = pSeries;

	return S_OK;
}

STDMETHODIMP CVsiStudyInfo::OnViewStateChange(VSI_VIEW_STATE state)
{
	m_ViewState = state;

	return S_OK;
}

STDMETHODIMP CVsiStudyInfo::PreTranslateMessage(MSG *pMsg, BOOL *pbHandled)
{
	*pbHandled = IsDialogMessage(pMsg);

	return S_OK;
}

STDMETHODIMP CVsiStudyInfo::OnParamUpdate(DWORD *pdwParamIds, DWORD dwCount)
{
	UNREFERENCED_PARAMETER(pdwParamIds);
	UNREFERENCED_PARAMETER(dwCount);

	if (IsWindow())
	{
		for (DWORD i = 0; i < dwCount; ++i, ++pdwParamIds)
		{
			if ((*pdwParamIds == m_dwEventOperatorListUpdate))
			{
				if (m_bAllowChangingOperatorOwner)
				{
					m_wndOperatorOwner.Fill(NULL);
				}

				if (m_bAllowChangingOperatorAcq)
				{
					m_wndOperatorAcq.Fill(NULL);
				}

				break;
			}
			else if ((*pdwParamIds == m_dwEventSysState))
			{
				ScanformatChange();
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CVsiStudyInfo::Execute(
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

LRESULT CVsiStudyInfo::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_VERIFY(VSI_SI_STATE_NONE != m_state, VSI_LOG_ERROR, NULL);

		// Set font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
		VsiThemeRecurSetFont(m_hWnd, hFont);

		// Layout
		CComPtr<IVsiWindowLayout> pLayout;
		hr = pLayout.CoCreateInstance(__uuidof(VsiWindowLayout));
		if (SUCCEEDED(hr))
		{
			hr = pLayout->Initialize(m_hWnd, VSI_WL_FLAG_AUTO_RELEASE);
			if (SUCCEEDED(hr))
			{				
				pLayout->AddControl(0, IDC_SI_COPY_PREV, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDOK, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDCANCEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_SI_MANDATORY_FIELD, VSI_WL_MOVE_X);

				pLayout.Detach();
			}
		}
		
		m_hbrushMandatory = CreateSolidBrush(VSI_COLOR_MANDATORY);
		VSI_WIN32_VERIFY(NULL != m_hbrushMandatory, VSI_LOG_ERROR, NULL);
		
		m_hilDate = ImageList_LoadImage(
			_AtlBaseModule.GetResourceInstance(),
			MAKEINTRESOURCE(IDB_CALENDAR),
			24,
			1,
			CLR_DEFAULT,
			IMAGE_BITMAP,
			LR_CREATEDIBSECTION);

		UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);

		VSI_SETIMAGE si = { 0 };
		for (int i = 0; i < _countof(si.uStyleFlags); ++i)
		{
			si.uStyleFlags[i] = VSI_BTN_IMAGE_STYLE_DRAWTHEME;
		}
		si.hImageLists[VSI_BTN_IMAGE_UNPUSHED] = m_hilDate;
		si.iImageIndex = 0;

		SendDlgItemMessage(
			IDC_SI_DATE_OF_BIRTH_PICKER,
			uiThemeCmd, VSI_BT_CMD_SETIMAGE, (LPARAM)&si);

		SendDlgItemMessage(
			IDC_SI_DATE_MATED_PICKER,
			uiThemeCmd, VSI_BT_CMD_SETIMAGE, (LPARAM)&si);

		SendDlgItemMessage(
			IDC_SI_DATE_PLUGGED_PICKER,
			uiThemeCmd, VSI_BT_CMD_SETIMAGE, (LPARAM)&si);

		m_bSecureMode = VsiGetBooleanValue<BOOL>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
			VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
		if (m_bSecureMode)
		{
			hr = m_pOperatorMgr->GetIsAdminAuthenticated(&m_bAdmin);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Init operator
		hr = m_wndOperatorOwner.Initialize(
			FALSE,
			TRUE,
			FALSE,
			GetDlgItem(IDC_SI_OPERATOR_NAME),
			m_pOperatorMgr,
			m_pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = m_wndOperatorAcq.Initialize(
			FALSE,
			TRUE,
			FALSE,
			GetDlgItem(IDC_SI_ACQ_OPERATOR_NAME),
			m_pOperatorMgr,
			m_pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Fills operator list
		FillOperatorList();

		// Get the list of available measurement packages.
		PopulatePackageSelection();

		// Must be set
		SendDlgItemMessage(
			IDC_SI_MANDATORY_FIELD,
			uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);

		WORD wIdFocus = IDC_SI_STUDY_NAME;

		switch (m_state)
		{
		case VSI_SI_STATE_NEW_STUDY:
			{
				// Not to change owner and acq by information when in new study
				if (m_bSecureMode)
				{
					UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);
					m_wndOperatorOwner.SendMessage(uiThemeCmd, VSI_CB_CMD_SETREADONLY, 1);
					m_wndOperatorOwner.EnableWindow(FALSE);
					m_wndOperatorAcq.SendMessage(uiThemeCmd, VSI_CB_CMD_SETREADONLY, 1);
					m_wndOperatorAcq.EnableWindow(FALSE);
				}
				else
				{
					SendDlgItemMessage(
						IDC_SI_OPERATOR_NAME_LABEL,
						uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
					SendDlgItemMessage(
						IDC_SI_ACQ_OPERATOR_NAME_LABEL,
						uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
				}

				SendDlgItemMessage(
					IDC_SI_STUDY_NAME_LABEL,
					uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);

				SendDlgItemMessage(
					IDC_SI_SERIES_NAME_LABEL,
					uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);

				SendDlgItemMessage(
					IDC_SI_MSMNT_PACKAGE_LABEL,
					uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
			}
			break;
		case VSI_SI_STATE_NEW_SERIES:
			{
				wIdFocus = IDC_SI_SERIES_NAME;

				// Not to change acq by information when in new series
				if (m_bSecureMode)
				{
					m_wndOperatorAcq.SendMessage(uiThemeCmd, VSI_CB_CMD_SETREADONLY, 1);
					m_wndOperatorAcq.EnableWindow(FALSE);
				}
				else
				{
					SendDlgItemMessage(
						IDC_SI_ACQ_OPERATOR_NAME_LABEL,
						uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
				}

				SendDlgItemMessage(
					IDC_SI_SERIES_NAME_LABEL,
					uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);

				SendDlgItemMessage(
					IDC_SI_MSMNT_PACKAGE_LABEL,
					uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);

				CWindow wndCopyPrev(GetDlgItem(IDC_SI_COPY_PREV));
				wndCopyPrev.EnableWindow();
				wndCopyPrev.ShowWindow(SW_SHOWNA);
			}
			break;
		case VSI_SI_STATE_REVIEW:
			{
				BOOL bStudyLockedAll(FALSE);

				BOOL bLockAll = VsiGetBooleanValue<BOOL>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
					VSI_PARAMETER_SYS_LOCK_ALL, m_pPdm);
				if (bLockAll)
				{
					CComVariant vLocked(false);
					hr = m_pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr && VARIANT_FALSE != V_BOOL(&vLocked))
					{
						bStudyLockedAll = TRUE;
					}
				}
			
				if (!bStudyLockedAll)
				{
					SendDlgItemMessage(
						IDC_SI_STUDY_NAME_LABEL,
						uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);

					SendDlgItemMessage(
						IDC_SI_SERIES_NAME_LABEL,
						uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
				}

				if (bStudyLockedAll)
				{
					wIdFocus = IDCANCEL;
				}
				else
				{
					if (m_bAllowChangingOperatorOwner)
					{
						wIdFocus = IDC_SI_OPERATOR_NAME;
					}
					else if (m_pSeries != NULL)
					{
						if (m_bAllowChangingOperatorAcq)
						{
							wIdFocus = IDC_SI_ACQ_OPERATOR_NAME;
						}
						else
						{
							wIdFocus = IDC_SI_SERIES_NAME;
						}
					}
					else
					{
						wIdFocus = IDC_SI_STUDY_NAME;
					}
				}
			}
			break;
		}

		m_wndDateOfBirth.SubclassWindow(GetDlgItem(IDC_SI_DATE_OF_BIRTH));
		m_wndDateMated.SubclassWindow(GetDlgItem(IDC_SI_DATE_MATED));
		m_wndDatePlugged.SubclassWindow(GetDlgItem(IDC_SI_DATE_PLUGGED));

		// Create Date Picker object
		hr = m_pDatePicker.CoCreateInstance(__uuidof(VsiDatePickerWnd));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (m_bSecureMode)
		{
			m_studyAccess.Initialize(GetDlgItem(IDC_SI_STUDY_ACCESS));
		}

		Initialize();

		ScanformatChange();

		UpdateUI();

		// not to highlight when application field is disabled(transducer disconnected)
		if (m_state == VSI_SI_STATE_NEW_STUDY || m_state == VSI_SI_STATE_NEW_SERIES)
		{
			UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);
			CVsiWindow wndApplication(GetDlgItem(IDC_SI_APPLICATION));
			if (wndApplication.IsWindowEnabled())  
			{
				SendDlgItemMessage(
					IDC_SI_APPLICATION_LABEL,
					uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
			}
		}

		RECT rc;
		GetClientRect(&rc);
		CVsiScrollImpl::SetScrollSize(rc.right, rc.bottom);
		CVsiScrollImpl::GetSystemSettings();

		// Sets focus
		PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)GetDlgItem(wIdFocus), TRUE);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyInfo::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	m_pDatePicker.Release();

	if (NULL != m_hbrushMandatory)
	{
		DeleteObject(m_hbrushMandatory);
		m_hbrushMandatory = NULL;
	}

	{
		CWindow wndPackages(GetDlgItem(IDC_SI_PACKAGE_SELECTION));
		int iCount = (int)wndPackages.SendMessage(CB_GETCOUNT, 0, 0);
		for (int i = 0; i < iCount; ++i)
		{
			CString *pstrOperator = (CString*)wndPackages.SendMessage(CB_GETITEMDATA, i);
			delete pstrOperator;
		}
		wndPackages.SendMessage(CB_RESETCONTENT);
	}

	m_wndOperatorOwner.Uninitialize();
	m_wndOperatorAcq.Uninitialize();

	if (NULL != m_hilDate)
	{
		ImageList_Destroy(m_hilDate);
		m_hilDate = NULL;
	}

	return 0;
}

LRESULT CVsiStudyInfo::OnCtlColorEdit(UINT, WPARAM wp, LPARAM lp, BOOL &bHandled)
{
	BOOL bMandatory(FALSE);

	if ((GetDlgItem(IDC_SI_STUDY_NAME) == (HWND)lp) ||
		(GetDlgItem(IDC_SI_SERIES_NAME) == (HWND)lp))
	{
		if (0 == ::GetWindowTextLength((HWND)lp))
		{
			bMandatory = TRUE;
		}
	}
	else if ((GetDlgItem(IDC_SI_OPERATOR_NAME) == (HWND)lp) ||
		(GetDlgItem(IDC_SI_ACQ_OPERATOR_NAME) == (HWND)lp))
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

LRESULT CVsiStudyInfo::OnBnClickedCopyPrev(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Study Info - Copy Prev");

	HRESULT hr = S_OK;

	try
	{
		CComPtr<IVsiEnumSeries> pEnum;
		hr = m_pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_CREATED_DATE, TRUE, &pEnum, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK == hr)
		{
			CComPtr<IVsiSeries> pSeries;
			// Get 1st
			hr = pEnum->Next(1, &pSeries, NULL);

			if (S_OK == hr && NULL != pSeries)
			{
				SetRedraw(FALSE);

				InitializeSeriesInfo(pSeries, true);

				// Acquired by
				{
					CComVariant vAcquiredBy;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vAcquiredBy);
					if (S_OK == hr)
					{
						SendDlgItemMessage(
							IDC_SI_ACQ_OPERATOR_NAME,
							CB_SELECTSTRING,
							static_cast<WPARAM>(-1),
							reinterpret_cast<LPARAM>(V_BSTR(&vAcquiredBy)));
					}
				}

				// Application
				{
					CComVariant vApplication;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_APPLICATION, &vApplication);
					if (S_OK == hr)
					{
						SendDlgItemMessage(
							IDC_SI_APPLICATION,
							CB_SELECTSTRING,
							static_cast<WPARAM>(-1),
							reinterpret_cast<LPARAM>(V_BSTR(&vApplication)));
					}
				}

				// Measurement package
				{
					CComVariant vMsmntPackage;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_MSMNT_PACKAGE, &vMsmntPackage);
					if (S_OK == hr)
					{
						SendDlgItemMessage(
							IDC_SI_PACKAGE_SELECTION,
							CB_SELECTSTRING,
							static_cast<WPARAM>(-1),
							reinterpret_cast<LPARAM>(V_BSTR(&vMsmntPackage)));
					}
				}
			}
		}
	}
	VSI_CATCH(hr);

	SetRedraw(TRUE);
	RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);

	UpdateUI();

	return 0;
}

LRESULT CVsiStudyInfo::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Study Info - OK");

	// Save settings
	HRESULT hr = S_OK;

	try
	{
		WORD wCtl(0);
		BOOL bValid;
		hr = CheckDates(&wCtl, &bValid);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_FALSE == hr)
		{
			VsiMessageBox(GetTopLevelParent(),
				L"Invalid date is specified",
				CString(MAKEINTRESOURCE(IDS_STYINFO_SERIES_INFORMATION)),
				MB_OK | MB_ICONERROR);
				
			CWindow wndCtl(GetDlgItem(wCtl));
			wndCtl.SendMessage(EM_SETSEL, 0, -1);
			wndCtl.SetFocus();

			VSI_THROW_HR(S_FALSE, VSI_LOG_INFO, NULL);
		}

		if (m_bAllowChangingOperatorOwner)
		{
			int index = (int)m_wndOperatorOwner.SendMessage(CB_GETCURSEL);
			if (CB_ERR != index)
			{
				LPCWSTR pszId = (LPCWSTR)m_wndOperatorOwner.SendMessage(CB_GETITEMDATA, index);

				if (m_bSecureMode && NULL != m_pStudy)
				{
					CComVariant vOwner;
					hr = m_pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// If selected study contains active series, close open series when change owner
					if (S_OK == hr)
					{
						CComPtr<IVsiOperator> pOperator;
						hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, V_BSTR(&vOwner), &pOperator);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (S_OK == hr)
						{
							CComHeapPtr<OLECHAR> pszOperatorId;

							hr = pOperator->GetId(&pszOperatorId);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if (0 != wcscmp(pszId, pszOperatorId))
							{
								if (SelectedStudyActive())
								{
									hr = m_pSession->CheckActiveSeries(TRUE, pszOperatorId, NULL);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								}				
							}
						}
					}
				}
				
				if (!m_bSecureMode)
				{
					// Reset current op
					if (! IS_INTRESOURCE(pszId))
					{
						hr = m_pOperatorMgr->SetCurrentOperator(pszId);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}
			}
		}
		else if (m_bAllowChangingOperatorAcq && !m_bSecureMode)
		{
			int index = (int)m_wndOperatorAcq.SendMessage(CB_GETCURSEL);
			if (CB_ERR != index)
			{
				LPCWSTR pszId = (LPCWSTR)m_wndOperatorAcq.SendMessage(CB_GETITEMDATA, index);
				if (! IS_INTRESOURCE(pszId))
				{
					m_pOperatorMgr->SetCurrentOperator(pszId);
				}
			}
		}

		CVsiWindow wndApplication(GetDlgItem(IDC_SI_APPLICATION));
		if (wndApplication.IsWindowEnabled() && wndApplication.IsWindowVisible())
		{
			CString strApplication;
			GetApplication(strApplication);
			
			SetDlgItemText(IDC_SI_STATIC_APPLICATION, strApplication);

			if (NULL != m_pScanFormat)
			{
				CString strPath;
				if (m_bSecureMode)
				{
					CComHeapPtr<OLECHAR> pszOpDataPath;
					HRESULT hr = m_pOperatorMgr->GetCurrentOperatorDataPath(&pszOpDataPath);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);
					strPath.Format(L"%s\\%s\\", pszOpDataPath, VSI_FOLDER_PRESETS);
				}

				hr = m_pScanFormat->Load(NULL, NULL, strPath, TRUE);  // Reload
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
				hr = m_pScanFormat->ActivateApplication(NULL, VSI_ACTIVATE_APP_FLAGS_KEEP_EKV_SETTINGS);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = m_pScanFormat->ActivateApplication((LPCWSTR)strApplication, VSI_ACTIVATE_APP_FLAGS_KEEP_EKV_SETTINGS);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

		// Measurement Package
		CWindow wndPackages(GetDlgItem(IDC_SI_PACKAGE_SELECTION));
		if (wndPackages.IsWindowEnabled() && wndPackages.IsWindowVisible())
		{
			CString strPackage;
			wndPackages.GetWindowText(strPackage);

			SetDlgItemText(IDC_SI_STATIC_MSMNT_PACKAGE, strPackage);

			GetMsmntPackage(strPackage);

			CComPtr<IVsiMsmntPackage> pPackage;
			hr = m_pMsmntMgr->GetPackage(strPackage, &pPackage);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CString strPackageFileName(strPackage + L".sxml");

			// Find out what is the system package for this package and set it
			CString strSysPackageFileName;

			BOOL bSystem(FALSE);
			hr = pPackage->GetIsSystem(&bSystem);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (bSystem)
			{
				strSysPackageFileName = strPackageFileName;
			}
			else
			{
				CComPtr<IVsiMsmntPackage> pSysPackage;
				hr = pPackage->GetSystemPackage(&pSysPackage);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComHeapPtr<OLECHAR> pszSysPackageFileName;
				hr = pSysPackage->GetPackageFileName(&pszSysPackageFileName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				strSysPackageFileName = pszSysPackageFileName;
			}

			// Update parameters
			VsiSetRangeValue<LPCWSTR>(strPackageFileName,
				VSI_PDM_ROOT_ACQ, VSI_PDM_GROUP_MODE,
				VSI_PARAMETER_SETTINGS_MSMNT_PACKAGE, m_pPdm);

			VsiSetRangeValue<LPCWSTR>(strSysPackageFileName,
				VSI_PDM_ROOT_ACQ, VSI_PDM_GROUP_MODE,
				VSI_PARAMETER_SETTINGS_MSMNT_PACKAGE_SYSTEM, m_pPdm);
		}

		switch (m_state)
		{
		case VSI_SI_STATE_NEW_SERIES:
			{
				// Gets date
				SYSTEMTIME st;
				GetSystemTime(&st);

				COleDateTime date(st);
				CComVariant vDate;
				V_VT(&vDate) = VT_DATE;
				V_DATE(&vDate) = DATE(date);

				// Get Study Id
				CComVariant vStudyId;
				hr = m_pStudy->GetProperty(VSI_PROP_STUDY_ID, &vStudyId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure setting property: study ID");

				// Get Study Path
				CComHeapPtr<OLECHAR> pszStudyPath;
				hr = m_pStudy->GetDataPath(&pszStudyPath);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting data path");

				LPWSTR pszSpt = wcsrchr(pszStudyPath, L'\\');
				CString strStudyFolder(pszStudyPath, (int)(pszSpt - pszStudyPath));

				LPCWSTR pszStudyId = V_BSTR(&vStudyId);

				WCHAR szSeriesId[128];
				VsiGetGuid(szSeriesId, _countof(szSeriesId));
				CComVariant vSeriesId(szSeriesId);
				hr = m_pSeries->SetProperty(VSI_PROP_SERIES_ID, &vSeriesId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure setting property: series ID");

				// Series namespace
				{
					CString strNs;
					strNs.Format(L"%s/%s", pszStudyId, szSeriesId);
					CComVariant vNs((LPCWSTR)strNs);
					hr = m_pSeries->SetProperty(VSI_PROP_SERIES_NS, &vNs);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure setting property: series NS");
				}

				// Set properties
				{
					CString strSeries;
					LPWSTR pszSeries = strSeries.GetBufferSetLength(VSI_FIELD_MAX + 1);
					GetDlgItemText(IDC_SI_SERIES_NAME, pszSeries, VSI_FIELD_MAX + 1);
					strSeries.ReleaseBuffer();
					strSeries.Trim();

					CComVariant vSeriesName(strSeries);
					hr = m_pSeries->SetProperty(VSI_PROP_SERIES_NAME, &vSeriesName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure setting property: series name");
				}

				// Date
				hr = m_pSeries->SetProperty(VSI_PROP_SERIES_CREATED_DATE, &vDate);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure setting property: series created date");

				// Version Created
				CString strSoftware;
				strSoftware.Format(
					L"%u.%u.%u.%u",
					VSI_SOFTWARE_MAJOR, VSI_SOFTWARE_MIDDLE, VSI_SOFTWARE_MINOR, VSI_BUILD_NUMBER);		

				CComVariant vVersionCreated((LPCWSTR)strSoftware);

				hr = m_pSeries->SetProperty(VSI_PROP_SERIES_VERSION_CREATED, &vVersionCreated);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Id
				hr = m_pSeries->SetProperty(VSI_PROP_SERIES_ID_STUDY, &vStudyId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure setting property: series ID study");

				// Acq Operator
				{
					// get current operator
					CString strAcquireByOperator;
					GetAcquireByOperator(strAcquireByOperator);

					CComVariant vAcqOperator((LPCWSTR)strAcquireByOperator);
					hr = m_pSeries->SetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vAcqOperator);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure setting property: series acquired by");
				}

				m_bSeriesChanged = true;
				UpdateLastModified();

				WCHAR szSeriesPath[MAX_PATH];
				WCHAR szSeriesFolder[MAX_PATH];

				hr = VsiCreateUniqueFolder(
					strStudyFolder, L"", szSeriesFolder, _countof(szSeriesFolder));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating unique folder");

				swprintf_s(szSeriesPath, _countof(szSeriesPath), L"%s\\%s",
					szSeriesFolder, CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME)));

				hr = m_pSeries->SetDataPath(szSeriesPath);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure setting series data path");

				// Set study
				hr = m_pSeries->SetStudy(m_pStudy);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure setting study");
			}
			break;

		case VSI_SI_STATE_REVIEW:
			{
				if (m_bAllowChangingOperatorOwner)
				{
					CString strOwnerOperator;
					GetOwnerOperator(strOwnerOperator);

					CComVariant vOwnerOperator((LPCWSTR)strOwnerOperator);
					m_pStudy->SetProperty(VSI_PROP_STUDY_OWNER, &vOwnerOperator);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				if ((m_pSeries != NULL) && (m_bAllowChangingOperatorAcq))
				{
					CString strAcquireByOperator;
					GetAcquireByOperator(strAcquireByOperator);

					CComVariant vAcquireByOperator((LPCWSTR)strAcquireByOperator);
					m_pSeries->SetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vAcquireByOperator);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				}

				UpdateLastModified();
			}
			break;

		case VSI_SI_STATE_NEW_STUDY:
			{
				CreateNewStudy();

				// Owner
				{
					CString strOwnerOperator;
					GetOwnerOperator(strOwnerOperator);

					CComVariant vOwnerOperator((LPCWSTR)strOwnerOperator);
					m_pStudy->SetProperty(VSI_PROP_STUDY_OWNER, &vOwnerOperator);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				// Acquire By
				{
					CString strAcquireByOperator;
					GetAcquireByOperator(strAcquireByOperator);

					CComVariant vAcquireByOperator((LPCWSTR)strAcquireByOperator);
					m_pSeries->SetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vAcquireByOperator);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}
			break;

		default:
			_ASSERT(0);
		}

		switch (m_state)
		{
		case VSI_SI_STATE_NEW_SERIES:
			{
				CWaitCursor wait;

				// save xml file
				SaveChanges(NULL, m_pSeries);

				hr = m_pStudy->AddSeries(m_pSeries);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComPtr<IVsiPropertyBag> pProperties;
				hr = pProperties.CoCreateInstance(__uuidof(VsiPropertyBag));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vStudy(m_pStudy);
				hr = pProperties->Write(L"study", &vStudy);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vSeries(m_pSeries);
				hr = pProperties->Write(L"series", &vSeries);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vSeriesId;
				hr = m_pSeries->GetProperty(VSI_PROP_SERIES_ID, &vSeriesId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pProperties->Write(L"seriesId", &vSeriesId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vSeriesNs;
				hr = m_pSeries->GetProperty(VSI_PROP_SERIES_NS, &vSeriesNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pProperties->Write(L"seriesNS", &vSeriesNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComHeapPtr<OLECHAR> pszSeriesFile;
				hr = m_pSeries->GetDataPath(&pszSeriesFile);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vSeriesFile(pszSeriesFile);
				hr = pProperties->Write(L"seriesFile", &vSeriesFile);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// release series
				m_pSeries.Release();

				CComVariant v(pProperties.p);
				m_pCmdMgr->InvokeCommand(ID_CMD_NEW_SERIES_COMMIT, &v);
			}
			break;

		case VSI_SI_STATE_NEW_STUDY:
			{
				CWaitCursor wait;

				// update XML files
				SaveChanges(m_pStudy, m_pSeries);

				CComPtr<IVsiPropertyBag> pProperties;
				hr = pProperties.CoCreateInstance(__uuidof(VsiPropertyBag));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vStudy(m_pStudy);
				hr = pProperties->Write(L"study", &vStudy);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vSeries(m_pSeries);
				hr = pProperties->Write(L"series", &vSeries);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant v(pProperties.p);
				m_pCmdMgr->InvokeCommand(ID_CMD_NEW_STUDY_COMMIT, &v);
			}
			break;

		case VSI_SI_STATE_REVIEW:
			{
				// Update XML files
				SaveChanges(m_pStudy, NULL);

				if (m_pSeries != NULL)
				{
					// Update XML files
					SaveChanges(NULL, m_pSeries);

					CComVariant v(m_pSeries);
					m_pCmdMgr->InvokeCommand(ID_CMD_STUDY_SERIES_INFO_COMMIT, &v);
				}
				else
				{
					CComVariant v(m_pStudy);
					m_pCmdMgr->InvokeCommand(ID_CMD_STUDY_SERIES_INFO_COMMIT, &v);
				}
			}
			break;

		default:
			_ASSERT(0);
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiStudyInfo::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Study Info - Cancel");

	switch (m_state)
	{
	case VSI_SI_STATE_NEW_STUDY:
		m_pCmdMgr->InvokeCommand(ID_CMD_NEW_STUDY_CANCEL, NULL);
		break;
	case VSI_SI_STATE_NEW_SERIES:
		m_pCmdMgr->InvokeCommand(ID_CMD_NEW_SERIES_CANCEL, NULL);
		break;
	case VSI_SI_STATE_REVIEW:
		m_pCmdMgr->InvokeCommand(ID_CMD_STUDY_SERIES_INFO_CANCEL, NULL);
		break;
	default:
		_ASSERT(0);
	}

	return 0;
}

LRESULT CVsiStudyInfo::OnOperatorOwnerChanged(WORD, WORD, HWND, BOOL&)
{
	int index = (int)m_wndOperatorOwner.SendMessage(CB_GETCURSEL);
	if (CB_ERR != index)
	{
		LPCWSTR pszId = (LPCWSTR)m_wndOperatorOwner.SendMessage(CB_GETITEMDATA, index);
		if (! IS_INTRESOURCE(pszId))
		{
			// Switch "acquired by" too if it is visible, enabled and nothing selected yet
			if (m_bAllowChangingOperatorAcq)
			{
				bool bSwitch(true);

				int indexAcq = m_wndOperatorAcq.SendMessage(CB_GETCURSEL);
				if (CB_ERR != indexAcq)
				{
					LPCWSTR pszId = (LPCWSTR)m_wndOperatorAcq.SendMessage(CB_GETITEMDATA, indexAcq);
					if (! IS_INTRESOURCE(pszId))
					{
						bSwitch = false;
					}
				}

				if (bSwitch)
				{
					m_wndOperatorAcq.SendMessage(CB_SETCURSEL, index, 0);
				}
			}

			m_bStudyChanged = true;
			UpdateUI();
		}
	}

	return 0;
}

LRESULT CVsiStudyInfo::OnOperatorAcqChanged(WORD, WORD, HWND, BOOL&)
{
	int index = (int)m_wndOperatorAcq.SendMessage(CB_GETCURSEL);
	if (CB_ERR != index)
	{
		LPCWSTR pszId = (LPCWSTR)m_wndOperatorAcq.SendMessage(CB_GETITEMDATA, index);
		if (! IS_INTRESOURCE(pszId))
		{
			// Switch "owner" too if it is visible, enabled and nothing selected yet
			if (m_bAllowChangingOperatorOwner)
			{
				bool bSwitch(true);

				int indexAcq = m_wndOperatorOwner.SendMessage(CB_GETCURSEL);
				if (CB_ERR != indexAcq)
				{
					LPCWSTR pszId = (LPCWSTR)m_wndOperatorOwner.SendMessage(CB_GETITEMDATA, indexAcq);
					if (! IS_INTRESOURCE(pszId))
					{
						bSwitch = false;
					}
				}

				if (bSwitch)
				{
					m_wndOperatorOwner.SendMessage(CB_SETCURSEL, index, 0);
				}
			}

			m_bSeriesChanged = true;
			UpdateUI();
		}
	}

	return 0;
}

/// <summary>
/// OnEnStudyChange
/// </summary>
LRESULT CVsiStudyInfo::OnEnStudyChange(WORD, WORD, HWND, BOOL&)
{
	m_bStudyChanged = true;
	UpdateUI();

	return 0;
}

/// <summary>
/// OnEnSeriesChange
/// </summary>
LRESULT CVsiStudyInfo::OnEnSeriesChange(WORD, WORD, HWND, BOOL&)
{
	m_bSeriesChanged = true;

	// remove trigures change of study
	//m_bStudyChanged = true;

	UpdateUI();

	return 0;
}

/// <summary>
/// Handles IDC_SI_MALE command - updates controls state
/// </summary>
LRESULT CVsiStudyInfo::OnCmdMale(WORD, WORD, HWND, BOOL&)
{
	m_bSeriesChanged = true;
	UpdateSexUI(FALSE);

	return 0;
}

/// <summary>
/// Handles IDC_SI_FEMALE command - updates controls state
/// </summary>
LRESULT CVsiStudyInfo::OnCmdFemale(WORD, WORD, HWND, BOOL&)
{
	m_bSeriesChanged = true;
	UpdateSexUI(TRUE);

	return 0;
}

/// <summary>
/// Handles IDC_SI_PREGNANT command - updates controls state
/// </summary>
LRESULT CVsiStudyInfo::OnCmdPregnant(WORD, WORD, HWND, BOOL&)
{
	m_bSeriesChanged = true;
	UpdateSexUI(TRUE);

	return 0;
}

/// <summary>
/// Handles IDC_SI_DATE_OF_BIRTH_PICKER command - bring up date picker control.
/// </summary>
LRESULT CVsiStudyInfo::OnCmdDateOfBirthPicker(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		hr = m_pDatePicker->ActivateCalendar(
			GetDlgItem(IDC_SI_DATE_OF_BIRTH_PICKER),
			GetDlgItem(IDC_SI_DATE_OF_BIRTH),
			&m_stDateOfBirth);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	m_bSeriesChanged = true;

	return 0;
}

/// <summary>
/// Handles IDC_SI_DATE_MATED_PICKER command - bring up date picker control.
/// </summary>
LRESULT CVsiStudyInfo::OnCmdDateMatedPicker(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		hr = m_pDatePicker->ActivateCalendar(
			GetDlgItem(IDC_SI_DATE_MATED_PICKER),
			GetDlgItem(IDC_SI_DATE_MATED),
			&m_stDateMated);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	m_bSeriesChanged = true;

	return 0;
}

/// <summary>
/// Handles IDC_SI_DATE_PLUGGED_PICKER command - bring up date picker control.
/// </summary>
LRESULT CVsiStudyInfo::OnCmdDatePluggedPicker(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		hr = m_pDatePicker->ActivateCalendar(
			GetDlgItem(IDC_SI_DATE_PLUGGED_PICKER),
			GetDlgItem(IDC_SI_DATE_PLUGGED),
			&m_stDatePlugged);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	m_bSeriesChanged = true;

	return 0;
}

HRESULT CVsiStudyInfo::VSI_CMD_HANDLER_IMP(ID_CMD_STUDY_SERIES_INFO)
{
	UNREFERENCED_PARAMETER(pvParam);
	UNREFERENCED_PARAMETER(pbContinueExecute);

	HRESULT hr = S_OK;

	try
	{
		if (VSI_VIEW_STATE_SHOW == m_ViewState)
		{
			switch (m_state)
			{
			case VSI_SI_STATE_REVIEW:
				if (GetDlgItem(IDOK).IsWindowEnabled())
				{
					// Commit and close
					SendMessage(WM_COMMAND, IDOK);
				}
				break;
			default:
				// Ignore
				break;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// Initialize
/// </summary>
void CVsiStudyInfo::Initialize() throw(...)
{
	HRESULT hr = S_OK;

	UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);

	InitializeDateFormat();

	// Set limits to the length of text
	static const DWORD pdwCtlId[] = {
		IDC_SI_GRANTING_INSTITUTION,
		IDC_SI_ANIMAL_ID, IDC_SI_COLOR, IDC_SI_HERITAGE,
		IDC_SI_SOURCE, IDC_SI_WEIGHT, IDC_SI_TYPE,
		IDC_SI_HEART_RATE, IDC_SI_TEMPERATURE,
		IDC_SI_C1, IDC_SI_C2, IDC_SI_C3, IDC_SI_C4, IDC_SI_C5, IDC_SI_C6,
		IDC_SI_C7, IDC_SI_C8, IDC_SI_C9, IDC_SI_C10, IDC_SI_C11, IDC_SI_C12,
	};
	for (int i = 0; i < _countof(pdwCtlId); ++i)
	{
		GetDlgItem(pdwCtlId[i]).SendMessage(EM_SETLIMITTEXT, VSI_FIELD_MAX, 0);
	}

	GetDlgItem(IDC_SI_STUDY_NAME).SendMessage(EM_SETLIMITTEXT, VSI_STUDY_NAME_MAX, 0);
	GetDlgItem(IDC_SI_SERIES_NAME).SendMessage(EM_SETLIMITTEXT, VSI_STUDY_NAME_MAX, 0);
	GetDlgItem(IDC_SI_STUDY_NOTES).SendMessage(EM_SETLIMITTEXT, 300, 0);
	GetDlgItem(IDC_SI_SERIES_NOTES).SendMessage(EM_SETLIMITTEXT, 300, 0);

	CWindow wndStaticApplication(GetDlgItem(IDC_SI_STATIC_APPLICATION));
	CWindow wndApplication(GetDlgItem(IDC_SI_APPLICATION));
	CWindow wndStaticPackages(GetDlgItem(IDC_SI_STATIC_MSMNT_PACKAGE));
	CWindow wndPackages(GetDlgItem(IDC_SI_PACKAGE_SELECTION));

	// User Management Mode
	if (m_bSecureMode)
	{
		for (int i = 0; i < _countof(g_pdwStudyAccessId); ++i)
		{
			CWindow wnd(GetDlgItem(g_pdwStudyAccessId[i]));

			wnd.EnableWindow(TRUE);
			wnd.ShowWindow(SW_SHOWNA);
		}
	}

	switch (m_state)
	{
	case VSI_SI_STATE_NEW_STUDY:
		{
			wndStaticApplication.EnableWindow(FALSE);
			wndStaticApplication.ShowWindow(SW_HIDE);
			wndApplication.EnableWindow(TRUE);
			wndApplication.ShowWindow(SW_SHOWNA);

			wndStaticPackages.EnableWindow(TRUE);
			wndStaticPackages.ShowWindow(SW_HIDE);
			wndPackages.EnableWindow(TRUE);
			wndPackages.ShowWindow(SW_SHOWNA);

			// Uses system institution by default
			CVsiRange<LPCWSTR> pInstitution;
			hr = m_pPdm->GetParameter(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_INSTITUTION,
				&pInstitution);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			SetDlgItemText(IDC_SI_GRANTING_INSTITUTION, pInstitution.GetValue());

			// Study Access
			if (m_bSecureMode)
			{
				CComPtr<IVsiOperator> pOperatorSel;

				hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorSel);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					VSI_OPERATOR_DEFAULT_STUDY_ACCESS access(VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE);
					hr = pOperatorSel->GetDefaultStudyAccess(&access);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE type(CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_NONE);

					switch (access)
					{
					case VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE:
						type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PRIVATE;
						break;
					case VSI_OPERATOR_DEFAULT_STUDY_ACCESS_GROUP:
						type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_GROUP;
						break;
					case VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PUBLIC:
						type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PUBLIC;
						break;
					}

					m_studyAccess.SetStudyAccess(type);
				}
			}

			CreateNewSeries();

			// Sets default series name
			{
				CComVariant vSeriesName(CString(MAKEINTRESOURCE(IDS_STYINFO_SERIES_ONE)));
				hr = m_pSeries->SetProperty(VSI_PROP_SERIES_NAME, &vSeriesName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}
		break;
	
	case VSI_SI_STATE_NEW_SERIES:
		{
			wndStaticApplication.EnableWindow(FALSE);
			wndStaticApplication.ShowWindow(SW_HIDE);
			wndApplication.EnableWindow(TRUE);
			wndApplication.ShowWindow(SW_SHOWNA);

			wndStaticPackages.EnableWindow(TRUE);
			wndStaticPackages.ShowWindow(SW_HIDE);
			wndPackages.EnableWindow(TRUE);
			wndPackages.ShowWindow(SW_SHOWNA);

			CreateNewSeries();

			_ASSERT(m_pStudy != NULL);
			CComQIPtr<IVsiPersistStudy> pPersist(m_pStudy);
			hr = pPersist->LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IVsiEnumSeries> pEnum;
			hr = m_pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_CREATED_DATE, FALSE, &pEnum, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Sets default series name
			CString strName;

			if (S_OK == hr)
			{
				LONG lCount(0);
				BOOL bExists = FALSE;
				do
				{
					bExists = FALSE;
					CString strText(MAKEINTRESOURCE(IDS_STYINFO_SERIES_N));

					strName.Format(strText, ++lCount);

					hr = pEnum->Reset();
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiSeries> pSeries;
					while (pEnum->Next(1, &pSeries, NULL) == S_OK)
					{
						CComVariant vName;
						hr = pSeries->GetProperty(VSI_PROP_SERIES_NAME, &vName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					
						if (strName == V_BSTR(&vName))
						{
							bExists = TRUE;
							break;
						}
					
						pSeries.Release();
					}
				}
				while (bExists);
			}
			else
			{
				CString strText(MAKEINTRESOURCE(IDS_STYINFO_SERIES_N));
				strName.Format(strText, 1);
			}

			CComVariant vSeriesName(strName);
			hr = m_pSeries->SetProperty(VSI_PROP_SERIES_NAME, &vSeriesName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		break;

	default:
		{
			wndStaticApplication.EnableWindow(TRUE);
			wndStaticApplication.ShowWindow(SW_SHOWNA);
			wndApplication.EnableWindow(FALSE);
			wndApplication.ShowWindow(SW_HIDE);

			wndStaticPackages.EnableWindow(TRUE);
			wndStaticPackages.ShowWindow(SW_SHOWNA);
			wndPackages.EnableWindow(FALSE);
			wndPackages.ShowWindow(SW_HIDE);
		}
		break;
	}

	// Sets value
	CComVariant vValue;

	//
	// Sets Study Elements
	if (m_pStudy != NULL)
	{
		for (int i = 0; i < _countof(g_pdwStudyId); i += 2)
		{
			if (NULL != g_pdwStudyId[i])
			{
				vValue.Clear();
				hr = m_pStudy->GetProperty(
					static_cast<VSI_PROP_STUDY>(g_pdwStudyId[i]), &vValue);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK != hr)
				{
					vValue = L"";
				}
				
				SetDlgItemText(g_pdwStudyId[i + 1], V_BSTR(&vValue));
			}
		}

		if (m_bSecureMode)
		{
			CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE type(CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_NONE);

			vValue.Clear();
			hr = m_pStudy->GetProperty(VSI_PROP_STUDY_ACCESS, &vValue);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr)
			{
				if (0 == wcscmp(V_BSTR(&vValue), VSI_STUDY_ACCESS_PRIVATE))
				{
					type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PRIVATE;
				}
				else if (0 == wcscmp(V_BSTR(&vValue), VSI_STUDY_ACCESS_GROUP))
				{
					type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_GROUP;
				}
				else if (0 == wcscmp(V_BSTR(&vValue), VSI_STUDY_ACCESS_PUBLIC))
				{
					type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PUBLIC;
				}
				else
				{
					_ASSERT(0);
					type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PUBLIC;
				}

			}
			else
			{
				// Old studies
				type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PUBLIC;
			}

			m_studyAccess.SetStudyAccess(type);
		}
	}

	//
	// Sets Series Elements
	if (m_pSeries != NULL)
	{
		InitializeSeriesInfo(m_pSeries);
	}
	else
	{
		ShowSeriesControls(FALSE);
	}

	switch (m_state)
	{
	case VSI_SI_STATE_NEW_STUDY:
		{
			SetWindowText(CString(MAKEINTRESOURCE(IDS_STYINFO_NEW_STUDY)));

			if (m_bSecureMode)
			{
				for (int i = 0; i < _countof(g_pdwStudyAccessId); ++i)
				{
					CWindow wnd(GetDlgItem(g_pdwStudyAccessId[i]));
					wnd.EnableWindow();
				}
			}
		}
		break;
	case VSI_SI_STATE_NEW_SERIES:
		{
			SetWindowText(CString(MAKEINTRESOURCE(IDS_STYINFO_NEW_SERIES)));

			LONG lStudies(0);
			hr = m_pStudyMgr->GetStudyCount(&lStudies);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			VSI_VERIFY(lStudies > 0, VSI_LOG_ERROR, L"Invalid option VSI_SI_STATE_NEW_SERIES");			
			
			GetDlgItem(IDC_SI_STUDY_NAME).ShowWindow(SW_SHOW);

			for (int i = 0; i < _countof(g_pdwStudyId); i += 2)
			{
				GetDlgItem(g_pdwStudyId[i + 1]).SendMessage(EM_SETREADONLY, 1);
			}

			CWindow wndAccessCombo((HWND)GetDlgItem(IDC_SI_STUDY_ACCESS).SendMessage(CBEM_GETCOMBOCONTROL));
			wndAccessCombo.SendMessage(uiThemeCmd, VSI_CB_CMD_SETREADONLY, 1);
			GetDlgItem(IDC_SI_STUDY_ACCESS).EnableWindow(FALSE);
		}
		break;
	case VSI_SI_STATE_REVIEW:
		{
			SetWindowText(CString(MAKEINTRESOURCE(IDS_STYINFO_STUDY_INFORMATION)));			

			bool bOwnerDeleted = false;
			bool bAcqByDeleted = false;
			if (m_bSecureMode)
			{
				// Only Admin or same owner can change study access
				BOOL bSameOwner(FALSE);
				CComPtr<IVsiOperator> pOperatorSel;
				hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorSel);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				if (S_OK == hr)
				{
					CComHeapPtr<OLECHAR> pszName;
					hr = pOperatorSel->GetName(&pszName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
 
					CComVariant vOwner(L"");
					hr = m_pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr && (0 != *V_BSTR(&vOwner)))
					{
						if (0 == wcscmp(pszName, V_BSTR(&vOwner)))
						{
							bSameOwner = TRUE;
						}
					}

					CComPtr<IVsiOperator> pOperator;
					hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, V_BSTR(&vOwner), &pOperator);
					if (S_OK != hr) 
					{
						GetDlgItem(IDC_SI_OPERATOR_NAME).SendMessage(CB_SETCUEBANNER, FALSE, (LPARAM)V_BSTR(&vOwner));
						bOwnerDeleted = true;
					}

					if (NULL != m_pSeries)
					{
						CComVariant vAcquireByOperator(L"");
						hr = m_pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vAcquireByOperator);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IVsiOperator> pAcquireByOperator;
						hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, V_BSTR(&vAcquireByOperator), &pAcquireByOperator);
						if (S_OK != hr) 
						{
							GetDlgItem(IDC_SI_ACQ_OPERATOR_NAME).SendMessage(CB_SETCUEBANNER, FALSE, (LPARAM)V_BSTR(&vAcquireByOperator));
							bAcqByDeleted = true;
						}
					}
				}

				// Study contains open series shall not allow to change study access information
				BOOL bStudyActive(FALSE);
				CComPtr<IVsiStudy> pStudyActive;
				hr = m_pSession->GetPrimaryStudy(&pStudyActive);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (pStudyActive != NULL)
				{
					CComVariant v1, v2;
					hr = m_pStudy->GetProperty(VSI_PROP_STUDY_ID, &v1);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);
					hr = pStudyActive->GetProperty(VSI_PROP_STUDY_ID, &v2);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					if (v1 == v2)
					{
						bStudyActive = TRUE;
					}
				}

				if ((bSameOwner || m_bAdmin) && !bStudyActive)
				{
					CWindow wndAccess(GetDlgItem(IDC_SI_STUDY_ACCESS));
					GetDlgItem(IDC_SI_STUDY_ACCESS).EnableWindow(TRUE);
				}
				else
				{
					CWindow wndAccessCombo((HWND)GetDlgItem(IDC_SI_STUDY_ACCESS).SendMessage(CBEM_GETCOMBOCONTROL));
					wndAccessCombo.SendMessage(uiThemeCmd, VSI_CB_CMD_SETREADONLY, 1);
					GetDlgItem(IDC_SI_STUDY_ACCESS).EnableWindow(FALSE);
				}
			}

			if (m_bAllowChangingOperatorOwner)
			{
				// Record old operator Id
				int index = (int)m_wndOperatorOwner.SendMessage(CB_GETCURSEL);
				if ((CB_ERR == index) && !bOwnerDeleted)
				{
					// Highlight the label if not already filled
					SendDlgItemMessage(
						IDC_SI_OPERATOR_NAME_LABEL,
						uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
				}
			}

			if (m_bAllowChangingOperatorAcq)
			{
				int index = (int)m_wndOperatorAcq.SendMessage(CB_GETCURSEL);
				if ((CB_ERR == index) && !bAcqByDeleted)  // not to highlight the label if already filled
				{
					SendDlgItemMessage(
						IDC_SI_ACQ_OPERATOR_NAME_LABEL,
						uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
				}
			}

			// Locked study cannot be modified
			// Get total lock state
			BOOL bLockAll = VsiGetBooleanValue<BOOL>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
				VSI_PARAMETER_SYS_LOCK_ALL, m_pPdm);

			// Check study lock state
			if (bLockAll)
			{
				UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);

				CComVariant vLocked(false);
				hr = m_pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
				if (S_OK == hr && VARIANT_FALSE != V_BOOL(&vLocked))
				{
					GetDlgItem(IDOK).ShowWindow(SW_HIDE);
					GetDlgItem(IDCANCEL).SetWindowText(L"Close");
					SendMessage(DM_SETDEFID, (WPARAM)IDCANCEL, 0);

					for (int i = 0; i < _countof(g_pdwStudyId); i += 2)
					{
						GetDlgItem(g_pdwStudyId[i + 1]).SendMessage(EM_SETREADONLY, 1);
					}

					CWindow wndAccessCombo((HWND)GetDlgItem(IDC_SI_STUDY_ACCESS).SendMessage(CBEM_GETCOMBOCONTROL));
					wndAccessCombo.SendMessage(uiThemeCmd, VSI_CB_CMD_SETREADONLY, 1);
					GetDlgItem(IDC_SI_STUDY_ACCESS).EnableWindow(FALSE);

					GetDlgItem(IDC_SI_OPERATOR_NAME).SendMessage(uiThemeCmd, VSI_CB_CMD_SETREADONLY, 1);
					GetDlgItem(IDC_SI_OPERATOR_NAME).EnableWindow(FALSE);

					GetDlgItem(IDC_SI_STATIC_OPERATOR_NAME).SendMessage(EM_SETREADONLY, 1);
					GetDlgItem(IDC_SI_STATIC_OPERATOR_NAME).EnableWindow(TRUE);

					for (int i = 0; i < _countof(g_pdwSeriesId); i += 2)
					{
						GetDlgItem(g_pdwSeriesId[i + 1]).SendMessage(EM_SETREADONLY, 1);
					}

					for (int i = 0; i < _countof(g_pdwSeriesDateId); i += 2)
					{
						GetDlgItem(g_pdwSeriesDateId[i + 1]).SendMessage(EM_SETREADONLY, 1);
					}

					for (int i = 0; i < _countof(g_pdwCustomId); i += 3)
					{
						GetDlgItem(g_pdwCustomId[i + 2]).SendMessage(EM_SETREADONLY, 1);
					}

					GetDlgItem(IDC_SI_ACQ_OPERATOR_NAME).SendMessage(uiThemeCmd, VSI_CB_CMD_SETREADONLY, 1);
					GetDlgItem(IDC_SI_ACQ_OPERATOR_NAME).EnableWindow(FALSE);
					GetDlgItem(IDC_SI_STATIC_ACQ_OPERATOR_NAME).SendMessage(EM_SETREADONLY, 1);
					GetDlgItem(IDC_SI_STATIC_ACQ_OPERATOR_NAME).EnableWindow(TRUE);
					GetDlgItem(IDC_SI_APPLICATION).EnableWindow(FALSE);
					GetDlgItem(IDC_SI_PACKAGE_SELECTION).EnableWindow(FALSE);
					UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);
					GetDlgItem(IDC_SI_MALE).SendMessage(uiThemeCmd, VSI_BT_CMD_SETREADONLY, 1);
					GetDlgItem(IDC_SI_FEMALE).SendMessage(uiThemeCmd, VSI_BT_CMD_SETREADONLY, 1);
					GetDlgItem(IDC_SI_PREGNANT).SendMessage(uiThemeCmd, VSI_BT_CMD_SETREADONLY, 1);
					GetDlgItem(IDC_SI_DATE_OF_BIRTH_PICKER).EnableWindow(FALSE);
					GetDlgItem(IDC_SI_DATE_MATED_PICKER).EnableWindow(FALSE);
					GetDlgItem(IDC_SI_DATE_PLUGGED_PICKER).EnableWindow(FALSE);
				}
			}
		}
		break;
	default:
		_ASSERT(0);
	}
}

void CVsiStudyInfo::InitializeSeriesInfo(IVsiSeries *pSeries, bool bSkipName) throw(...)
{
	HRESULT hr;
	CComVariant vValue;

	_ASSERT(VSI_PROP_SERIES_NAME == g_pdwSeriesId[0]);
	for (int i = bSkipName ? 2 : 0; i < _countof(g_pdwSeriesId); i += 2)
	{
		if (NULL != g_pdwSeriesId[i])
		{
			vValue.Clear();
			hr = pSeries->GetProperty(static_cast<VSI_PROP_SERIES>(g_pdwSeriesId[i]), &vValue);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	
			if (S_OK != hr)
			{
				vValue = L"";
			}
	
			SetDlgItemText(g_pdwSeriesId[i + 1], V_BSTR(&vValue));
	
			if (g_pdwSeriesId[i] >= VSI_PROP_SERIES_CUSTOM1)
			{
				int iIndex = g_pdwSeriesId[i] - VSI_PROP_SERIES_CUSTOM1;
				int iLabel = g_pdwCustomId[iIndex * 3 + 1];
				GetDlgItem(iLabel).ShowWindow(SW_SHOWNA);
				int iValue = g_pdwCustomId[iIndex * 3 + 2];
				GetDlgItem(iValue).ShowWindow(SW_SHOWNA);
			}
		}
	}

	// Special props
	//	VSI_PROP_SERIES_SEX, IDC_SI_MALE, IDC_SI_FEMALE
	//	VSI_PROP_SERIES_PREGNANT, IDC_SI_PREGNANT,
	{
		vValue.Clear();
		hr = pSeries->GetProperty(VSI_PROP_SERIES_SEX, &vValue);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK == hr)
		{
			if (0 == lstrcmp(VSI_SERIES_XML_VAL_MALE, V_BSTR(&vValue)))
			{
				UpdateSexUI(FALSE);
			}
			else if (0 == lstrcmp(VSI_SERIES_XML_VAL_FEMALE, V_BSTR(&vValue)))
			{
				UpdateSexUI(TRUE);
			}
		}

		vValue.Clear();
		hr = pSeries->GetProperty(VSI_PROP_SERIES_PREGNANT, &vValue);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK == hr)
		{
			if (VARIANT_FALSE != V_BOOL(&vValue))
			{
				CheckDlgButton(IDC_SI_PREGNANT, BST_CHECKED);
				UpdateSexUI(TRUE);
			}
		}
	}

	// Special props
	//	VSI_PROP_SERIES_DATE_OF_BIRTH, IDC_SI_DATE_OF_BIRTH,
	//	VSI_PROP_SERIES_DATE_MATED, IDC_SI_DATE_MATED,
	//	VSI_PROP_SERIES_DATE_PLUGGED, IDC_SI_DATE_PLUGGED,
	SYSTEMTIME* ppst[] = {
		&m_stDateOfBirth,
		&m_stDateMated,
		&m_stDatePlugged
	};
	for (int i = 0; i < _countof(g_pdwSeriesDateId); i += 2)
	{
		WCHAR szDate[100];
		SYSTEMTIME st;

		vValue.Clear();
		hr = pSeries->GetProperty(
			static_cast<VSI_PROP_SERIES>(g_pdwSeriesDateId[i]), &vValue);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK == hr)
		{
			COleDateTime date(vValue);

			date.GetAsSystemTime(*ppst[i / 2]);
			// SystemTime already in a local TMZ
			st = *ppst[i / 2];

			GetDateFormat(
				LOCALE_USER_DEFAULT,
				DATE_SHORTDATE,
				&st,
				NULL,
				szDate,
				_countof(szDate));
			SetDlgItemText(g_pdwSeriesDateId[i + 1], szDate);
		}
	}

	// Custom fields label
	for (int i = 0; i < _countof(g_pdwCustomId); i += 3)
	{
		CComHeapPtr<OLECHAR> pszLabel;
		hr = pSeries->GetPropertyLabel(static_cast<VSI_PROP_SERIES>(g_pdwCustomId[i]), &pszLabel);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		int iLabel = g_pdwCustomId[i + 1];
		int iValue = g_pdwCustomId[i + 2];

		if (S_OK == hr)
		{
			GetDlgItem(iLabel).SetWindowText(pszLabel);
			GetDlgItem(iLabel).ShowWindow(SW_SHOWNA);
			GetDlgItem(iValue).ShowWindow(SW_SHOWNA);
		}
		else
		{
			GetDlgItem(iLabel).ShowWindow(SW_HIDE);
			GetDlgItem(iValue).ShowWindow(SW_HIDE);
		}
	}
}

void CVsiStudyInfo::UpdateUI()
{
	HRESULT hr = S_OK;

	try
	{
		CString strStudy;
		GetDlgItemText(IDC_SI_STUDY_NAME, strStudy);
		strStudy.Trim();

		CString strSeries;
		GetDlgItemText(IDC_SI_SERIES_NAME, strSeries);
		strSeries.Trim();

		switch (m_state)
		{
		case VSI_SI_STATE_NEW_STUDY:
		case VSI_SI_STATE_REVIEW:
			{
				BOOL bOkEnable(!strStudy.IsEmpty());

				if (bOkEnable)
				{
					bool bOwnerDeleted = false;
					bool bAcqByDeleted = false;

					if (NULL != m_pStudy)
					{
						CComVariant vOwner(L"");
						hr = m_pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IVsiOperator> pOperator;
						HRESULT hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, V_BSTR(&vOwner), &pOperator);
						if (S_OK != hr) 
						{
							bOwnerDeleted = true;
						}
					}

					if (NULL != m_pSeries)
					{
						CComVariant vAcquireByOperator(L"");
						hr = m_pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vAcquireByOperator);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IVsiOperator> pAcquireByOperator;
						HRESULT hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, V_BSTR(&vAcquireByOperator), &pAcquireByOperator);
						if (S_OK != hr) 
						{
							bAcqByDeleted = true;
						}
					}					

					// Get operator
					CString strOwnerOperator;
					GetOwnerOperator(strOwnerOperator);

					bOkEnable = (bOwnerDeleted || (!strOwnerOperator.IsEmpty()));
					if (bOkEnable)
					{
						if (NULL != m_pSeries)
						{
							bOkEnable = (bAcqByDeleted || (!strSeries.IsEmpty()));

							if (bOkEnable)
							{
								CString strAcquireByOperator;
								GetAcquireByOperator(strAcquireByOperator);

								bOkEnable =	!strAcquireByOperator.IsEmpty();
							}
						}
					}
				}

				if (NULL != m_pStudy)
				{
					// Get lock all state
					BOOL bLockAll = VsiGetBooleanValue<BOOL>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
						VSI_PARAMETER_SYS_LOCK_ALL, m_pPdm);

					// Check study lock state
					if (bLockAll)
					{
						CComVariant vLocked(false);
						hr = m_pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
						if (VARIANT_FALSE != V_BOOL(&vLocked))
						{
							bOkEnable = FALSE;
						}
					}
				}

				GetDlgItem(IDOK).EnableWindow(bOkEnable);
			}
			break;
			
		case VSI_SI_STATE_NEW_SERIES:
			{
				// Get operator
				CString strAcquireByOperator;
				GetAcquireByOperator(strAcquireByOperator);

				GetDlgItem(IDOK).EnableWindow(
					!strAcquireByOperator.IsEmpty() &&
					!strStudy.IsEmpty() &&
					!strSeries.IsEmpty());
			}
			break;
		default:
			_ASSERT(0);
		}
	}
	VSI_CATCH(hr);
}

// Enable the pregnant, date mated and date plugged fields depending on
// the setting of the SEX control.
void CVsiStudyInfo::UpdateSexUI(BOOL bFemale)
{
	// Set the sex of the mouse. No, we are not giving it an operation we are just
	// turning on the appropriate radio control
	CheckDlgButton(IDC_SI_FEMALE, bFemale ? BST_CHECKED : BST_UNCHECKED);
	CheckDlgButton(IDC_SI_MALE, bFemale ? BST_UNCHECKED : BST_CHECKED);

	static const WORD pwID[] = {
		IDC_SI_DATE_MATED_LABEL,
		IDC_SI_DATE_MATED,
		IDC_SI_DATE_MATED_PICKER,
		IDC_SI_DATE_PLUGGED_LABEL,
		IDC_SI_DATE_PLUGGED,
		IDC_SI_DATE_PLUGGED_PICKER,
		IDC_SI_DATEFORMAT2,
		IDC_SI_DATEFORMAT3
	};

	// If the mouse is male then the pregnant, date of mating and date plugged
	// should be disabled.
	if (bFemale)
	{
		// If the mouse is female enable the "Pregnant" check box.
		CWindow wndPregnant(GetDlgItem(IDC_SI_PREGNANT));
		wndPregnant.EnableWindow(bFemale);
		wndPregnant.ShowWindow(SW_SHOWNA);

		// If the "Pregnant" check box is on enable the date mated & plugged controls
		int isPregnant = (int)wndPregnant.SendMessage(BM_GETCHECK);
		BOOL bEnable = (isPregnant == BST_CHECKED);

		int iCount = _countof(pwID);
		for (int i = 0; i < iCount; i++)
		{
			CWindow wnd(GetDlgItem(pwID[i]));
			wnd.EnableWindow(bEnable);
			wnd.ShowWindow(bEnable ? SW_SHOWNA : SW_HIDE);
		}
	}
	else
	{
		CWindow wndPregnant(GetDlgItem(IDC_SI_PREGNANT));
		wndPregnant.EnableWindow(FALSE);
		wndPregnant.ShowWindow(SW_HIDE);

		int iCount = _countof(pwID);
		for (int i = 0; i < iCount; i++)
		{
			CWindow wnd(GetDlgItem(pwID[i]));
			wnd.EnableWindow(FALSE);
			wnd.ShowWindow(SW_HIDE);
		}
	}
}

// Populate the operator list
void CVsiStudyInfo::FillOperatorList()
{
	HRESULT hr = S_OK;

	try
	{
		CWindow wndStaticOwnerOperator(GetDlgItem(IDC_SI_STATIC_OPERATOR_NAME));
		CWindow wndStaticAcqOperator(GetDlgItem(IDC_SI_STATIC_ACQ_OPERATOR_NAME));

		// UI
		m_wndOperatorOwner.Clear();
		m_wndOperatorAcq.Clear();

		CComPtr<IVsiOperator> pOperatorSel;
		CComHeapPtr<OLECHAR> pszOperatorName;

		hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorSel);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		BOOL bServiceAccount  = FALSE;

		if (S_OK == hr)
		{
			hr = pOperatorSel->GetName(&pszOperatorName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		    VSI_OPERATOR_TYPE dwType;
			hr = pOperatorSel->GetType(&dwType);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			bServiceAccount = 0 != (VSI_OPERATOR_TYPE_SERVICE_MODE & dwType);  
		}

		switch (m_state)
		{
		case VSI_SI_STATE_NEW_STUDY:
			{
				m_wndOperatorOwner.Fill(pszOperatorName);
				m_wndOperatorAcq.Fill(pszOperatorName);

				if (m_bSecureMode)
				{
					m_bAllowChangingOperatorAcq = FALSE;

					wndStaticOwnerOperator.SetWindowText(pszOperatorName);
					wndStaticAcqOperator.SetWindowText(pszOperatorName);

					if (!m_bAdmin || bServiceAccount)
					{
						// Owner operator
						wndStaticOwnerOperator.SendMessage(EM_SETREADONLY, 1);
						wndStaticOwnerOperator.EnableWindow(TRUE);
						wndStaticOwnerOperator.ShowWindow(SW_SHOWNA);

						m_wndOperatorOwner.ShowWindow(SW_HIDE);
						m_wndOperatorOwner.EnableWindow(FALSE);

						m_bAllowChangingOperatorOwner = FALSE;

						// Acquired by
						wndStaticAcqOperator.SendMessage(EM_SETREADONLY, 1);
						wndStaticAcqOperator.EnableWindow(TRUE);
						wndStaticAcqOperator.ShowWindow(SW_SHOWNA);

						m_wndOperatorAcq.ShowWindow(SW_HIDE);
						m_wndOperatorAcq.EnableWindow(FALSE);
					}
				}
			}
			break;

		case VSI_SI_STATE_NEW_SERIES:
			{
				// Owner operator
				CComVariant vOwner;
				hr = m_pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr && (0 < lstrlen(V_BSTR(&vOwner))))
				{
					wndStaticOwnerOperator.SetWindowText((LPCWSTR)V_BSTR(&vOwner));
				}

				wndStaticOwnerOperator.SendMessage(EM_SETREADONLY, 1);
				wndStaticOwnerOperator.EnableWindow(TRUE);
				wndStaticOwnerOperator.ShowWindow(SW_SHOWNA);

				m_wndOperatorOwner.ShowWindow(SW_HIDE);
				m_wndOperatorOwner.EnableWindow(FALSE);

				m_bAllowChangingOperatorOwner = FALSE;

				// Acquired by
				m_wndOperatorAcq.Fill(pszOperatorName);

				if (m_bSecureMode)
				{
					m_bAllowChangingOperatorAcq = FALSE;

					wndStaticAcqOperator.SetWindowText(pszOperatorName);

					if (!m_bAdmin)
					{
						wndStaticAcqOperator.SendMessage(EM_SETREADONLY, 1);
						wndStaticAcqOperator.EnableWindow(TRUE);
						wndStaticAcqOperator.ShowWindow(SW_SHOWNA);

						m_wndOperatorAcq.ShowWindow(SW_HIDE);
						m_wndOperatorAcq.EnableWindow(FALSE);
					}
				}
			}
			break;

		case VSI_SI_STATE_REVIEW:
			{
				// Owner operator
				CComVariant vOwner;
				hr = m_pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr && (0 < lstrlen(V_BSTR(&vOwner))))
				{
					// Not allow to change owner (+ Acq by) for open series study
					if (m_bSecureMode && m_bAdmin && !SelectedStudyActive())
					{
						m_bAllowChangingOperatorOwner = TRUE;
					}
					else
					{
						m_bAllowChangingOperatorOwner = FALSE;
					}
				}
				else
				{
					m_bAllowChangingOperatorOwner = TRUE;
				}

				if (m_bAllowChangingOperatorOwner)
				{
					wndStaticOwnerOperator.ShowWindow(SW_HIDE);
					wndStaticOwnerOperator.EnableWindow(FALSE);

					m_wndOperatorOwner.Fill(V_BSTR(&vOwner));
					m_wndOperatorOwner.EnableWindow(TRUE);
					m_wndOperatorOwner.ShowWindow(SW_SHOWNA);
				}
				else
				{
					wndStaticOwnerOperator.SetWindowText((LPCWSTR)V_BSTR(&vOwner));
					wndStaticOwnerOperator.SendMessage(EM_SETREADONLY, 1);
					wndStaticOwnerOperator.EnableWindow(TRUE);
					wndStaticOwnerOperator.ShowWindow(SW_SHOWNA);

					m_wndOperatorOwner.ShowWindow(SW_HIDE);
					m_wndOperatorOwner.EnableWindow(FALSE);
				}

				if (NULL != m_pSeries)
				{
					// Acq operator
					CComVariant vAcqOper;
					hr = m_pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vAcqOper);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr && (0 < lstrlen(V_BSTR(&vAcqOper))))
					{
						if (m_bClosed)
						{
							m_bAllowChangingOperatorAcq = FALSE;

							wndStaticAcqOperator.SetWindowText((LPCWSTR)V_BSTR(&vAcqOper));
							wndStaticAcqOperator.SendMessage(EM_SETREADONLY, 1);
							wndStaticAcqOperator.EnableWindow(TRUE);
							wndStaticAcqOperator.ShowWindow(SW_SHOWNA);

							m_wndOperatorAcq.ShowWindow(SW_HIDE);
							m_wndOperatorAcq.EnableWindow(FALSE);
						}
						else
						{
							m_wndOperatorAcq.Fill((LPCWSTR)V_BSTR(&vAcqOper));
						}
					}
					else
					{
						m_wndOperatorAcq.Fill(pszOperatorName);
					}
				}
			}
			break;

		default:
			_ASSERT(0);
		}
	}
	VSI_CATCH(hr);
}

void CVsiStudyInfo::GetOwnerOperator(CString &strOperator) throw(...)
{
	strOperator.Empty();

	if (m_bAllowChangingOperatorOwner)
	{
		int iIndex = (int)m_wndOperatorOwner.SendMessage(CB_GETCURSEL);
		if (CB_ERR != iIndex)
		{
			LPCWSTR pszOperator = (LPCWSTR)m_wndOperatorOwner.SendMessage(CB_GETITEMDATA, iIndex);
			if (! IS_INTRESOURCE(pszOperator))
			{
				WCHAR szOper[200];
				int iLen = (int)m_wndOperatorOwner.SendMessage(CB_GETLBTEXT, iIndex, (LPARAM)szOper);
				if (CB_ERR != iLen)
				{
					strOperator = szOper;
				}
			}
		}
	}
	else
	{
		GetDlgItemText(IDC_SI_STATIC_OPERATOR_NAME, strOperator);
	}
}

void CVsiStudyInfo::GetAcquireByOperator(CString &strOperator) throw(...)
{
	strOperator.Empty();

	if (m_bAllowChangingOperatorAcq)
	{
		int iIndex = (int)m_wndOperatorAcq.SendMessage(CB_GETCURSEL);
		if (CB_ERR != iIndex)
		{
			LPCWSTR pszOperator = (LPCWSTR)m_wndOperatorAcq.SendMessage(CB_GETITEMDATA, iIndex);
			if (! IS_INTRESOURCE(pszOperator))
			{
				WCHAR szOper[200];
				int iLen = (int)m_wndOperatorAcq.SendMessage(CB_GETLBTEXT, iIndex, (LPARAM)szOper);
				if (CB_ERR != iLen)
				{
					strOperator = szOper;
				}
			}
		}
	}
	else
	{
		GetDlgItemText(IDC_SI_STATIC_ACQ_OPERATOR_NAME, strOperator);
	}
}

void CVsiStudyInfo::GetMsmntPackage(CString &strPackage) throw(...)
{
	strPackage.Empty();
	CWindow wndPackages(GetDlgItem(IDC_SI_PACKAGE_SELECTION));
	if (wndPackages.IsWindowEnabled())
	{
		int iIndex = (int)wndPackages.SendMessage(CB_GETCURSEL);
		if (CB_ERR != iIndex)
		{
			CString *pstrPackage = (CString*)wndPackages.SendMessage(CB_GETITEMDATA, iIndex);
			strPackage = *pstrPackage;
		}
	}
	else
	{
		GetDlgItemText(IDC_SI_STATIC_MSMNT_PACKAGE, strPackage);
	}
}

void CVsiStudyInfo::GetApplication(CString &strApplication) throw(...)
{
	strApplication.Empty();
	CWindow wndApplication(GetDlgItem(IDC_SI_APPLICATION));

	if (wndApplication.IsWindowEnabled())
	{
		wndApplication.GetWindowText(strApplication);
	}
	else
	{
		GetDlgItemText(IDC_SI_STATIC_APPLICATION, strApplication);
	}
}

void CVsiStudyInfo::SaveChanges(IVsiStudy *pStudy, IVsiSeries *pSeries) throw(...)
{
	HRESULT hr = S_OK;

	// Series
	if (pSeries != NULL)
	{
		for (int i = 0; i < _countof(g_pdwSeriesId); i += 2)
		{
			int iLenMax = (VSI_PROP_SERIES_NOTES == g_pdwSeriesId[i] ? VSI_NOTES_MAX : VSI_FIELD_MAX) + 1;

			CString strFieldName;
			LPWSTR pszFieldName = strFieldName.GetBufferSetLength(iLenMax);
			GetDlgItemText(g_pdwSeriesId[i + 1], pszFieldName, iLenMax);
			strFieldName.ReleaseBuffer();
			strFieldName.Trim();

			// Sets value
			CComVariant vValue(strFieldName);
			hr = pSeries->SetProperty(static_cast<VSI_PROP_SERIES>(g_pdwSeriesId[i]), &vValue);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Special props
		//	VSI_PROP_SERIES_DATE_OF_BIRTH, IDC_SI_DATE_OF_BIRTH,
		{
			CWindow wndDate(GetDlgItem(IDC_SI_DATE_OF_BIRTH));
			if (wndDate.GetWindowTextLength() > 0)
			{
				COleDateTime date(m_stDateOfBirth);
				CComVariant vDate;
				V_VT(&vDate) = VT_DATE;
				V_DATE(&vDate) = DATE(date);
				hr = pSeries->SetProperty(VSI_PROP_SERIES_DATE_OF_BIRTH, &vDate);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else
			{
				CComVariant vEmpty;
				hr = pSeries->SetProperty(VSI_PROP_SERIES_DATE_OF_BIRTH, &vEmpty);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

		//	VSI_PROP_SERIES_SEX, IDC_SI_MALE, IDC_SI_FEMALE
		//	VSI_PROP_SERIES_PREGNANT, IDC_SI_PREGNANT,
		//	VSI_PROP_SERIES_DATE_MATED, IDC_SI_DATE_MATED,
		//	VSI_PROP_SERIES_DATE_PLUGGED, IDC_SI_DATE_PLUGGED,
		if (BST_CHECKED == IsDlgButtonChecked(IDC_SI_MALE))
		{
			CComVariant vValue(VSI_SERIES_XML_VAL_MALE);
			hr = pSeries->SetProperty(VSI_PROP_SERIES_SEX, &vValue);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			vValue.Clear();
			hr = pSeries->SetProperty(VSI_PROP_SERIES_PREGNANT, &vValue);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			hr = pSeries->SetProperty(VSI_PROP_SERIES_DATE_MATED, &vValue);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			hr = pSeries->SetProperty(VSI_PROP_SERIES_DATE_PLUGGED, &vValue);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (BST_CHECKED == IsDlgButtonChecked(IDC_SI_FEMALE))
		{
			CComVariant vValue(VSI_SERIES_XML_VAL_FEMALE);
			hr = pSeries->SetProperty(VSI_PROP_SERIES_SEX, &vValue);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			vValue = (BST_CHECKED == IsDlgButtonChecked(IDC_SI_PREGNANT));
			hr = pSeries->SetProperty(VSI_PROP_SERIES_PREGNANT, &vValue);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CWindow wndDateMated(GetDlgItem(IDC_SI_DATE_MATED));
			if (wndDateMated.GetWindowTextLength() > 0)
			{
				COleDateTime date(m_stDateMated);
				CComVariant vDate;
				V_VT(&vDate) = VT_DATE;
				V_DATE(&vDate) = DATE(date);
				hr = pSeries->SetProperty(VSI_PROP_SERIES_DATE_MATED, &vDate);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else
			{
				CComVariant vEmpty;
				hr = pSeries->SetProperty(VSI_PROP_SERIES_DATE_MATED, &vEmpty);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			CWindow wndDatePlugged(GetDlgItem(IDC_SI_DATE_PLUGGED));
			if (wndDatePlugged.GetWindowTextLength() > 0)
			{
				COleDateTime date(m_stDatePlugged);
				CComVariant vDate;
				V_VT(&vDate) = VT_DATE;
				V_DATE(&vDate) = DATE(date);
				hr = pSeries->SetProperty(VSI_PROP_SERIES_DATE_PLUGGED, &vDate);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else
			{
				CComVariant vEmpty;
				hr = pSeries->SetProperty(VSI_PROP_SERIES_DATE_PLUGGED, &vEmpty);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

		CComQIPtr<IVsiPersistSeries> pPersistSeries(pSeries);
		hr = pPersistSeries->SaveSeriesData(NULL, 0);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}

	// Study
	if (pStudy != NULL)
	{
		for (int i = 0; i < _countof(g_pdwStudyId); i += 2)
		{
			int iLenMax = (VSI_PROP_STUDY_NOTES == g_pdwStudyId[i] ? VSI_NOTES_MAX : VSI_FIELD_MAX) + 1;

			CString strFieldName;
			LPWSTR pszFieldName = strFieldName.GetBufferSetLength(iLenMax);
			GetDlgItemText(g_pdwStudyId[i + 1], pszFieldName, iLenMax);
			strFieldName.ReleaseBuffer();
			strFieldName.Trim();

			// Sets value
			CComVariant vValue(strFieldName);
			hr = pStudy->SetProperty(static_cast<VSI_PROP_STUDY>(g_pdwStudyId[i]), &vValue);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		if (m_bSecureMode)
		{
			// Access info
			CComVariant vStudyAccess;
			CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE type = m_studyAccess.GetStudyAccess();
			switch (type)
			{
			case CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PRIVATE:
				vStudyAccess = VSI_STUDY_ACCESS_PRIVATE;
				break;
			case CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_GROUP:
				vStudyAccess = VSI_STUDY_ACCESS_GROUP;
				break;
			case CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PUBLIC:
				vStudyAccess = VSI_STUDY_ACCESS_PUBLIC;
				break;
			default:
				_ASSERT(0);
				vStudyAccess = VSI_STUDY_ACCESS_PRIVATE;
				break;
			}

			if (VT_BSTR == V_VT(&vStudyAccess))
			{
				hr = pStudy->SetProperty(VSI_PROP_STUDY_ACCESS, &vStudyAccess);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

		CComQIPtr<IVsiPersistStudy> pStudyPersist(pStudy);
		hr = pStudyPersist->SaveStudyData(NULL, 0);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}

	// Signals update event (not for new study or series)
	if (VSI_SI_STATE_REVIEW == m_state)
	{
		VsiSetParamEvent(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
			VSI_PARAMETER_EVENTS_GENERAL_STUDY_INFO_UPDATE, m_pPdm);
	}
}

void CVsiStudyInfo::ShowSeriesControls(BOOL bShow)
{
	for (int i = 0; i < _countof(g_pdwSeriesId); i += 2)
	{
		GetDlgItem(g_pdwSeriesId[i+1]).ShowWindow(bShow ? SW_SHOWNA : SW_HIDE);
	}

	for (int i = 0; i < _countof(g_pdwSeriesDateId); i += 2)
	{
		GetDlgItem(g_pdwSeriesDateId[i+1]).ShowWindow(bShow ? SW_SHOWNA : SW_HIDE);
	}

	for (int i = 0; i < _countof(g_pdwSeriesLabelId); ++i)
	{
		GetDlgItem(g_pdwSeriesLabelId[i]).ShowWindow(bShow ? SW_SHOWNA : SW_HIDE);
	}

	// Acq Operator
	GetDlgItem(IDC_SI_ACQ_OPERATOR_NAME_LABEL).ShowWindow(bShow ? SW_SHOWNA : SW_HIDE);
	GetDlgItem(IDC_SI_ACQ_OPERATOR_NAME).ShowWindow(bShow ? SW_SHOWNA : SW_HIDE);
	GetDlgItem(IDC_SI_PACKAGE_SELECTION).ShowWindow(bShow ? SW_SHOWNA : SW_HIDE);
}

void CVsiStudyInfo::CreateNewSeries() throw(...)
{
	HRESULT hr = S_OK;

	// Create new series

	m_pSeries.Release();

	hr = m_pSeries.CoCreateInstance(__uuidof(VsiSeries));
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	// Loads custom labels
	{
		// Get data path
		CVsiRange<LPCWSTR> pParamPathSettings;
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_PATH_SETTINGS, &pParamPathSettings);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiParameterRange> pRange(pParamPathSettings.m_pParam);
		VSI_CHECK_INTERFACE(pRange, VSI_LOG_ERROR, NULL);

		CComVariant vPath;
		hr = pRange->GetValue(&vPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CString strPath(V_BSTR(&vPath));
		strPath += L"\\" VSI_SERIES_FIELDS_FILE;

		CComQIPtr<IVsiPersistSeries> pPersist(m_pSeries);
		VSI_CHECK_INTERFACE(pPersist, VSI_LOG_ERROR, NULL);

		hr = pPersist->LoadPropertyLabel(strPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
}

void CVsiStudyInfo::CreateNewStudy() throw(...)
{
	HRESULT hr = S_OK;

	// Gets date
	SYSTEMTIME st;
	GetSystemTime(&st);

	COleDateTime date(st);
	CComVariant vDate;
	V_VT(&vDate) = VT_DATE;
	V_DATE(&vDate) = DATE(date);

	// Create new study
	if (NULL != m_pStudy)
		m_pStudy.Release();

	hr = m_pStudy.CoCreateInstance(__uuidof(VsiStudy));
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	// ID
	WCHAR szStudyId[128];
	VsiGetGuid(szStudyId, _countof(szStudyId));
	CComVariant vStudyId(szStudyId);
	hr = m_pStudy->SetProperty(VSI_PROP_STUDY_ID_CREATED, &vStudyId);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	hr = m_pStudy->SetProperty(VSI_PROP_STUDY_ID, &vStudyId);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	hr = m_pStudy->SetProperty(VSI_PROP_STUDY_NS, &vStudyId);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	// Date
	hr = m_pStudy->SetProperty(VSI_PROP_STUDY_CREATED_DATE, &vDate);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								
	// Data path
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

	WCHAR szStudyPath[MAX_PATH];
	WCHAR szStudyFolder[MAX_PATH];

	hr = VsiCreateUniqueFolder(
		V_BSTR(&vPath), L"", szStudyFolder, _countof(szStudyFolder));
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	swprintf_s(szStudyPath, _countof(szStudyPath), L"%s\\%s",
		szStudyFolder, CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));
	hr = m_pStudy->SetDataPath(szStudyPath);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	WCHAR szSeriesId[128];
	VsiGetGuid(szSeriesId, _countof(szSeriesId));
	CComVariant vSeriesId(szSeriesId);
	hr = m_pSeries->SetProperty(VSI_PROP_SERIES_ID, &vSeriesId);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	// Series namespace
	{
		CString strNs;
		strNs.Format(L"%s/%s", szStudyId, szSeriesId);
		CComVariant vNs((LPCWSTR)strNs);
		hr = m_pSeries->SetProperty(VSI_PROP_SERIES_NS, &vNs);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}

	// Date
	hr = m_pSeries->SetProperty(VSI_PROP_SERIES_CREATED_DATE, &vDate);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	// Version Created
	CString strSoftware;
	strSoftware.Format(
		L"%u.%u.%u.%u",
		VSI_SOFTWARE_MAJOR, VSI_SOFTWARE_MIDDLE, VSI_SOFTWARE_MINOR, VSI_BUILD_NUMBER);		

	CComVariant vVersionCreated((LPCWSTR)strSoftware);

	hr = m_pSeries->SetProperty(VSI_PROP_SERIES_VERSION_CREATED, &vVersionCreated);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	hr = m_pSeries->SetProperty(VSI_PROP_SERIES_ID_STUDY, &vStudyId);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	m_bStudyChanged = true;
	m_bSeriesChanged = true;
	UpdateLastModified();

	WCHAR szSeriesPath[MAX_PATH];
	WCHAR szSeriesFolder[MAX_PATH];

	hr = VsiCreateUniqueFolder(
		szStudyFolder, L"", szSeriesFolder, _countof(szSeriesFolder));
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	_ASSERT(m_pSeries);

	// set study
	hr = m_pSeries->SetStudy(m_pStudy);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	swprintf_s(szSeriesPath, _countof(szSeriesPath), L"%s\\%s",
		szSeriesFolder, CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME)));

	hr = m_pSeries->SetDataPath(szSeriesPath);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
}

HRESULT CVsiStudyInfo::ScanformatChange()
{
	HRESULT hr = S_OK;

	try
	{
		BOOL bCheckTransducer(FALSE);

		switch (m_state)
		{
		case VSI_SI_STATE_NEW_STUDY:
		case VSI_SI_STATE_NEW_SERIES:
			bCheckTransducer = TRUE;
			break;
		}
			
		CWindow wndApplication(GetDlgItem(IDC_SI_APPLICATION));
		if (bCheckTransducer && (NULL != m_pScanFormatMgr) && NULL != m_pSeries)//only do this if combo is enabled
		{
			CVsiBitfield<ULONG> pState;
			hr = m_pPdm->GetParameter(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_SYS_STATE, &pState);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			wndApplication.SendMessage(CB_RESETCONTENT);

			m_pScanFormat.Release();

			BOOL bTransducerConnected = (pState.GetValue() & VSI_SYS_STATE_TRANSDUCER_CONNECT);
			if (bTransducerConnected)
			{
				// Fill the applications combo
				hr = m_pScanFormatMgr->GetActiveScanformat(&m_pScanFormat);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (NULL == m_pScanFormat)
				{
					CVsiMemory<DWORD> pParamId;

					hr = m_pPdm->GetParameter(
						VSI_PDM_ROOT_HARDWARE, VSI_PDM_GROUP_GLOBAL,
						VSI_PARAMETER_HW_SYSCONT_PROBE_ID,
						&pParamId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
					BYTE id = static_cast<BYTE>(pParamId.GetSingleValue());
					if (0xFF != id)
					{
						CComPtr<IVsiEnumScanFormatInfo> pEnumInfo;
						HRESULT hr = m_pScanFormatMgr->GetScanformatInfo(id, &pEnumInfo);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (S_OK == hr)
						{
							// Pick the 1st one
							CComPtr<IVsiScanFormatInfo> pInfo;
							if (S_OK == pEnumInfo->Next(1, &pInfo, NULL))
							{
								CComHeapPtr<OLECHAR> pszPath;
								hr = pInfo->GetPath(&pszPath);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								hr = m_pScanFormatMgr->LoadScanformat(pszPath, TRUE, &m_pScanFormat);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							}
						}
					}
				}

				if (NULL != m_pScanFormat)
				{
					CComPtr<IEnumString> pEnumApp;
					hr = m_pScanFormat->GetApplications(&pEnumApp);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
					CComHeapPtr<OLECHAR> pszApp;
					while (S_OK == pEnumApp->Next(1, &pszApp, NULL))
					{
						CComPtr<IVsiApplication> pApp;
						hr = m_pScanFormat->GetApplication(pszApp, &pApp);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
						CComHeapPtr<OLECHAR> pszDesc;
						hr = pApp->GetDescription(&pszDesc);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						int iIndex = (int)wndApplication.SendMessage(CB_ADDSTRING, 0, (LPARAM)pszDesc.m_pData);
						VSI_VERIFY(CB_ERR != iIndex, VSI_LOG_ERROR, NULL);

						pszApp.Free();
					}

					// Selection
					{
						CComHeapPtr<OLECHAR> pszDefaultApp;
						hr = m_pScanFormat->GetDefaultApplication(&pszDefaultApp);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						int iIndex = wndApplication.SendMessage(CB_FINDSTRINGEXACT, (WPARAM)-1, (LPARAM)pszDefaultApp.operator wchar_t *());
						if (CB_ERR == iIndex)
						{
							// Select 1st
							iIndex = 0;
						}

						wndApplication.SendMessage(CB_SETCURSEL, iIndex);
					}
				}
			}

			wndApplication.EnableWindow(bTransducerConnected);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
///	Populates the package selection drop down from the currently installed measurement
///	packages.
/// </summary>
HRESULT CVsiStudyInfo::PopulatePackageSelection()
{
	HRESULT hr = S_OK;
	try
	{
		CWindow wndPackages(GetDlgItem(IDC_SI_PACKAGE_SELECTION));
		CWindow wndStaticPackages(GetDlgItem(IDC_SI_STATIC_MSMNT_PACKAGE));

		CComVariant vPackage;
		if (NULL != m_pSeries)
		{
			hr = m_pSeries->GetProperty(VSI_PROP_SERIES_MSMNT_PACKAGE, &vPackage);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		if ((VT_BSTR == V_VT(&vPackage)) && (0 < lstrlen(V_BSTR(&vPackage))))
		{
			wndStaticPackages.SetWindowText(V_BSTR(&vPackage));

			wndStaticPackages.EnableWindow(TRUE);
			wndStaticPackages.ShowWindow(SW_SHOWNA);

			wndPackages.EnableWindow(FALSE);
			wndPackages.ShowWindow(SW_HIDE);
		}
		else
		{
			// Get the name of the selected package.
			CString strCurrentPackage(VsiGetRangeValue<CString>(
				VSI_PDM_ROOT_ACQ, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_SETTINGS_MSMNT_PACKAGE, m_pPdm));

			PathRemoveExtension(strCurrentPackage.GetBuffer());
			strCurrentPackage.ReleaseBuffer();

			CComPtr<IVsiEnumMsmntPackage> pEnumPackages;
			hr = m_pMsmntMgr->GetPackages(&pEnumPackages);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IVsiMsmntPackage> pPackage;
			while (S_OK == pEnumPackages->Next(1, &pPackage, NULL))
			{
				// Get the description and add it to the drop down
				CComHeapPtr<OLECHAR> pszDescription;
				hr = pPackage->GetDescription(&pszDescription);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComHeapPtr<OLECHAR> pszName;
				hr = pPackage->GetPackageName(&pszName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Add the measurement package to the list control and attach the file name to it.
				int iIndex = (int)wndPackages.SendMessage(CB_ADDSTRING, 0, (LPARAM)(LPWSTR)pszDescription);
				if (CB_ERR != iIndex)
				{
					wndPackages.SendMessage(CB_SETITEMDATA, iIndex, (LPARAM)new CString(pszName));

					if (0 == strCurrentPackage.CompareNoCase(pszName))
					{
						wndPackages.SendMessage(CB_SETCURSEL, iIndex, 0);
					}
				}

				pPackage.Release();
			}

			// Select the first if none selected
			if (CB_ERR == wndPackages.SendMessage(CB_GETCURSEL))
			{
				wndPackages.SendMessage(CB_SETCURSEL, 0, 0);
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

void CVsiStudyInfo::InitializeDateFormat()
{
	// Get presentation format of the date
	WCHAR szSDateFormat[81];
	GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, szSDateFormat, _countof(szSDateFormat));

	_wcslwr_s(szSDateFormat, _countof(szSDateFormat));

	// Show date as 'dd', month as 'mm' and year as 'yyyy' for now
	CString strDate;
	LPCWSTR pszDate = szSDateFormat;
	while (NULL != *pszDate)
	{
		WCHAR cSkip = 0;
	
		switch (*pszDate)
		{
		case L'd':
			strDate += L"dd";
			cSkip = L'd';
			break;
			
		case L'm':
			strDate += L"mm";
			cSkip = L'm';
			break;

		case L'y':
			strDate += L"yyyy";
			cSkip = L'y';
			break;
			
		default:
			strDate += *pszDate;
		}
		
		if (0 != cSkip)
		{
			while (*pszDate == cSkip)
				++pszDate;
		}
		else
		{
			++pszDate;
		}
	}

	SetDlgItemText(IDC_SI_DATEFORMAT1, strDate);
	SetDlgItemText(IDC_SI_DATEFORMAT2, strDate);
	GetDlgItem(IDC_SI_DATEFORMAT2).ShowWindow(SW_HIDE);
	SetDlgItemText(IDC_SI_DATEFORMAT3, strDate);
	GetDlgItem(IDC_SI_DATEFORMAT3).ShowWindow(SW_HIDE);

	// Sets limits
	SendDlgItemMessage(IDC_SI_DATE_OF_BIRTH, EM_SETLIMITTEXT, VSI_MAX_DATE_LENGTH);
	SendDlgItemMessage(IDC_SI_DATE_MATED, EM_SETLIMITTEXT, VSI_MAX_DATE_LENGTH);
	SendDlgItemMessage(IDC_SI_DATE_PLUGGED, EM_SETLIMITTEXT, VSI_MAX_DATE_LENGTH);
}

HRESULT CVsiStudyInfo::CheckDates(WORD *pwCtl, BOOL *pbValid)
{
	HRESULT hr = S_OK;

	*pbValid = FALSE;

	try
	{
		// Read date of birth
		WCHAR szDate[100];
		GetDlgItemText(IDC_SI_DATE_OF_BIRTH, szDate, _countof(szDate));

		SYSTEMTIME stDate;
		hr = m_pDatePicker->ConvertStrToSystemDate(szDate, &stDate);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_FALSE == hr && lstrlen(szDate) > 0)
		{
			// Date provided is not valid
			*pwCtl = IDC_SI_DATE_OF_BIRTH;
			VSI_THROW_HR(S_FALSE, VSI_LOG_INFO, L"No valid Date of Birth provided");
		}

		if (lstrlen(szDate) > 0)
		{
			*pbValid = TRUE;
			m_stDateOfBirth = stDate;
		}
		else // empty string removes the date
		{
			hr = S_OK;
		}

		if (BST_CHECKED == IsDlgButtonChecked(IDC_SI_PREGNANT))
		{
			// Mated date
			GetDlgItemText(IDC_SI_DATE_MATED, szDate, _countof(szDate));

			hr = m_pDatePicker->ConvertStrToSystemDate(szDate, &stDate);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_FALSE == hr && lstrlen(szDate) > 0)
			{
				// Date provided is not valid
				*pwCtl = IDC_SI_DATE_MATED;
				VSI_THROW_HR(S_FALSE, VSI_LOG_INFO, L"No valid Date Mated provided");
			}

			if (lstrlen(szDate) > 0)
				m_stDateMated = stDate;
			else // empty string removes the date
				hr = S_OK;

			// Plugged date
			GetDlgItemText(IDC_SI_DATE_PLUGGED, szDate, _countof(szDate));

			hr = m_pDatePicker->ConvertStrToSystemDate(szDate, &stDate);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_FALSE == hr && lstrlen(szDate) > 0)
			{
				// Date provided is not valid
				*pwCtl = IDC_SI_DATE_PLUGGED;
				VSI_THROW_HR(S_FALSE, VSI_LOG_INFO, L"No valid Date Plugged provided");
			}

			if (lstrlen(szDate) > 0)
				m_stDatePlugged = stDate;
			else // empty string removes the date
				hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

void CVsiStudyInfo::UpdateLastModified()
{
	HRESULT hr = S_OK;
 
	try
	{
		bool bOwnerExist = false;
		CComVariant vOwner(L"");
		CComVariant vDate;
		if(m_bStudyChanged || m_bSeriesChanged)
		{
			// Get current operator name
			CComPtr<IVsiOperator> pOperatorCurrent;
 
			hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorCurrent);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			if (NULL != pOperatorCurrent)   // Safe to assume not null
			{
				CComHeapPtr<OLECHAR> pszCurrentOperName;
				hr = pOperatorCurrent->GetName(&pszCurrentOperName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
 
				vOwner = pszCurrentOperName;
				bOwnerExist = true;
			}

			// Gets date
			SYSTEMTIME st;
			GetSystemTime(&st);

			COleDateTime date(st);
			V_VT(&vDate) = VT_DATE;
			V_DATE(&vDate) = DATE(date);
		}

		if(m_bStudyChanged)
		{
			if(bOwnerExist)
			{
				hr = m_pStudy->SetProperty(VSI_PROP_STUDY_LAST_MODIFIED_BY, &vOwner);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);	
			}

			hr = m_pStudy->SetProperty(VSI_PROP_STUDY_LAST_MODIFIED_DATE, &vDate);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Reset
			m_bStudyChanged = false;
		}

		if(m_bSeriesChanged)
		{
			if(bOwnerExist)
			{
				hr = m_pSeries->SetProperty(VSI_PROP_SERIES_LAST_MODIFIED_BY, &vOwner);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);	
			}

			hr = m_pSeries->SetProperty(VSI_PROP_SERIES_LAST_MODIFIED_DATE, &vDate);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Reset
			m_bSeriesChanged = false;
		}
	}
	VSI_CATCH(hr);

}

bool CVsiStudyInfo::SelectedStudyActive()
{
	bool bActive = false;
	HRESULT hr(S_OK);

	CComVariant vOwner;
	hr = m_pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	if (S_OK == hr)
	{
		CComPtr<IVsiOperator> pOperator;
		hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, V_BSTR(&vOwner), &pOperator);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		if (S_OK == hr)
		{
			CComHeapPtr<OLECHAR> pszOperatorId;

			hr = pOperator->GetId(&pszOperatorId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszSeriesNs;
			hr = m_pSession->CheckActiveSeries(FALSE, pszOperatorId, &pszSeriesNs);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (NULL != pszSeriesNs)
			{
				CComPtr<IUnknown> pUnkSeries;
				hr = m_pStudyMgr->GetItem(pszSeriesNs, &pUnkSeries);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					CComQIPtr<IVsiSeries> pSeries(pUnkSeries);
					VSI_CHECK_INTERFACE(pSeries, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiStudy> pStudy;
					hr = pSeries->GetStudy(&pStudy);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						CComVariant vActiveId;
						hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vActiveId);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						if (S_OK == hr)
						{
							CComVariant vId;
							hr = m_pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							if (S_OK == hr)
							{
								if (vActiveId == vId)
								{
									bActive = true;
								}
							}
						}
					}
				}
			}
		}
	}
	
	return bActive;
}