/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiMoveSeries.cpp
**
**	Description:
**		Implementation of CVsiMoveSeries
**
********************************************************************************/

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
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiStudyXml.h>
#include <VsiStudyModule.h>
#include "VsiMoveSeries.h"


// CVsiMoveSeries

CVsiMoveSeries::CVsiMoveSeries() :
	m_iSelectedStudies(0)
{
}

CVsiMoveSeries::~CVsiMoveSeries()
{
	_ASSERT(m_pApp == NULL);
	_ASSERT(m_pCmdMgr == NULL);
	_ASSERT(m_pPdm == NULL);
	_ASSERT(m_pStudyMgr == NULL);
}

STDMETHODIMP CVsiMoveSeries::Activate(
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

		// Gets session
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_SESSION_MGR,
			__uuidof(IVsiSession),
			(IUnknown**)&m_pSession);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HWND hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiMoveSeries::Deactivate()
{
	HRESULT hr = S_FALSE;

	if (IsWindow())
	{
		BOOL bRet = DestroyWindow();
		_ASSERT(bRet);

		hr = S_OK;
	}

	m_pSession.Release();

	m_pPdm.Release();

	m_pCmdMgr.Release();

	m_pStudyMgr.Release();

	m_pApp.Release();

	return hr;
}

STDMETHODIMP CVsiMoveSeries::GetWindow(HWND *phWnd)
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

STDMETHODIMP CVsiMoveSeries::GetAccelerator(HACCEL *phAccel)
{
	UNREFERENCED_PARAMETER(phAccel);
	return E_NOTIMPL;
}

STDMETHODIMP CVsiMoveSeries::GetIsBusy(
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

STDMETHODIMP CVsiMoveSeries::SetSeries(IVsiSeries *pSeries)
{
	HRESULT hr = S_OK;

	try
	{
		m_pSeries = pSeries;
		VSI_CHECK_INTERFACE(m_pSeries, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

LRESULT CVsiMoveSeries::OnInitDialog(UINT, WPARAM, LPARAM, BOOL& bHandled)
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
		hr = pLayout.CoCreateInstance(__uuidof(VsiWindowLayout));
		if (SUCCEEDED(hr))
		{
			hr = pLayout->Initialize(m_hWnd, VSI_WL_FLAG_AUTO_RELEASE);
			if (SUCCEEDED(hr))
			{
				pLayout->AddControl(0, IDOK, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDCANCEL, VSI_WL_MOVE_X);

				pLayout->AddControl(0, IDC_MS_STUDY_LIST, VSI_WL_SIZE_XY);
				pLayout->AddControl(0, IDC_MS_INFO, VSI_WL_MOVE_X | VSI_WL_SIZE_Y);

				pLayout->AddControl(0, IDC_IE_SELECTED, VSI_WL_MOVE_X);

				pLayout.Detach();
			}
		}

		// Initialize Study List
		static VSI_LVCOLUMN plvc[] =
		{
			{ LVCF_WIDTH | LVCF_TEXT, 0, 500, IDS_MOVESER_COLUMN_NAME, VSI_DATA_LIST_COL_NAME, -1, 0, TRUE },
			{ LVCF_WIDTH | LVCF_TEXT, 0, 120, IDS_MOVESER_COLUMN_DATE, VSI_DATA_LIST_COL_DATE, -1, 1, TRUE },
			{ LVCF_WIDTH | LVCF_TEXT, 0, 200, IDS_MOVESER_COLUMN_OWNER, VSI_DATA_LIST_COL_OWNER, -1, 2, TRUE },
		};

		hr = m_pDataList.CoCreateInstance(__uuidof(VsiDataListWnd));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating Data List");

		hr = m_pDataList->Initialize(
			GetDlgItem(IDC_MS_STUDY_LIST),
			VSI_DATA_LIST_FLAG_ITEM_STATUS_CALLBACK,
			VSI_DATA_LIST_STUDY, m_pApp, plvc, _countof(plvc));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = m_pDataList->SetEmptyMessage(L"Loading studies...");
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HFONT hFontNormal, hFontReviewed;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFontNormal);
		VsiThemeGetFont(VSI_THEME_FONT_M_T, &hFontReviewed);
		hr = m_pDataList->SetFont(hFontNormal, hFontReviewed);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = m_pDataList->SetTextColor(
			VSI_DATA_LIST_ITEM_STATUS_ACTIVE,
			VSI_COLOR_ITEM_ACTIVE);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Sorting
		{
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
		}

		// Init study overview control
		m_studyOverview.SubclassWindow(GetDlgItem(IDC_MS_INFO));
		m_studyOverview.Initialize();

		SetWindowText(CString(MAKEINTRESOURCE(IDS_MOVESER_MOVE_SELECTED_SERIES)));

		// Delay load study list
		PostMessage(WM_VSI_SB_FILL_LIST);

		UpdateUI();
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiMoveSeries::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	return 0;
}

LRESULT CVsiMoveSeries::OnUpdatePreview(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr;

	try
	{
		CComPtr<IVsiEnumStudy> pEnumStudy;
		CComPtr<IVsiEnumSeries> pEnumSeries;
		int iStudy = 0, iSeries = 0; //, iStudyPreFiltered = 0, iSeriesPreFiltered = 0;
		VSI_DATA_LIST_COLLECTION selection;

		memset((void*)&selection, 0, sizeof(VSI_DATA_LIST_COLLECTION));
		selection.dwFlags = VSI_DATA_LIST_FILTER_SELECT_PARENT;
		selection.ppEnumStudy = &pEnumStudy;
		selection.piStudyCount = &iStudy;
		selection.piSeriesCount = &iSeries;
		hr = m_pDataList->GetSelection(&selection);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		m_iSelectedStudies = iStudy;

		if (1 == iStudy)
		{
			CComPtr<IVsiStudy> pStudy;
			hr = pEnumStudy->Next(1, &pStudy, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (! m_pItemPreview.IsEqualObject(pStudy))
			{
				m_pItemPreview = pStudy;

				// Preview study Note
				SetStudyInfo(pStudy);
			}
		}
		else
		{
			SetStudyInfo(NULL);
		}

		UpdateUI();
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiMoveSeries::OnFillList(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		SetStudyInfo(NULL);

		hr = FillDataList();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Change empty message
		hr = m_pDataList->SetEmptyMessage(L"No studies data available");
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiMoveSeries::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Move Series - OK");

	HRESULT hr = S_OK;

	try
	{
		CComPtr<IUnknown> pUnkItem;
		VSI_DATA_LIST_TYPE type(VSI_DATA_LIST_STUDY);

		int iFocused(-1);
		hr = m_pDataList->GetSelectedItem(&iFocused);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = m_pDataList->GetItem(iFocused, &type, &pUnkItem);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiStudy> pDesStudy(pUnkItem);
		VSI_CHECK_INTERFACE(pDesStudy, VSI_LOG_ERROR, NULL);

		// Check study lock state
		BOOL bStudyLocked = FALSE;
		// Get total lock state
		BOOL bLockAll = VsiGetBooleanValue<BOOL>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
			VSI_PARAMETER_SYS_LOCK_ALL, m_pPdm);
 
		// Check study lock state
		if (bLockAll)
		{
			CComVariant vLocked;
			hr = pDesStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
 
			if (VARIANT_TRUE == V_BOOL(&vLocked))
			{
				bStudyLocked = TRUE;
			}
		}

		if (bStudyLocked)
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
			// Close images
			{
				CComPtr<IVsiEnumImage> pEnumImage;
				hr = m_pSeries->GetImageEnumerator(
					VSI_PROP_IMAGE_MIN, TRUE, &pEnumImage, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					CComPtr<IVsiImage> pImageTest;
					while (pEnumImage->Next(1, &pImageTest, NULL) == S_OK)
					{
						CComVariant vNs;
						hr = pImageTest->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						int iSlot(-1);
						HRESULT hrLoaded = m_pSession->GetIsItemLoaded(V_BSTR(&vNs), &iSlot);
						if (S_OK == hrLoaded)
						{
							CComPtr<IVsiImage> pImage;
							hr = m_pSession->GetImage(iSlot, &pImage);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							CComVariant vParam(pImage.p);
							hr = m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_IMAGE, &vParam);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}

						pImageTest.Release();
					}
				}

				// Check any mode at all in session
				BOOL bMode(FALSE);
				for (int i = 0; i < VSI_SESSION_SLOT_MAX; ++i)
				{
					CComPtr<IUnknown> pUnkModeView;
					hr = m_pSession->GetModeView(i, &pUnkModeView);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						bMode = TRUE;
						break;
					}
				}

				if (!bMode)
				{
					hr = m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_MODE_VIEW, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}

			hr = m_pSeries->RemoveAllImages();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Get old study
			CComPtr<IVsiStudy> pStudyOld;
			hr = m_pSeries->GetStudy(&pStudyOld);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszResult;
			hr = m_pStudyMgr->MoveSeries(m_pSeries, pDesStudy, &pszResult);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Update session
			if (S_OK == hr)
			{
				CComVariant vNs;
				hr = m_pSeries->GetProperty(VSI_PROP_SERIES_NS, &vNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				int iSlot(-1);
				HRESULT hrLoaded = m_pSession->GetIsItemLoaded(V_BSTR(&vNs), &iSlot);
				if (S_OK == hrLoaded)
				{
					hr = m_pSession->SetStudy(iSlot, pDesStudy, FALSE);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				CComPtr<IVsiStudy> pStudyPrimary;
				hr = m_pSession->GetPrimaryStudy(&pStudyPrimary);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (pStudyOld.IsEqualObject(pStudyPrimary))
				{
					hr = m_pSession->SetPrimaryStudy(pDesStudy, TRUE);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				// Check if active series is moved
				CComPtr<IVsiSeries> pActiveSeries;
				hr = m_pSession->GetPrimarySeries(&pActiveSeries);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (pActiveSeries.IsEqualObject(m_pSeries))
				{
					CComHeapPtr<OLECHAR> pszSeries;
					hr = m_pSeries->GetDataPath(&pszSeries);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vSeries(pszSeries);

					CComVariant vSeriesNs;
					hr = m_pSeries->GetProperty(VSI_PROP_SERIES_NS, &vSeriesNs);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Update active series file is the active series is moved to a new study
					CComQIPtr<IVsiSessionJournal> pJournal(m_pSession);
					pJournal->Record(VSI_SESSION_EVENT_SERIES_NEW, &vSeriesNs, &vSeries);
				}
			}
			else
			{
				CString strMsg(MAKEINTRESOURCE(IDS_MOVESER_FAILED_MOVE_SELECTED_SERIES));
				if (NULL != pszResult)
				{
					strMsg += L" (";
					strMsg += pszResult;
					strMsg += L")";
				}
				VsiMessageBox(
					GetTopLevelParent(), 
					strMsg, 
					CString(MAKEINTRESOURCE(IDS_MOVESER_MOVE_SERIES)), 
					MB_OK | MB_ICONERROR);
			}
		}
	}
	VSI_CATCH(hr);

	m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_MOVE_SERIES, NULL);

	return hr;
}

LRESULT CVsiMoveSeries::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Move Series - Cancel");

	m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_MOVE_SERIES, NULL);

	return 0;
}

LRESULT CVsiMoveSeries::OnItemChanged(int, LPNMHDR, BOOL &bHandled)
{
	MSG msg;
	PeekMessage(&msg, *this, WM_VSI_UPDATE_PREVIEW, WM_VSI_UPDATE_PREVIEW, PM_REMOVE);
	PostMessage(WM_VSI_UPDATE_PREVIEW);

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiMoveSeries::OnListGetItemStatus(int, LPNMHDR pnmh, BOOL&)
{
	HRESULT hr = S_OK;
	LRESULT lRet = VSI_DATA_LIST_ITEM_STATUS_NORMAL;

	try
	{
		NM_DL_ITEM *pItem = (NM_DL_ITEM*)pnmh;
		if (VSI_DATA_LIST_STUDY == pItem->type)
		{
			VSI_CHECK_INTERFACE(m_pSeries, VSI_LOG_ERROR, NULL);

			CComQIPtr<IVsiStudy> pStudy(pItem->pUnkItem);

			CComVariant vCurrentNs;
			hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vCurrentNs);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IVsiStudy> pStudySrc;
			hr = m_pSeries->GetStudy(&pStudySrc);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vSrcNs;
			hr = pStudySrc->GetProperty(VSI_PROP_STUDY_NS, &vSrcNs);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (0 == lstrcmp(V_BSTR(&vSrcNs), V_BSTR(&vCurrentNs)))
			{
				lRet = VSI_DATA_LIST_ITEM_STATUS_HIDDEN;
			}
		}
	}
	VSI_CATCH(hr);

	return lRet;
}

void CVsiMoveSeries::UpdateUI()
{
	BOOL bEnableOK = (m_iSelectedStudies == 1);

	GetDlgItem(IDOK).EnableWindow(bEnableOK);
}

void CVsiMoveSeries::SetStudyInfo(IVsiStudy *pStudy)
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

		VSI_PROP_SERIES prop(VSI_PROP_SERIES_CREATED_DATE);

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

		m_studyOverview.SetStudy(
			pStudy,
			prop,
			(VSI_SYS_SORT_ORDER_DESCENDING == iOrder));
	}
	VSI_CATCH(hr);
}

HRESULT CVsiMoveSeries::FillDataList()
{
	HRESULT hr = S_OK;

	try
	{
		CComPtr<IVsiEnumStudy> pEnum;
		hr = m_pStudyMgr->GetStudyEnumerator(TRUE, &pEnum);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Display only the study level
		hr = m_pDataList->Fill(pEnum, 1);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

