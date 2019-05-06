/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiDataListWnd.cpp
**
**	Description:
**		Implementation of CVsiDataListWnd
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiWtl.h>
#include <VsiRes.h>
#include <VsiServiceProvider.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiThemeColor.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiOperatorModule.h>
#include <VsiStudyModule.h>
#include <VsiStudyXml.h>
#include <VsiServiceKey.h>
#include "VsiImportExportXml.h"
#include "VsiImpExpUtility.h"
#include "VsiDataListWnd.h"


#define VSI_COLOR_VISITED_R 140 
#define VSI_COLOR_VISITED_G 0 
#define VSI_COLOR_VISITED_B 140

#define VSI_STATE_IMAGE_WIDTH 13
#define VSI_SMALL_IMAGE_WIDTH 16


CVsiDataListWnd::CVsiDataListWnd() :
	m_dwFlags(0),
	m_wIdCtl(0),
	m_bTrackLeave(FALSE),
	m_iLastHotItem(-1),
	m_iLastHotState(-1),
	m_iLevelMax(2),
	m_bNotifyStateChanged(TRUE),
	m_typeRoot(VSI_DATA_LIST_NONE),
	m_hfontNormal(NULL),
	m_hfontReviewed(NULL)
{
}

CVsiDataListWnd::~CVsiDataListWnd()
{
	m_hfontNormal = NULL;  // Owned by outside
	m_hfontReviewed = NULL;  // Owned by outside

	_ASSERT(m_ilHeader.IsNull());
	_ASSERT(m_ilState.IsNull());
	_ASSERT(m_ilSmall.IsNull());
}

STDMETHODIMP CVsiDataListWnd::Initialize(
	HWND hwnd,
	DWORD dwFlags,
	VSI_DATA_LIST_TYPE typeRoot,
	IVsiApp *pApp,
	const VSI_LVCOLUMN *pColumns,
	int iColumns)
{
	HRESULT hr = S_OK;

	CCritSecLock csl(m_cs);

	try
	{
		m_dwFlags = dwFlags;

		baseClass::SubclassWindow(hwnd);

		m_wIdCtl = (WORD)GetWindowLong(GWL_ID);

		m_typeRoot = typeRoot;

		m_pApp = pApp;

		m_osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
		m_osvi.dwMajorVersion = 0;
		GetVersionEx(&m_osvi);

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Cache PDM
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Cache Study Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_STUDY_MGR,
			__uuidof(IVsiStudyManager),
			(IUnknown**)&m_pStudyMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Cache Operator Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_OPERATOR_MGR,
			__uuidof(IVsiOperatorManager),
			(IUnknown**)&m_pOperatorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Cache Mode Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_MODE_MANAGER,
			__uuidof(IVsiModeManager),
			(IUnknown**)&m_pModeManager);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Init list control
		ListView_SetExtendedListViewStyleEx(
			*this,
			LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP | LVS_EX_FULLROWSELECT | LVS_EX_SUBITEMIMAGES |
				LVS_EX_INFOTIP | LVS_EX_COLUMNSNAPPOINTS,
			LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP | LVS_EX_FULLROWSELECT | LVS_EX_SUBITEMIMAGES |
				LVS_EX_INFOTIP | LVS_EX_COLUMNSNAPPOINTS);
		ListView_SetHotCursor(*this, LoadCursor(NULL, IDC_ARROW));
		ListView_SetCallbackMask(*this, LVIS_FOCUSED | LVIS_SELECTED | LVIS_STATEIMAGEMASK);

		for (int i = 0; i < _countof(m_prgbState); ++i)
		{
			m_prgbState[i] = ListView_GetTextColor(*this);
		}

		COLORREF rgbNormal = m_prgbState[VSI_DATA_STATE_IMAGE_NORMAL];
		int iR = (GetRValue(rgbNormal) + VSI_COLOR_VISITED_R) / 2;
		int iG = (GetGValue(rgbNormal) + VSI_COLOR_VISITED_G) / 2;
		int iB = (GetBValue(rgbNormal) + VSI_COLOR_VISITED_B) / 2;
		m_prgbState[VSI_DATA_STATE_IMAGE_REVIEWED] = RGB(iR, iG, iB);
		m_prgbState[VSI_DATA_STATE_IMAGE_ERROR] = VSI_COLOR_TEXT_ERROR;
		m_prgbState[VSI_DATA_STATE_IMAGE_WARNING] = VSI_COLOR_TEXT_WARNING;
		m_prgbState[VSI_DATA_STATE_IMAGE_FILTER] = VSI_COLOR_TEXT_FILTER;

		m_ilState.CreateFromImage(IDB_LIST_STATE, VSI_STATE_IMAGE_WIDTH, 1, CLR_DEFAULT,
			IMAGE_BITMAP, LR_CREATEDIBSECTION);

		HIMAGELIST hilOld = ListView_SetImageList(*this, m_ilState, LVSIL_STATE);
		if (NULL != hilOld)
		{
			ImageList_Destroy(hilOld);
		}

		m_ilSmall.CreateFromImage(IDB_STUDY_LIST, VSI_SMALL_IMAGE_WIDTH, 1, CLR_DEFAULT,
			IMAGE_BITMAP, LR_CREATEDIBSECTION);

		hilOld = ListView_SetImageList(*this, m_ilSmall, LVSIL_SMALL);
		if (NULL != hilOld)
		{
			ImageList_Destroy(hilOld);
		}

		// Header
		CWindow wndHeader(ListView_GetHeader(*this));

		m_ilHeader.CreateFromImage(IDB_SB_HEADER, 9, 1, CLR_NONE, IMAGE_BITMAP, LR_CREATEDIBSECTION);

		Header_SetImageList(wndHeader, m_ilHeader.operator HIMAGELIST());
		Header_SetBitmapMargin(wndHeader, 2);

		// Add columns
		hr = SetColumns(pColumns, iColumns);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		UpdateSortIndicator(TRUE);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::Uninitialize()
{
	CCritSecLock csl(m_cs);

	baseClass::UnsubclassWindow();

	Clear();

	m_ilHeader.Destroy();

	if (!m_ilState.IsNull())
	{
		ListView_SetImageList(*this, NULL, LVSIL_STATE);
		m_ilState.Destroy();
	}

	if (!m_ilSmall.IsNull())
	{
		ListView_SetImageList(*this, NULL, LVSIL_SMALL);
		m_ilSmall.Destroy();
	}

	m_pModeManager.Release();
	m_pStudyMgr.Release();
	m_pPdm.Release();
	m_pApp.Release();

	return S_OK;
}

STDMETHODIMP CVsiDataListWnd::LockData()
{
	return m_cs.TryEnter() ? S_OK : S_FALSE;
}

STDMETHODIMP CVsiDataListWnd::UnlockData()
{
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CVsiDataListWnd::SetColumns(
	const VSI_LVCOLUMN *pColumns, int iColumns)
{
	CWindow wndHeader(ListView_GetHeader(*this));

	// Clear old
	{
		int iCount = Header_GetItemCount(wndHeader);

		for (int i = 0; i < iCount; ++i)
		{
			ListView_DeleteColumn(*this, 0);
		}

		m_mapIndexToCol.clear();
		m_mapColToIndex.clear();
	}

	// Add columns
	std::unique_ptr<int[]> piOrder(new int[iColumns]);

	LVCOLUMN col = { 0 };

	int iColumnSet = 0;
	for (int i = 0; i < iColumns; ++i)
	{
		if (pColumns[i].bVisible)
		{
			CString strText(MAKEINTRESOURCE(pColumns[i].dwTextId));

			col.mask = pColumns[i].mask | LVCF_SUBITEM | LVCF_MINWIDTH;
			col.pszText = (LPWSTR)strText.GetBuffer();
			col.iSubItem = pColumns[i].iSubItem;
			col.cx = pColumns[i].cx;
			col.fmt = pColumns[i].fmt;
			col.cxMin = 64;
			ListView_InsertColumn(*this, iColumnSet, &col);

			// Set order array, added extra error checking
			if (pColumns[i].iOrder < iColumns)
			{
				piOrder[pColumns[i].iOrder] = iColumnSet;
			}

			if (0 <= pColumns[i].iImage)
			{
				HDITEM hdi = { HDI_FORMAT | HDI_IMAGE };
				hdi.fmt = HDF_LEFT | HDF_IMAGE;
				hdi.iImage = pColumns[i].iImage;
				Header_SetItem(wndHeader, iColumnSet, &hdi);
			}

			m_mapIndexToCol[iColumnSet] = (VSI_DATA_LIST_COL)pColumns[i].iSubItem;
			m_mapColToIndex[(VSI_DATA_LIST_COL)pColumns[i].iSubItem] = iColumnSet;

			++iColumnSet;
		}
	}

	ListView_SetColumnOrderArray(*this, iColumnSet, piOrder.get());

	UpdateSortIndicator(TRUE);

	return S_OK;
}

STDMETHODIMP CVsiDataListWnd::SetTextColor(
	VSI_DATA_LIST_ITEM_STATUS dwState, COLORREF rgbState)
{
	switch (dwState)
	{
	case VSI_DATA_LIST_ITEM_STATUS_ERROR:
		m_prgbState[VSI_DATA_STATE_IMAGE_ERROR] = rgbState;
		break;
	case VSI_DATA_LIST_ITEM_STATUS_WARNING:
		m_prgbState[VSI_DATA_STATE_IMAGE_WARNING] = rgbState;
		break;
	case VSI_DATA_LIST_ITEM_STATUS_ACTIVE_SINGLE:
		m_prgbState[VSI_DATA_STATE_IMAGE_SINGLE] = rgbState;
		break;
	case VSI_DATA_LIST_ITEM_STATUS_ACTIVE_LEFT:
		m_prgbState[VSI_DATA_STATE_IMAGE_LEFT] = rgbState;
		break;
	case VSI_DATA_LIST_ITEM_STATUS_ACTIVE_RIGHT:
		m_prgbState[VSI_DATA_STATE_IMAGE_RIGHT] = rgbState;
		break;
	case VSI_DATA_LIST_ITEM_STATUS_ACTIVE:
		{
			int iR = (GetRValue(rgbState) + VSI_COLOR_VISITED_R) / 2;
			int iG = (GetGValue(rgbState) + VSI_COLOR_VISITED_G) / 2;
			int iB = (GetBValue(rgbState) + VSI_COLOR_VISITED_B) / 2;
			m_prgbState[VSI_DATA_STATE_IMAGE_ACTIVE_NORMAL] = rgbState;
			m_prgbState[VSI_DATA_STATE_IMAGE_ACTIVE_REVIEWED] = RGB(iR, iG, iB);
		}
		break;
	case VSI_DATA_LIST_ITEM_STATUS_INACTIVE:
		{
			int iR = (GetRValue(rgbState) + VSI_COLOR_VISITED_R) / 2;
			int iG = (GetGValue(rgbState) + VSI_COLOR_VISITED_G) / 2;
			int iB = (GetBValue(rgbState) + VSI_COLOR_VISITED_B) / 2;
			m_prgbState[VSI_DATA_STATE_IMAGE_INACTIVE_NORMAL] = rgbState;
			m_prgbState[VSI_DATA_STATE_IMAGE_INACTIVE_REVIEWED] = RGB(iR, iG, iB);
		}
		break;
	default:
		_ASSERT(0);
	}

	return S_OK;
}

STDMETHODIMP CVsiDataListWnd::SetFont(HFONT hfontNormal, HFONT hfontReviewed)
{
	m_hfontNormal = hfontNormal;
	m_hfontReviewed = hfontReviewed;

	return S_OK;
}

STDMETHODIMP CVsiDataListWnd::SetEmptyMessage(LPCOLESTR pszMessage)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pszMessage, VSI_LOG_ERROR, NULL);

		m_strEmptyMsg = pszMessage;

		if (0 == m_vecItem.size())
		{
			Invalidate();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::GetSortSettings(VSI_DATA_LIST_TYPE type, VSI_DATA_LIST_COL *pCol, BOOL *pbDescending)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pCol, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(pbDescending, VSI_LOG_ERROR, NULL);

		switch (type)
		{
		case VSI_DATA_LIST_STUDY:
			*pCol = m_sort.m_sortStudyCol;
			*pbDescending = m_sort.m_bStudyDescending;
			break;
		case VSI_DATA_LIST_SERIES:
			*pCol = m_sort.m_sortSeriesCol;
			*pbDescending = m_sort.m_bSeriesDescending;
			break;
		default:
			*pCol = m_sort.m_sortImageCol;
			*pbDescending = m_sort.m_bImageDescending;
			break;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::SetSortSettings(VSI_DATA_LIST_TYPE type, VSI_DATA_LIST_COL col, BOOL bDescending)
{
	HRESULT hr = S_OK;

	UpdateSortIndicator(FALSE);

	m_sort.m_sortCol = col;
	m_sort.m_bDescending = bDescending;

	switch (type)
	{
	case VSI_DATA_LIST_STUDY:
		m_sort.m_sortStudyCol = col;
		m_sort.m_bStudyDescending = bDescending;
		break;
	case VSI_DATA_LIST_SERIES:
		m_sort.m_sortSeriesCol = col;
		m_sort.m_bSeriesDescending = bDescending;
		break;
	default:
		m_sort.m_sortImageCol = col;
		m_sort.m_bImageDescending = bDescending;
		break;
	}

	UpdateSortIndicator(TRUE);
	UpdateList();

	return hr;
}

STDMETHODIMP CVsiDataListWnd::Clear()
{
	// Clear list
	if (IsWindow())
	{
		ListView_SetItemCount(*this, 0);
	}

	m_vecItem.clear();

	m_listRoot.clear();

	return S_OK;
}

STDMETHODIMP CVsiDataListWnd::Fill(IUnknown *pUnkEnum, int iLevelMax)
{
	CComQIPtr<IVsiEnumStudy> pEnumStudy(pUnkEnum);
	if (NULL != pEnumStudy)
	{
		CCritSecLock csl(m_cs);

		Fill(pEnumStudy, iLevelMax);
		return S_OK;
	}

	CComQIPtr<IVsiEnumSeries> pEnumSeries(pUnkEnum);
	if (NULL != pEnumSeries)
	{
		CCritSecLock csl(m_cs);

		Fill(pEnumSeries, iLevelMax);
		return S_OK;
	}

	CComQIPtr<IVsiEnumImage> pEnumImage(pUnkEnum);
	if (NULL != pEnumImage)
	{
		CCritSecLock csl(m_cs);

		Fill(pEnumImage);
		return S_OK;
	}

	return S_OK;
}

STDMETHODIMP CVsiDataListWnd::Filter(LPCWSTR pszSearch, BOOL *pbStop)
{
	HRESULT hr = S_OK;

	try
	{
		CVsiFilterJob job;

		if (0 != *pszSearch)
		{
			// Decode filter string
			int iLen = wcslen(pszSearch);
			bool bQuote(false);

			LPCWSTR pszBegin = pszSearch;
			for (int i = 0; i < iLen; ++i)
			{
				switch (pszSearch[i])
				{
				case L'\"':
					if (bQuote)
					{
						LPCWSTR pszEnd = pszSearch + i - 1;
						if (pszEnd >= pszBegin)
						{
							CString str(pszBegin, pszEnd - pszBegin + 1);
							str.Trim();

							if (0 == str.Compare(L"OR"))
							{
								if (VSI_FILTER_TYPE_NONE == job.m_filterType)
								{
									job.m_filterType = VSI_FILTER_TYPE_OR;
								}
							}
							else if (0 == str.Compare(L"AND"))
							{
								if (VSI_FILTER_TYPE_NONE == job.m_filterType)
								{
									job.m_filterType = VSI_FILTER_TYPE_AND;
								}
							}
							else
							{
								job.m_vecFilters.push_back(str);
							}
						}

						pszBegin = pszSearch + i + 1;
					}
					else
					{
						pszBegin = pszSearch + i + 1;  // Skip "
					}
					bQuote = !bQuote;
					break;
				case L' ':
					if (!bQuote)
					{
						LPCWSTR pszEnd = pszSearch + i - 1;
						if (pszEnd >= pszBegin)
						{
							CString str(pszBegin, pszEnd - pszBegin + 1);
							str.Trim();

							if (0 == str.Compare(L"OR"))
							{
								if (VSI_FILTER_TYPE_NONE == job.m_filterType)
								{
									job.m_filterType = VSI_FILTER_TYPE_OR;
								}
							}
							else if (0 == str.Compare(L"AND"))
							{
								if (VSI_FILTER_TYPE_NONE == job.m_filterType)
								{
									job.m_filterType = VSI_FILTER_TYPE_AND;
								}
							}
							else
							{
								job.m_vecFilters.push_back(str);
							}
						}

						pszBegin = pszSearch + i + 1;
					}
					break;
				}
			}

			if (0 != *pszBegin)
			{
				LPCWSTR pszEnd = pszSearch + iLen;
				if (pszEnd >= pszBegin)
				{
					CString str(pszBegin, pszEnd - pszBegin + 1);
					str.Trim();

					if (0 == str.Compare(L"OR"))
					{
						if (VSI_FILTER_TYPE_NONE == job.m_filterType)
						{
							job.m_filterType = VSI_FILTER_TYPE_OR;
						}
					}
					else if (0 == str.Compare(L"AND"))
					{
						if (VSI_FILTER_TYPE_NONE == job.m_filterType)
						{
							job.m_filterType = VSI_FILTER_TYPE_AND;
						}
					}
					else
					{
						job.m_vecFilters.push_back(str);
					}
				}
			}

			if (VSI_FILTER_TYPE_NONE == job.m_filterType)
			{
				// Default is AND
				job.m_filterType = VSI_FILTER_TYPE_AND;
			}
		}

		bool bNewJob(false);
		if (0 == m_vecFilterJobs.size())
		{
			// 1st job
			bNewJob = true;
		}
		else
		{
			// The filters are different?
			auto const &lastJob = m_vecFilterJobs.back();
			if (lastJob.m_filterType != job.m_filterType ||
				lastJob.m_vecFilters != job.m_vecFilters)
			{
				// A different job
				bNewJob = true;
			}
		}

		if (bNewJob)
		{
			// New search

			ListView_SetItemCountEx(*this, 0, LVSICF_NOSCROLL);

			{
				CCritSecLock csl(m_cs);
				m_vecItem.clear();

				if ((0 == m_vecFilterJobs.size()) && (0 < job.m_vecFilters.size()))
				{
					// 1st - expand all
					ExpandAll(m_listRoot, *pbStop);
				}

				if (!(*pbStop))
				{
					GetNextFilterJob(job);

					Filter(m_listRoot, job, pbStop);

					if (!(*pbStop))
					{
						m_vecFilterJobs.clear();

						if (0 < job.m_vecFilters.size())
						{
							job.m_pItemFilterStop = nullptr;
							job.m_scope = VSI_SEARCH_ALL;
							m_vecFilterJobs.push_back(job);
						}

						UpdateItemStatus(LPCWSTR(NULL), TRUE);

						UpdateList();
					}
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::GetItemCount(int *piCount)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(piCount, VSI_LOG_ERROR, NULL);

		*piCount = (int)m_vecItem.size();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::GetSelectedItem(int *piIndex)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(piIndex, VSI_LOG_ERROR, NULL);

		*piIndex = -1;

		int iLast = (int)m_vecItem.size();

		// Find focused
		for (int i = 0; i < iLast; ++i)
		{
			CVsiItem *pItem = m_vecItem[i];

			if (pItem->m_bFocused)
			{
				*piIndex = i;
				return S_OK;
			}
		}

		// Find selected
		for (int i = 0; i < iLast; ++i)
		{
			CVsiItem *pItem = m_vecItem[i];

			if (pItem->m_bSelected)
			{
				*piIndex = i;
				return S_OK;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::SetSelectedItem(int iIndex, BOOL bEnsureVisible)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(iIndex < (int)m_vecItem.size(), VSI_LOG_ERROR, NULL);

		CVsiItem *pItem = m_vecItem[iIndex];

		ListView_SetItemState(*this, iIndex, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
		if (bEnsureVisible)
		{
			ListView_EnsureVisible(*this, iIndex, FALSE);
		}

		// Selects all children if item is opened
		if (pItem->m_bExpanded)
		{
			int i = iIndex;
			for (CVsiListItem::size_type j = 0; j < pItem->GetVisibleChildrenCount(); ++j)
			{
				ListView_SetItemState(*this, ++i, LVIS_SELECTED, LVIS_SELECTED);
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::SelectAll()
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		LV_ITEM lvi;
		lvi.stateMask = LVIS_SELECTED;
		lvi.state = LVIS_SELECTED;
		
		int iLast = (int)m_vecItem.size();
		for (int i = 0; i < iLast; ++i)
		{
			CVsiItem *pItem = m_vecItem[i];

			if (!pItem->m_bSelected)
			{
				SendMessage(LVM_SETITEMSTATE, (WPARAM)i, (LPARAM)&lvi);
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::GetSelection(VSI_DATA_LIST_COLLECTION *pSelection)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	CVsiListItem listStudySelected;
	CVsiListItem listSeriesSelected;
	CVsiListItem listImageSelected;

	try
	{
		VSI_CHECK_ARG(pSelection, VSI_LOG_ERROR, NULL);

		CVsiItem *pItem = NULL;

		// Generate new selection lists
		for (auto iter = m_vecItem.begin(); iter != m_vecItem.end(); ++iter)
		{
			pItem = *iter;
			if (VSI_DATA_LIST_STUDY == pItem->m_type)
			{
				if (pItem->m_bSelected)
				{
					if (VSI_DATA_LIST_FILTER_SKIP_ERRORS_AND_WARNINGS & pSelection->dwFlags)
					{
						CVsiStudy* pStudy = dynamic_cast<CVsiStudy*>(pItem);
						VSI_VERIFY(nullptr != pStudy, VSI_LOG_ERROR, NULL);

						VSI_STUDY_ERROR dwErrorCode(VSI_STUDY_ERROR_NONE);
						pStudy->m_pStudy->GetErrorCode(&dwErrorCode);
						if (VSI_STUDY_ERROR_NONE != dwErrorCode)
						{
							continue;
						}
					}

					listStudySelected.push_back(pItem);
				}
			}
		}
		for (auto iter = m_vecItem.begin(); iter != m_vecItem.end(); ++iter)
		{
			pItem = *iter;
			if (VSI_DATA_LIST_SERIES == pItem->m_type)
			{
				if (pItem->m_bSelected)
				{
					if (VSI_DATA_LIST_FILTER_SKIP_ERRORS_AND_WARNINGS & pSelection->dwFlags)
					{
						CVsiSeries* pSeries = dynamic_cast<CVsiSeries*>(pItem);
						VSI_VERIFY(nullptr != pSeries, VSI_LOG_ERROR, NULL);

						VSI_SERIES_ERROR dwErrorCode(VSI_SERIES_ERROR_NONE);
						pSeries->m_pSeries->GetErrorCode(&dwErrorCode);
						if (VSI_SERIES_ERROR_NONE != dwErrorCode)
						{
							continue;
						}
					}

					listSeriesSelected.push_back(pItem);
				}
			}
		}
		for (auto iter = m_vecItem.begin(); iter != m_vecItem.end(); ++iter)
		{
			pItem = *iter;
			if (VSI_DATA_LIST_IMAGE == pItem->m_type)
			{
				if (pItem->m_bSelected)
				{
					if (VSI_DATA_LIST_FILTER_SKIP_ERRORS_AND_WARNINGS & pSelection->dwFlags)
					{
						CVsiImage* pImage = dynamic_cast<CVsiImage*>(pItem);
						VSI_VERIFY(nullptr != pImage, VSI_LOG_ERROR, NULL);

						VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_NONE);
						pImage->m_pImage->GetErrorCode(&dwErrorCode);
						if (VSI_IMAGE_ERROR_NONE != dwErrorCode)
						{
							continue;
						}
					}

					listImageSelected.push_back(pItem);

					if (VSI_DATA_LIST_FILTER_SELECT_ALL_IMAGES & pSelection->dwFlags)
					{
						listSeriesSelected.push_back(pItem->m_pParent);
					}
				}
			}
		}

		switch (pSelection->dwFlags & VSI_DATA_LIST_FILTER_BASE_TYPE_MASK)
		{
		case VSI_DATA_LIST_FILTER_NONE:
			// No filtering
			break;
		case VSI_DATA_LIST_FILTER_SELECT_FOR_DELETE:
			{
				// This case only removes active study from the list and 
				// falls through to VSI_DATA_LIST_FILTER_SELECT_PARENT
				hr = AdjustForActiveSeries(
					listStudySelected,
					listSeriesSelected, 
					listImageSelected,
					pSelection->pSeriesActive);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}  // fall through
		case VSI_DATA_LIST_FILTER_SELECT_PARENT:
			{
				// Flip the selection to the parent
				hr = RemoveDuplicateChildren(listSeriesSelected, listImageSelected);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = RemoveDuplicateChildren(listStudySelected, listSeriesSelected);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			break;
		case VSI_DATA_LIST_FILTER_SELECT_CHILDREN:
			{
				// Flip the selection to children
				hr = RemoveDuplicateParent(listStudySelected, listSeriesSelected);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = RemoveDuplicateParent(listSeriesSelected, listImageSelected);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			break;
		case VSI_DATA_LIST_FILTER_MOVE_TO_CHILDREN:
			{
				// This case is handles differently because selection is moved to the
				// children and the children may not be even populated yet

				// Move parent to children
				hr = MoveToChildren(listStudySelected, listSeriesSelected);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = MoveToChildren(listSeriesSelected, listImageSelected);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			break;
		case VSI_DATA_LIST_FILTER_SELECT_ALL_IMAGES:
			{
				// This case is handles differently because selection is moved to the
				// children and the children may not be even populated yet

				//Move parent to children - we have series from studies
				hr = MoveToChildren(listStudySelected, listSeriesSelected, false);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = RemoveDuplicates(listSeriesSelected);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				listImageSelected.empty();

				hr = MoveToChildren(listSeriesSelected, listImageSelected, false);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			break;
		}

		// Gets enumerated lists
		if (pSelection->ppEnumStudy != NULL)
		{
			hr = GetSelectedStudies(
				pSelection->ppEnumStudy, pSelection->piStudyCount, &listStudySelected);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (pSelection->piStudyCount != NULL)
		{
			*pSelection->piStudyCount = (int)listStudySelected.size();
		}

		if (pSelection->ppEnumSeries != NULL)
		{
			hr = GetSelectedSeries(
				pSelection->ppEnumSeries, pSelection->piSeriesCount, &listSeriesSelected);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (pSelection->piSeriesCount != NULL)
		{
			*pSelection->piSeriesCount = (int)listSeriesSelected.size();
		}

		if (pSelection->ppEnumImage != NULL)
		{
			hr = GetSelectedImages(
				pSelection->ppEnumImage, pSelection->piImageCount, &listImageSelected);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (pSelection->piImageCount != NULL)
		{
			*pSelection->piImageCount = (int)listImageSelected.size();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::SetSelection(LPCOLESTR pszNamespace, BOOL bEnsureVisible)
{
	HRESULT hr = S_FALSE;  // Namespace not found

	CCritSecLock csl(m_cs);

	// Remove all selection
	ListView_SetItemState(*this, -1, 0, LVIS_FOCUSED | LVIS_SELECTED);

	LPCWSTR pszNext = pszNamespace;
	CString strId;
	CVsiListItem *pList = &m_listRoot;	

	while ((NULL != pszNext) && (0 != *pszNext))
	{
		LPCWSTR pszSpt = wcschr(pszNext, L'/');
		if (pszSpt != NULL)
		{
			strId.SetString(pszNext, (int)(pszSpt - pszNext));
			pszNext = pszSpt + 1;
		}
		else
		{
			strId = pszNext;
			pszNext = NULL;
		}

		for (auto iter = pList->begin(); iter != pList->end(); ++iter)
		{
			CVsiItem *pItem = *iter;
			CComHeapPtr<OLECHAR> pszId;
			pItem->GetId(&pszId);

			if (pszId != NULL && strId == pszId)
			{
				if (NULL == pszNext)
				{
					// Reach the end
					pItem->m_bSelected = TRUE;
					pItem->m_bFocused = TRUE;
				}
				else if (0 == *pszNext)
				{
					// Ended with '/'
					pItem->m_bSelected = TRUE;
					pItem->m_bFocused = TRUE;
					
					pItem->Expand(TRUE);
				}
				else
				{
					// Found a node, expand it
					pItem->Expand();

					pList = &pItem->m_listChild;
				}
				
				hr = S_OK;  // Done
				break;
			}
		}
	}

	UpdateList();

	if (bEnsureVisible)
	{
		int iSel = ListView_GetNextItem(*this, -1, LVNI_SELECTED);
		if (iSel >= 0)
		{
			ListView_EnsureVisible(*this, iSel, FALSE);
		}
	}

	return hr;
}

STDMETHODIMP CVsiDataListWnd::GetLastSelectedItem(
	VSI_DATA_LIST_TYPE *pType,
	IUnknown **ppUnk)
{
	HRESULT hr = S_FALSE;
	CVsiItem *pItem = NULL;

	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pType, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(ppUnk, VSI_LOG_ERROR, NULL);
		
		*pType = VSI_DATA_LIST_NONE;
		*ppUnk = NULL;

		int iIndex = -1;

		int iLast = (int)m_vecItem.size();

		for (int i = 0; i < iLast; ++i)
		{
			pItem = m_vecItem[i];

			if (pItem->m_bFocused)
			{
				iIndex = i;
				break;
			}
		}

		if (NULL != pItem)
		{
			if (VSI_DATA_LIST_STUDY == pItem->m_type)
			{
				CVsiStudy* pStudy = dynamic_cast<CVsiStudy*>(pItem);
				VSI_VERIFY(nullptr != pStudy, VSI_LOG_ERROR, NULL);

				*ppUnk = pStudy->m_pStudy;
			}
			else if (VSI_DATA_LIST_SERIES == pItem->m_type)
			{
				CVsiSeries* pSeries = dynamic_cast<CVsiSeries*>(pItem);
				VSI_VERIFY(nullptr != pSeries, VSI_LOG_ERROR, NULL);

				*ppUnk = pSeries->m_pSeries;
			}
			else if (VSI_DATA_LIST_IMAGE == pItem->m_type)
			{
				CVsiImage* pImage = dynamic_cast<CVsiImage*>(pItem);
				VSI_VERIFY(nullptr != pImage, VSI_LOG_ERROR, NULL);

				*ppUnk = pImage->m_pImage;
			}

			if (*ppUnk != NULL)
			{
				(*ppUnk)->AddRef();
				*pType = pItem->m_type;

				hr = S_OK;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::UpdateItem(
	LPCOLESTR pszNamespace, BOOL bRecurChildren)
{
	CCritSecLock csl(m_cs);

	if (NULL == pszNamespace)
	{
		for (auto iter = m_listRoot.begin(); iter != m_listRoot.end(); ++iter)
		{
			CVsiItem *pItem = *iter;
			if (pItem->m_bExpanded)
			{
				pItem->Collapse();
				pItem->Expand();
			}
			UpdateItemStatus(pItem, bRecurChildren);
		}

		UpdateList();
	}
	else
	{
		CVsiItem *pItem = GetItem(pszNamespace, true, NULL);

		if (pItem != NULL)
		{
			if (pItem->m_bExpanded)
			{
				pItem->Collapse();
				pItem->Expand();
			}

			UpdateItemStatus(pItem, bRecurChildren);

			UpdateList();
		}
	}

	return S_OK;
}

STDMETHODIMP CVsiDataListWnd::UpdateItemStatus(
	LPCOLESTR pszNamespace, BOOL bRecurChildren)
{
	CCritSecLock csl(m_cs);

	if (NULL == pszNamespace)
	{
		for (auto iter = m_listRoot.begin(); iter != m_listRoot.end(); ++iter)
		{
			CVsiItem *pItem = *iter;
			UpdateItemStatus(pItem, bRecurChildren);
		}
	}
	else
	{
		CVsiItem *pItem = GetItem(pszNamespace, true, NULL);

		if (pItem != NULL)
		{
			UpdateItemStatus(pItem, bRecurChildren);
		}
	}

	return S_OK;
}

STDMETHODIMP CVsiDataListWnd::UpdateList(IXMLDOMDocument *pXmlDoc)
{
	HRESULT hr = S_OK;
	CVsiItem *pNewSelItem = NULL;
	BOOL bDeletedItems = FALSE;
	BOOL bAddSelectedItems = FALSE;
	BOOL bRefreshSelectedItems = FALSE;

	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pXmlDoc, VSI_LOG_ERROR, NULL);

		// Update XML example
		// <update resetSelection="true">
		//   <add>
		//		<item ns="1234" path="c:/study/1234567"/>
		//   </add>
		//   <remove>
		//      <item ns="3456/4567/5678"/>
		//   </remove>
		//   <refresh>
		//      <item ns="6789"/>
		//   </refresh>
		// </update>

		CVsiMapItemLocation itemLocation;

		// Global flags
		CComVariant vResetSel(false);
		{
			CComPtr<IXMLDOMElement> pXmlElemRoot;
			hr = pXmlDoc->get_documentElement(&pXmlElemRoot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
			hr = pXmlElemRoot->getAttribute(CComBSTR(VSI_UPDATE_XML_ATT_RESET_SELECTION), &vResetSel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
			if (S_OK == hr)
			{
				hr = vResetSel.ChangeType(VT_BOOL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else
			{
				vResetSel = false;
			}
		}

		// Remove all selection?
		if (VARIANT_TRUE == V_BOOL(&vResetSel))
		{
			ListView_SetItemState(*this, -1, 0, LVIS_FOCUSED | LVIS_SELECTED);
		}

		// Remove
		{
			// Get the number of deleted studies/series/images
			int iNumStudy = 0, iNumSeries = 0, iNumImage = 0;
			DWORD dwError = 0;
			CComVariant vNum;
			CString strPrevStudy, strPrevSeries;
			BOOL bMultipleStudies = FALSE;
			BOOL bMultipleSeries = FALSE;
			BOOL bFindNext = FALSE;
			BOOL bSelectParent = FALSE;

			CComPtr<IXMLDOMNode> pItemRemove;
			hr = pXmlDoc->selectSingleNode(
				CComBSTR(L"//" VSI_UPDATE_XML_ELM_UPDATE L"/"
					VSI_UPDATE_XML_ELM_REMOVE),
				&pItemRemove);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr)
			{
				CComQIPtr<IXMLDOMElement> pElemRemove(pItemRemove);

				vNum.Clear();
				hr = pElemRemove->getAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_STUDIES), &vNum);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_REMOVE_XML_ATT_NUMBER_STUDIES));
				if (S_OK == hr)
				{
					VsiVariantDecodeValue(vNum, VT_I4);
					iNumStudy = V_I4(&vNum);
				}

				vNum.Clear();
				hr = pElemRemove->getAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_SERIES), &vNum);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_REMOVE_XML_ATT_NUMBER_SERIES));
				if (S_OK == hr)
				{
					VsiVariantDecodeValue(vNum, VT_I4);
					iNumSeries = V_I4(&vNum);
				}

				vNum.Clear();
				hr = pElemRemove->getAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_IMAGES), &vNum);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_REMOVE_XML_ATT_NUMBER_IMAGES));
				if (S_OK == hr)
				{
					VsiVariantDecodeValue(vNum, VT_I4);
					iNumImage = V_I4(&vNum);
				}

				long lCount = 0;

				if (iNumStudy + iNumSeries + iNumImage > 0)
				{
					CComVariant vError;
					hr = pElemRemove->getAttribute(CComBSTR(VSI_REMOVE_XML_ATT_ERROR), &vError);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure getting attribute (%s)", VSI_REMOVE_XML_ATT_ERROR));
					if (S_OK == hr)
					{
						hr = vError.ChangeType(VT_UI4);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						dwError = V_UI4(&vError);
					}

					CComPtr<IXMLDOMNodeList> pListItemRemove;
					hr = pXmlDoc->selectNodes(
						CComBSTR(L"//" VSI_UPDATE_XML_ELM_UPDATE L"/"
							VSI_UPDATE_XML_ELM_REMOVE L"/" VSI_UPDATE_XML_ELM_ITEM),
						&pListItemRemove);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pListItemRemove->get_length(&lCount);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

					if (lCount > 0)
					{
						itemLocation.clear();
						UpdateItemLocationMap(itemLocation, L"", m_listRoot);

						CComVariant vNs;
						for (long l = 0; l < lCount; ++l)
						{
							CComPtr<IXMLDOMNode> pNode;
							hr = pListItemRemove->get_item(l, &pNode);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

							CComQIPtr<IXMLDOMElement> pElemParam(pNode);

							vNs.Clear();
							hr = pElemParam->getAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), &vNs);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR,
								VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_ATT_NS));

							LPCWSTR pszNs = V_BSTR(&vNs);

							if (VSI_DATA_LIST_STUDY == m_typeRoot)
							{
								NULL;
							}
							else if (VSI_DATA_LIST_SERIES == m_typeRoot)
							{
								pszNs = wcschr(pszNs, L'/');
								VSI_VERIFY(pszNs != NULL, VSI_LOG_ERROR, NULL);
								++pszNs;  // Skip '/'
							}
							else if (VSI_DATA_LIST_IMAGE == m_typeRoot)
							{
								pszNs = wcschr(pszNs, L'/');
								VSI_VERIFY(pszNs != NULL, VSI_LOG_ERROR, NULL);
								++pszNs;  // Skip '/'
								pszNs = wcschr(pszNs, L'/');
								VSI_VERIFY(pszNs != NULL, VSI_LOG_ERROR, NULL);
								++pszNs;  // Skip '/'
							}
							else
							{
								_ASSERT(0);
							}

							CString strNs(pszNs);
							auto iter = itemLocation.find(strNs);
							if (iter != itemLocation.end())
							{
								CVsiItemLocation &loc = iter->second;

								// Special case - 1 item - find next/previous item
								if (0 == dwError && 1 == lCount)
								{
									// Find next peer item and select it
									CVsiItem* pItem = *loc.m_iter;
									VSI_DATA_LIST_TYPE type = pItem->m_type;

									int iItem = -1;
									for (int i = 0; i < (int)m_vecItem.size(); ++i)
									{
										if (pItem == m_vecItem[i])
										{
											iItem = i;
											break;
										}
									}

									if (iItem >= 0)
									{
										BOOL bNextItemFound = FALSE;
										int iCount = (int)m_vecItem.size();
										CVsiItem* pParent = m_vecItem[iItem]->m_pParent;

										// Walk the list forward
										int iCurrItem = iItem + 1;
										while (iCurrItem < iCount)
										{
											CVsiItem *pCurrItem = m_vecItem[iCurrItem];

											if (type == pCurrItem->m_type &&
												pParent == m_vecItem[iCurrItem]->m_pParent)
											{
												bNextItemFound = TRUE;
												pNewSelItem = pCurrItem;
												break;
											}

											++iCurrItem;
										}
									
										if (!bNextItemFound)
										{
											// Walk the list backward
											iCurrItem = iItem - 1;
											while (iCurrItem >= 0)
											{
												CVsiItem *pCurrItem = m_vecItem[iCurrItem];

												if (type == pCurrItem->m_type &&
													pParent == m_vecItem[iCurrItem]->m_pParent)
												{
													bNextItemFound = TRUE;
													pNewSelItem = pCurrItem;
													break;
												}

												--iCurrItem;
											}
										}

										if (!bNextItemFound && iItem > 0)
										{
											// Select parent
											bNextItemFound = TRUE;
											pNewSelItem = m_vecItem[iItem - 1];
										}
									}
								}

								// Remove item from corresponding maps
								CVsiItem *pItem = *(loc.m_iter);

								if (VSI_DATA_LIST_SERIES == pItem->m_type)
								{
									CComVariant vId;
									CVsiSeries *pSeries = static_cast<CVsiSeries*>(*(loc.m_iter));
									hr = pSeries->m_pSeries->GetProperty(VSI_PROP_SERIES_ID, &vId);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									// Remove series from the map of series in the study
									CVsiStudy *pStudy = static_cast<CVsiStudy*>(pItem->m_pParent);
									hr = pStudy->m_pStudy->RemoveSeries(V_BSTR(&vId));
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									if (0 < m_vecFilterJobs.size())
									{
										pSeries->m_bVisible = FALSE;

										auto & job = m_vecFilterJobs.back();
										FilterParent(pStudy, job.m_vecFilters, job.m_filterType);
									}
								}
								else if (VSI_DATA_LIST_IMAGE == pItem->m_type)
								{
									CComVariant vId;
									CVsiImage *pImage = static_cast<CVsiImage*>(*(loc.m_iter));
									hr = pImage->m_pImage->GetProperty(VSI_PROP_IMAGE_ID, &vId);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									// Remove image from the map of images in the series
									CVsiSeries *pSeries = static_cast<CVsiSeries*>(pItem->m_pParent);
									hr = pSeries->m_pSeries->RemoveImage(V_BSTR(&vId));
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									if (0 < m_vecFilterJobs.size())
									{
										pImage->m_bVisible = FALSE;
										auto & job = m_vecFilterJobs.back();
										FilterParent(pSeries, job.m_vecFilters, job.m_filterType);
									}
								}

								// Remove item
								loc.m_pListItem->erase(loc.m_iter);

								// Find whether all items are from the same study/series
								if (!bMultipleStudies)
								{
									LPCWSTR pszStudyId = wcschr(pszNs, L'/');
									if (NULL != pszStudyId)
									{
										CString strCurrStudy(pszNs, (int)(pszStudyId - pszNs));
										if (strPrevStudy.IsEmpty())
										{
											strPrevStudy = strCurrStudy;
										}
										else
										{
											if (0 != lstrcmp(strCurrStudy, strPrevStudy))
											{
												bMultipleStudies = TRUE;
											}
										}

										// Find series
										if (!bMultipleSeries)
										{
											++pszStudyId;
											LPCWSTR pszSeriesId = wcschr(pszStudyId, L'/');
											if (NULL != pszSeriesId)
											{
												CString strCurrSeries(pszNs, (int)(pszSeriesId - pszNs));
												if (strPrevSeries.IsEmpty())
												{
													strPrevSeries = strCurrSeries;
												}
												else
												{
													if (0 != lstrcmp(strCurrSeries, strPrevSeries))
													{
														bMultipleSeries = TRUE;
													}
												}
											}
											else if (!strPrevSeries.IsEmpty())
											{
												bMultipleSeries = TRUE;
											}
											else
											{
												strPrevSeries = pszNs;
											}
										}
									}
									else if (!strPrevStudy.IsEmpty())
									{
										bMultipleStudies = TRUE;
									}
									else
									{
										strPrevStudy = pszNs;
									}
								}
							}
						}
					}
				}

				// Update selection

				// If the was an error during Delete operation - don't update selection
				// Items that failed to delete should still be selected

				CString strParent(L"");
				if (0 == dwError && lCount > 0)
				{
					bDeletedItems = TRUE;

					// 1 study
					if (1 == iNumStudy && 0 == iNumSeries && 0 == iNumImage)
					{
						bFindNext = TRUE;
					}
					// multiple studies
					else if (
						(iNumStudy > 0 && iNumSeries > 0) ||
						(iNumStudy > 1 && 0 == iNumSeries))
					{
						NULL;
					}
					// 1 series
					else if (1 == iNumSeries && 0 == iNumImage)
					{
						bFindNext = TRUE;
					}
					// Multiple series or mix of series/images
					else if (
						(iNumSeries > 0 && iNumImage > 0) ||
						(iNumSeries > 1 && 0 == iNumImage))
					{
						if (!bMultipleStudies)
						{
							bSelectParent = TRUE;
							strParent = strPrevStudy;
						}
					}
					// 1 image
					else if (
						1 == iNumImage &&
						!bMultipleStudies &&
						!bMultipleSeries)
					{
						bFindNext = TRUE;
					}
					// Multiple images
					else if (iNumImage > 1)
					{
						if (!bMultipleStudies)
						{
							bSelectParent = TRUE;
							if (bMultipleSeries)
								strParent = strPrevStudy;
							else
								strParent = strPrevSeries;
						}
					}

					// Find new selection
					if (bFindNext)
					{
						if (0 == m_listRoot.size())
						{
							bFindNext = FALSE;
						}
					}
					else if (bSelectParent)
					{
						auto iter = itemLocation.find(strParent);
						if (iter != itemLocation.end())
						{
							CVsiItemLocation &loc = iter->second;
							pNewSelItem = *(loc.m_iter);
						}
					}
				}
			}
		}

		// Add
		{
			// Get session
			CComPtr<IVsiSession> pSession;
			if (0 == (m_dwFlags & VSI_DATA_LIST_FLAG_NO_SESSION_LINK))
			{
				CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
				VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

				hr = pServiceProvider->QueryService(
					VSI_SERVICE_SESSION_MGR,
					__uuidof(IVsiSession),
					(IUnknown**)&pSession);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			CComPtr<IXMLDOMNodeList> pListItemAdd;
			hr = pXmlDoc->selectNodes(
				CComBSTR(L"//" VSI_UPDATE_XML_ELM_UPDATE L"/"
					VSI_UPDATE_XML_ELM_ADD L"/" VSI_UPDATE_XML_ELM_ITEM),
				&pListItemAdd);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			long lCount = 0;
			hr = pListItemAdd->get_length(&lCount);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

			CComVariant vNs;
			CComVariant vPath;
			BOOL bSelected = FALSE;
			CComVariant vSelected;
			for (long l = 0; l < lCount; ++l)
			{
				CComPtr<IXMLDOMNode> pNode;
				hr = pListItemAdd->get_item(l, &pNode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

				CComQIPtr<IXMLDOMElement> pElemParam(pNode);

				vPath.Clear();
				hr = pElemParam->getAttribute(CComBSTR(VSI_UPDATE_XML_ATT_PATH), &vPath);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_ATT_PATH));

				vNs.Clear();
				hr = pElemParam->getAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), &vNs);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_ATT_NS));

				bSelected = FALSE;

				vSelected.Clear();
				hr = pElemParam->getAttribute(CComBSTR(VSI_UPDATE_XML_SELECT), &vSelected);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_SELECT));

				if (S_OK == hr)
				{
					hr = vSelected.ChangeType(VT_BOOL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
					bSelected = (VARIANT_FALSE != V_BOOL(&vSelected));
				}

				LPCWSTR pszNs = V_BSTR(&vNs);
				CString strOrignalNs(pszNs);

				if (VSI_DATA_LIST_STUDY == m_typeRoot)
				{
					NULL;
				}
				else if (VSI_DATA_LIST_SERIES == m_typeRoot)
				{
					pszNs = wcschr(pszNs, L'/');
					VSI_VERIFY(pszNs != NULL, VSI_LOG_ERROR, NULL);
					++pszNs;  // Skip '/'
				}
				else if (VSI_DATA_LIST_IMAGE == m_typeRoot)
				{
					pszNs = wcschr(pszNs, L'/');
					VSI_VERIFY(pszNs != NULL, VSI_LOG_ERROR, NULL);
					++pszNs;  // Skip '/'
					pszNs = wcschr(pszNs, L'/');
					VSI_VERIFY(pszNs != NULL, VSI_LOG_ERROR, NULL);
					++pszNs;  // Skip '/'
				}
				else
				{
					_ASSERT(0);
				}

				// Find parents
				LPCWSTR pszItemNs(NULL); 
				LPWSTR pszSpt = (LPWSTR)wcsrchr(pszNs, L'/');
				if (NULL != pszSpt)
				{
					pszItemNs = pszSpt + 1;
					*pszSpt = 0;
				}
				else
				{
					pszItemNs = pszNs;
					pszNs = NULL;
				}

				if (NULL == pszNs)
				{
					// No parent
					if (VSI_DATA_LIST_STUDY == m_typeRoot)
					{
						CComPtr<IVsiStudy> pStudy;
						hr = m_pStudyMgr->GetStudy(V_BSTR(&vNs), &pStudy);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

						if (S_OK == hr)
						{
							AddStudy(pStudy, TRUE);
							bRefreshSelectedItems = TRUE;
						}
					}
					else if (VSI_DATA_LIST_SERIES == m_typeRoot)
					{
						CComPtr<IVsiSeries> pSeries;
						hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

						CComQIPtr<IVsiPersistSeries> pPersist(pSeries);
						hr = pPersist->LoadSeriesData(V_BSTR(&vPath), 0, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

						AddSeries(pSeries, TRUE);
						bRefreshSelectedItems = TRUE;
					}
					else if (VSI_DATA_LIST_IMAGE == m_typeRoot)
					{
						CComPtr<IVsiImage> pImage;
						hr = pImage.CoCreateInstance(__uuidof(VsiImage));
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

						CComQIPtr<IVsiPersistImage> pPersist(pImage);
						hr = pPersist->LoadImageData(V_BSTR(&vPath), NULL, NULL, 0);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

						AddImage(pImage, TRUE);
						bRefreshSelectedItems = TRUE;
					}
				}
				else
				{
					// Refresh location map
					itemLocation.clear();
					UpdateItemLocationMap(itemLocation, L"", m_listRoot);

					// Find parent
					CString strNs(pszNs);
					auto iter = itemLocation.find(strNs);
					if (iter != itemLocation.end())
					{
						CVsiItemLocation &loc = iter->second;

						CVsiItem *pItem = *(loc.m_iter);

						switch (pItem->m_type)
						{
						case VSI_DATA_LIST_STUDY:
							{
								CVsiStudy *pStudy = (CVsiStudy*)pItem;
								if (pStudy->m_bExpanded)
								{
									CComPtr<IVsiSeries> pSeries;
									hr = pStudy->m_pStudy->GetSeries(pszItemNs, &pSeries);
									VSI_CHECK_S_OK(hr, VSI_LOG_ERROR,	NULL);

									pStudy->AddItem(pSeries);
									pStudy->Sort();
								}
								else
								{
									pStudy->Expand(FALSE);
								}

								// Find new item if it needs to be selected
								if (bSelected)
								{
									itemLocation.clear();
									UpdateItemLocationMap(itemLocation, L"", m_listRoot);

									auto iterSel = itemLocation.find(strOrignalNs);
									if (iterSel != itemLocation.end())
									{
										CVsiItemLocation &locSel = iterSel->second;
										pNewSelItem = *(locSel.m_iter);
										bAddSelectedItems = TRUE;
									}
								}
							}
							break;

						case VSI_DATA_LIST_SERIES:
							{
								CVsiSeries *pSeries = (CVsiSeries*)pItem;
								if (pSeries->m_bExpanded)
								{
									CComPtr<IVsiImage> pImage;

									// Use image in session if available
									if (NULL != pSession)
									{
										int iSlot;
										if (S_OK == pSession->GetIsImageLoaded(strOrignalNs, &iSlot))
										{
											hr = pSession->GetImage(iSlot, &pImage);
											VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
										}
									}

									if (NULL == pImage)
									{ 
										hr = pImage.CoCreateInstance(__uuidof(VsiImage));
										VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

										hr = pImage->SetSeries(pSeries->m_pSeries);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

										CComQIPtr<IVsiPersistImage> pPersist(pImage);
										hr = pPersist->LoadImageData(V_BSTR(&vPath), NULL, NULL, 0);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
									
										hr = pSeries->m_pSeries->AddImage(pImage);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
									}

									pSeries->AddItem(pImage);
									pSeries->Sort();
								}
								else
								{
									pSeries->Expand(FALSE);
								}

								// find new item if it needs to be selected
								if (bSelected)
								{
									itemLocation.clear();
									UpdateItemLocationMap(itemLocation, L"", m_listRoot);

									auto iterSel = itemLocation.find(strOrignalNs);
									if (iterSel != itemLocation.end())
									{
										CVsiItemLocation &locSel = iterSel->second;
										pNewSelItem = *(locSel.m_iter);
										bAddSelectedItems = TRUE;
									}
								}
							}
							break;

						default:
							_ASSERT(0);
						}

						UpdateChildrenStatus(pItem);
					}
					else
					{
						// Image is being added to the existing series which is
						// under the study that is collapsed
					}
				}
			}
		}

		// Refresh
		{
			CComPtr<IXMLDOMNodeList> pListItemRefresh;
			hr = pXmlDoc->selectNodes(
				CComBSTR(L"//" VSI_UPDATE_XML_ELM_UPDATE L"/"
					VSI_UPDATE_XML_ELM_REFRESH L"/" VSI_UPDATE_XML_ELM_ITEM),
				&pListItemRefresh);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			long lCount = 0;
			hr = pListItemRefresh->get_length(&lCount);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

			if (lCount > 0)
			{
				// Refresh location map
				itemLocation.clear();
				UpdateItemLocationMap(itemLocation, L"", m_listRoot);

				CComVariant vNs;
				BOOL bSelected = FALSE, bCollapse = FALSE;
				CComVariant vSelected, vCollapse;
				for (long l = 0; l < lCount; ++l)
				{
					CComPtr<IXMLDOMNode> pNode;
					hr = pListItemRefresh->get_item(l, &pNode);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

					CComQIPtr<IXMLDOMElement> pElemParam(pNode);

					vNs.Clear();
					hr = pElemParam->getAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), &vNs);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_ATT_NS));

					LPCWSTR pszNs = V_BSTR(&vNs);

					bSelected = FALSE;

					vSelected.Clear();
					hr = pElemParam->getAttribute(CComBSTR(VSI_UPDATE_XML_SELECT), &vSelected);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_SELECT));

					if (S_OK == hr)
					{
						hr = vSelected.ChangeType(VT_BOOL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
						bSelected = (VARIANT_FALSE != V_BOOL(&vSelected));
					}

					vCollapse.Clear();
					hr = pElemParam->getAttribute(CComBSTR(VSI_UPDATE_XML_COLLAPSE), &vCollapse);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_COLLAPSE));

					if (S_OK == hr)
					{
						hr = vCollapse.ChangeType(VT_BOOL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
						bCollapse = (VARIANT_FALSE != V_BOOL(&vCollapse));
					}

					if (VSI_DATA_LIST_STUDY == m_typeRoot)
					{
						NULL;
					}
					else if (VSI_DATA_LIST_SERIES == m_typeRoot)
					{
						pszNs = wcschr(pszNs, L'/');
						VSI_VERIFY(pszNs != NULL, VSI_LOG_ERROR, NULL);
						++pszNs;  // Skip '/'
					}
					else if (VSI_DATA_LIST_IMAGE == m_typeRoot)
					{
						pszNs = wcschr(pszNs, L'/');
						VSI_VERIFY(pszNs != NULL, VSI_LOG_ERROR, NULL);
						++pszNs;  // Skip '/'
						pszNs = wcschr(pszNs, L'/');
						VSI_VERIFY(pszNs != NULL, VSI_LOG_ERROR, NULL);
						++pszNs;  // Skip '/'
					}
					else
					{
						_ASSERT(0);
					}

					CString strNs(pszNs);
					auto iter = itemLocation.find(strNs);
					if (iter != itemLocation.end())
					{
						// Item visible
						CVsiItemLocation &loc = iter->second;
						CVsiItem *pItem = *(loc.m_iter);
						if (bSelected)
						{
							pItem->m_bSelected = TRUE;

							bRefreshSelectedItems = TRUE;
						}

						if (bCollapse && pItem->m_bExpanded)
						{
							pItem->Collapse();
						}

						// Update item's status
						pItem->Reload();
						UpdateItemStatus(pItem, FALSE);
					}
				}
			}
		}

		UpdateList();

		if (0 < ListView_GetItemCount(*this))
		{
			// Set new selection
			if (bDeletedItems)
			{
				if (NULL != pNewSelItem)
				{
					int iItem = ListView_GetTopIndex(*this);
					if (iItem >= 0)
					{
						int iCurIndex = iItem;
						CVsiItem *pItem = NULL;
						while ((iCurIndex <= (iItem + ListView_GetCountPerPage(*this))) &&
								iCurIndex < (int)m_vecItem.size())
						{
							pItem = m_vecItem[iCurIndex];
							if (pItem == pNewSelItem)
							{
								iItem = iCurIndex;
								break;
							}

							++iCurIndex;
						}
					}
					ListView_SetItemState(
						*this, iItem,
						LVIS_FOCUSED | LVIS_SELECTED,
						LVIS_FOCUSED | LVIS_SELECTED);
				}
				else
				{
					int iSel(-1);
					GetSelectedItem(&iSel);
					if (-1 == iSel)
					{
						// No selection , select top item
						iSel = ListView_GetTopIndex(*this);
						ListView_SetItemState(
							*this, iSel,
							LVIS_FOCUSED | LVIS_SELECTED,
							LVIS_FOCUSED | LVIS_SELECTED);
					}
					ListView_EnsureVisible(*this, iSel, TRUE);
				}
			}
			else if (bAddSelectedItems)
			{
				if (NULL != pNewSelItem)
				{
					BOOL bFound = FALSE;
					int iCurIndex = 0;
					CVsiItem *pItem = NULL;
					while (iCurIndex < (int)m_vecItem.size())
					{
						pItem = m_vecItem[iCurIndex];
						if (pItem == pNewSelItem)
						{
							bFound = TRUE;
							break;
						}

						++iCurIndex;
					}

					if (bFound)
					{
						ListView_EnsureVisible(*this, iCurIndex, TRUE);
						ListView_SetItemState(
							*this, iCurIndex,
							LVIS_FOCUSED | LVIS_SELECTED,
							LVIS_FOCUSED | LVIS_SELECTED);
					}
				}
			}
			else if (bRefreshSelectedItems)
			{
				int iSel(-1);
				GetSelectedItem(&iSel);
				if (-1 != iSel)
				{
					ListView_EnsureVisible(*this, iSel, TRUE);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::GetItem(
	int iIndex, VSI_DATA_LIST_TYPE *pType, IUnknown **ppUnk)
{
	HRESULT hr = S_OK;

	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pType, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(ppUnk, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(iIndex < (int)m_vecItem.size(), VSI_LOG_ERROR, NULL);

		CVsiItem *pItem = m_vecItem[iIndex];

		*pType = pItem->m_type;

		if (VSI_DATA_LIST_STUDY == pItem->m_type)
		{
			CVsiStudy* pStudy = dynamic_cast<CVsiStudy*>(pItem);
			VSI_VERIFY(nullptr != pStudy, VSI_LOG_ERROR, NULL);

			*ppUnk = pStudy->m_pStudy;
		}
		else if (VSI_DATA_LIST_SERIES == pItem->m_type)
		{
			CVsiSeries* pSeries = dynamic_cast<CVsiSeries*>(pItem);
			VSI_VERIFY(nullptr != pSeries, VSI_LOG_ERROR, NULL);

			*ppUnk = pSeries->m_pSeries;
		}
		else if (VSI_DATA_LIST_IMAGE == pItem->m_type)
		{
			CVsiImage* pImage = dynamic_cast<CVsiImage*>(pItem);
			VSI_VERIFY(nullptr != pImage, VSI_LOG_ERROR, NULL);

			*ppUnk = pImage->m_pImage;
		}

		if (*ppUnk != NULL)
			(*ppUnk)->AddRef();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::GetItemStatus(
	LPCOLESTR pszNamespace,
	VSI_DATA_LIST_ITEM_STATUS *pdwStatus)
{
	HRESULT hr = S_OK;

	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pdwStatus, VSI_LOG_ERROR, NULL);

		*pdwStatus = VSI_DATA_LIST_ITEM_STATUS_NORMAL;

		if (0 < m_vecFilterJobs.size())
		{
			CComPtr<IUnknown> pUnkItem;
			hr = m_pStudyMgr->GetItem(pszNamespace, &pUnkItem);
			for (;S_OK == hr;)
			{
				CComQIPtr<IVsiStudy> pStudy(pUnkItem);
				if (NULL != pStudy)
				{
					CVsiStudy study;
					study.m_pRoot = this;
					study.m_pStudy = pStudy;
					study.m_pSort = &m_sort;

					auto const & job = m_vecFilterJobs.back();
					Filter(&study, job.m_vecFilters, job.m_filterType);

					if (!study.m_bVisible)
					{
						*pdwStatus |= VSI_DATA_LIST_ITEM_STATUS_HIDDEN;
					}

					break;
				}

				CComQIPtr<IVsiSeries> pSeries(pUnkItem);
				if (NULL != pSeries)
				{
					CVsiSeries series;
					CVsiStudy study;

					CComPtr<IVsiStudy> pStudy;
					hr = pSeries->GetStudy(&pStudy);
					if (S_OK == hr)
					{
						study.m_pRoot = this;
						study.m_pStudy = pStudy;
						study.m_pSort = &m_sort;
						series.m_pParent = &study;
					}

					series.m_pRoot = this;
					series.m_pSeries = pSeries;
					series.m_pSort = &m_sort;

					auto const & job = m_vecFilterJobs.back();
					Filter(&series, job.m_vecFilters, job.m_filterType);

					if (!series.m_bVisible)
					{
						*pdwStatus |= VSI_DATA_LIST_ITEM_STATUS_HIDDEN;
					}

					break;
				}

				CComQIPtr<IVsiImage> pImage(pUnkItem);
				if (NULL != pImage)
				{
					CVsiImage image;
					CVsiSeries series;
					CVsiStudy study;

					CComPtr<IVsiSeries> pSeries;
					hr = pImage->GetSeries(&pSeries);
					if (S_OK == hr)
					{
						CComPtr<IVsiStudy> pStudy;
						hr = pSeries->GetStudy(&pStudy);
						if (S_OK == hr)
						{
							study.m_pRoot = this;
							study.m_pStudy = pStudy;
							study.m_pSort = &m_sort;
							series.m_pParent = &study;
						}

						series.m_pRoot = this;
						series.m_pSeries = pSeries;
						series.m_pSort = &m_sort;
						image.m_pParent = &series;
					}

					image.m_pRoot = this;
					image.m_pImage = pImage;
					image.m_pSort = &m_sort;

					auto const & job = m_vecFilterJobs.back();
					Filter(&image, job.m_vecFilters, job.m_filterType);

					if (!image.m_bVisible)
					{
						*pdwStatus |= VSI_DATA_LIST_ITEM_STATUS_HIDDEN;
					}

					break;
				}

				break;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::AddItem(
	IUnknown *pUnkItem, VSI_DATA_LIST_SELECTION dwSelect)
{
	HRESULT hr = S_OK;

	CCritSecLock csl(m_cs);

	try
	{
		if (VSI_DATA_LIST_SELECTION_CLEAR & dwSelect)
		{
			ListView_SetItemState(*this, -1, 0, LVIS_FOCUSED | LVIS_SELECTED);
		}

		CComQIPtr<IVsiStudy> pStudy(pUnkItem);
		if (pStudy != NULL)
		{
			CComVariant vId;
			hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CVsiStudy *pStudyTest = NULL;
			for (auto iter = m_listRoot.begin(); iter != m_listRoot.end(); ++iter)
			{
				pStudyTest = (CVsiStudy*)*iter;

				CComVariant vIdTest;
				hr = pStudyTest->m_pStudy->GetProperty(
					VSI_PROP_STUDY_ID, &vIdTest);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (0 == lstrcmp(V_BSTR(&vId), V_BSTR(&vIdTest)))
				{
					break;
				}

				pStudyTest = NULL;
			}

			if (NULL == pStudyTest)
			{
				CVsiStudy *p = new CVsiStudy;
				VSI_CHECK_MEM(p, VSI_LOG_ERROR, NULL);

				VSI_STUDY_ERROR dwErrorCode(VSI_STUDY_ERROR_NONE);
				pStudy->GetErrorCode(&dwErrorCode);
				if (VSI_STUDY_ERROR_VERSION_INCOMPATIBLE & dwErrorCode)
				{
					p->m_iStateImage = VSI_DATA_STATE_IMAGE_WARNING;
				}
				else if (VSI_STUDY_ERROR_NONE != dwErrorCode)
				{
					p->m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
				}
				p->m_pRoot = this;
				p->m_pStudy.Attach(pStudy.Detach());
				p->m_pSort = &m_sort;
				if (0 < (VSI_DATA_LIST_SELECTION_UPDATE & dwSelect))
				{
					p->m_bSelected = TRUE;
					p->m_bFocused = TRUE;
				}

				if (0 < m_vecFilterJobs.size())
				{
					auto const & job = m_vecFilterJobs.back();
					Filter(p, job.m_vecFilters, job.m_filterType);
				}

				m_listRoot.push_back(p);
			}

			UpdateList();

			if (0 < (VSI_DATA_LIST_SELECTION_ENSURE_VISIBLE & dwSelect))
			{
				int iSel = ListView_GetNextItem(*this, -1, LVNI_SELECTED);
				if (iSel >= 0)
				{
					ListView_EnsureVisible(*this, iSel, FALSE);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiDataListWnd::RemoveItem(
	LPCOLESTR pszNamespace, VSI_DATA_LIST_SELECTION dwSelect)
{
	HRESULT hr = S_OK;

	CCritSecLock csl(m_cs);

	try
	{
		if (VSI_DATA_LIST_SELECTION_CLEAR & dwSelect)
		{
			ListView_SetItemState(*this, -1, 0, LVIS_FOCUSED | LVIS_SELECTED);
		}

		CVsiItem *pItem = GetItem(pszNamespace, true, NULL);

		if (pItem != NULL)
		{
			CVsiListItem *pListChild = NULL;
			CVsiItem *pItemParent = NULL;

			if (NULL != pItem->m_pParent)
			{
				pItemParent = pItem->m_pParent;
				pListChild = &pItemParent->m_listChild;
			}
			else
			{
				pListChild = &m_listRoot;
			}
			
			if (NULL != pListChild)
			{
				auto iterSel = pListChild->end();
				for (auto iter = pListChild->begin(); iter != pListChild->end(); ++iter)
				{
					CVsiItem *pItemTest = (CVsiItem*)*iter;
					if (pItem == pItemTest)
					{
						iterSel = pListChild->erase(iter);
						break;
					}
				}
				
				if (VSI_DATA_LIST_SELECTION_UPDATE & dwSelect)
				{
					if (0 < pListChild->size())
					{
						if (iterSel == pListChild->end())
						{
							--iterSel;
						}

						CVsiItem *pItemSel = (CVsiItem*)*iterSel;
						pItemSel->m_bSelected = TRUE;
						pItemSel->m_bFocused = TRUE;
					}
					else if (NULL != pItemParent)
					{
						pItemParent->m_bSelected = TRUE;
						pItemParent->m_bFocused = TRUE;
					}
				}
			}

			UpdateList();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <Summary>
///	Get enumerator of the root items
/// </Summary>
STDMETHODIMP CVsiDataListWnd::GetItems(
	VSI_DATA_LIST_COLLECTION *pItems)
{
	HRESULT hr = S_OK;

	CCritSecLock csl(m_cs);

	if (NULL != pItems->ppEnumStudy ||
		NULL != pItems->piStudyCount)
	{
		typedef CComObject<
			CComEnum<
				IVsiEnumStudy,
				&__uuidof(IVsiEnumStudy),
				IVsiStudy*, _CopyInterface<IVsiStudy>
			>
		> CVsiEnumStudy;

		IVsiStudy **ppStudy = NULL;
		CVsiEnumStudy *pEnum = NULL;
		int i = 0;

		try
		{
			VSI_CHECK_ARG(pItems, VSI_LOG_ERROR, NULL);

			typedef std::list<IVsiStudy*> CVsiListStudy;
			typedef CVsiListStudy::iterator CVsiListStudyIter;

			CVsiListStudy listStudy;
			IVsiStudy *piStudy = NULL;
			CVsiItem *pItem = NULL;

			for (auto iter = m_listRoot.begin(); iter != m_listRoot.end(); ++iter)
			{
				pItem = *iter;
				if (VSI_DATA_LIST_STUDY == pItem->m_type)
				{
					CVsiStudy *pStudy = dynamic_cast<CVsiStudy*>(pItem);
					VSI_VERIFY(nullptr != pStudy, VSI_LOG_ERROR, NULL);

					piStudy = pStudy->m_pStudy;
					listStudy.push_back(piStudy);
				}
			}

			if (NULL != pItems->piStudyCount)
			{
				*pItems->piStudyCount = (int)listStudy.size();
			}

			if (NULL != pItems->ppEnumStudy)
			{
				ppStudy = new IVsiStudy*[listStudy.size()];
				VSI_CHECK_MEM(ppStudy, VSI_LOG_ERROR, NULL);

				CVsiListStudyIter iterStudy;
				for (iterStudy = listStudy.begin();
					iterStudy != listStudy.end();
					++iterStudy)
				{
					IVsiStudy *pStudy = *iterStudy;
					pStudy->AddRef();

					ppStudy[i++] = pStudy;
				}

				pEnum = new CVsiEnumStudy;
				VSI_CHECK_MEM(pEnum, VSI_LOG_ERROR, NULL);

				hr = pEnum->Init(ppStudy, ppStudy + i, NULL, AtlFlagTakeOwnership);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				ppStudy = NULL;  // pEnum is the owner

				hr = pEnum->QueryInterface(
					__uuidof(IVsiEnumStudy),
					(void**)pItems->ppEnumStudy);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				pEnum = NULL;
			}
		}
		VSI_CATCH(hr);

		if (NULL != pEnum) delete pEnum;
		if (NULL != ppStudy)
		{
			for (; i > 0; i--)
			{
				if (NULL != ppStudy[i])
					ppStudy[i]->Release();
			}

			delete [] ppStudy;
		}
	}

	return hr;
}

HRESULT CVsiDataListWnd::RemoveDuplicateChildren(
	CVsiListItem &listParent,
	CVsiListItem &listChild)
{
	for (auto iterParent = listParent.begin(); iterParent != listParent.end(); ++iterParent)
	{
		CVsiItem *pItemParent = *iterParent;

		for (auto iterChild = listChild.begin(); iterChild != listChild.end();)
		{
			CVsiItem *pItemChild = *iterChild;

			if (pItemChild->m_pParent == pItemParent)
			{
				iterChild = listChild.erase(iterChild);
			}
			else
			{
				++iterChild;
			}
		}
	}

	// If all visible children are selected, switch selection to parent
	for (auto iterChild = listChild.begin(); iterChild != listChild.end();)
	{
		CVsiItem *pItemChild = *iterChild;
		if (NULL != pItemChild->m_pParent)
		{
			bool bMove(true);

			CVsiItem *pItemParent = pItemChild->m_pParent;

			for (auto iter = pItemParent->m_listChild.begin();
				iter != pItemParent->m_listChild.end();
				++iter)
			{
				CVsiItem *pItemTest = *iter;
				if (pItemTest->m_bVisible)
				{
					auto find = std::find(listChild.begin(), listChild.end(), pItemTest);
					if (find == listChild.end())
					{
						bMove = false;
						break;
					}
				}
			}

			if (bMove)
			{
				auto find = std::find(listParent.begin(), listParent.end(), pItemParent);
				if (find == listParent.end())
				{
					listParent.push_back(pItemParent);
				}

				iterChild = listChild.erase(iterChild);
			}
			else
			{
				++iterChild;
			}
		}
	}

	return S_OK;
}

HRESULT CVsiDataListWnd::RemoveDuplicateParent(
	CVsiListItem &listParent,
	CVsiListItem &listChild)
{
	for (auto iterParent = listParent.begin(); iterParent != listParent.end(); )
	{
		CVsiItem *pItemParent = *iterParent;

		BOOL bFound = FALSE;
		for (auto iterChild = listChild.begin(); iterChild != listChild.end(); ++iterChild)
		{
			CVsiItem *pItemChild = *iterChild;
			bFound = pItemParent->IsParentOf(*pItemChild);
			if (bFound)
				break;
		}
		if (bFound)
		{
			iterParent = listParent.erase(iterParent);
		}
		else
		{
			++iterParent;
		}
	}

	return S_OK;
}

// for now it's used ONLY for series
HRESULT CVsiDataListWnd::RemoveDuplicates(CVsiListItem &setIndex)
{
	HRESULT hr = S_OK;

	for (auto iterPrime = setIndex.begin(); iterPrime != setIndex.end(); ++iterPrime)
	{
		CVsiItem *pItemPrime = *iterPrime;

		auto iterBelow = iterPrime;
		while (++iterBelow != setIndex.end())
		{
			CVsiItem *pItemBelow = *iterBelow;

			// Compare Series Ids
			CVsiSeries* pSeriesPrime = (CVsiSeries*)pItemPrime;

			CComVariant vNsSeriesPrime;
			hr = pSeriesPrime->m_pSeries->GetProperty(VSI_PROP_SERIES_NS, &vNsSeriesPrime);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CVsiSeries* pSeriesBelow = (CVsiSeries*)pItemBelow;

			CComVariant vNsSeriesBelow;
			hr = pSeriesBelow->m_pSeries->GetProperty(VSI_PROP_SERIES_NS, &vNsSeriesBelow);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (vNsSeriesPrime == vNsSeriesBelow)
				iterBelow = setIndex.erase(iterBelow);
		}
	}

	return S_OK;
}

HRESULT CVsiDataListWnd::MoveToChildren(
	CVsiListItem &listParent,
	CVsiListItem &listChild,
	bool bVisibleOnly)
{
	for (auto iterParent = listParent.begin();
		iterParent != listParent.end(); ++iterParent)
	{
		CVsiItem *pItemParent = *iterParent;

		BOOL bExpanded = pItemParent->m_bExpanded;
		if (!bExpanded)
			pItemParent->Expand(TRUE);

		for (auto iter = pItemParent->m_listChild.begin();
			iter != pItemParent->m_listChild.end();
			++iter)
		{
			CVsiItem *pItemTest = *iter;
			if (!bVisibleOnly || pItemTest->m_bVisible)
			{
				listChild.push_back(pItemTest);
			}
		}

		if (!bExpanded)
		{
			pItemParent->m_bExpanded = FALSE;
		}
	}

	listParent.empty();

	return S_OK;
}

// This function will adjust selection for special case when we have a last image
// selected in a active study, in which case we do not want to delete the study itself
// but image only.
HRESULT CVsiDataListWnd::AdjustForActiveSeries(
	CVsiListItem &listIndexStudy,
	CVsiListItem &listIndexSeries,
	CVsiListItem &listIndexImage,
	IVsiSeries *pSeriesActive)
{
	HRESULT hr(S_OK);

	try
	{
		if (NULL == pSeriesActive)
			return hr;

		CComVariant vNsSeriesActive;
		hr = pSeriesActive->GetProperty(VSI_PROP_SERIES_NS, &vNsSeriesActive);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Found the active series
		for (auto iterSeries = listIndexSeries.begin(); iterSeries != listIndexSeries.end(); ++iterSeries)
		{
			CVsiItem *pItemSeries = *iterSeries;
			
			CVsiSeries *pSeries = dynamic_cast<CVsiSeries*>(pItemSeries);
			VSI_VERIFY(nullptr != pSeries, VSI_LOG_ERROR, NULL);
			
			CComVariant vNsSeries;
			hr = pSeries->m_pSeries->GetProperty(VSI_PROP_SERIES_NS, &vNsSeries);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
			if (vNsSeries != vNsSeriesActive)
				continue;

			for (auto iterStudy = listIndexStudy.begin(); iterStudy != listIndexStudy.end(); ++iterStudy)
			{
				CVsiItem *pItemStudy = *iterStudy;

				if (pItemStudy->IsParentOf(*pItemSeries))
				{
					long lImageCnt = 0;
					BOOL bFocused = FALSE;
					for (auto iterImage = listIndexImage.begin(); iterImage != listIndexImage.end(); ++iterImage)
					{
						CVsiItem *pItemImage = *iterImage;
						if (pItemSeries->IsParentOf(*pItemImage))
						{
							bFocused = pItemImage->m_bFocused;
							lImageCnt++;
						}
					}
					// Here we check if this active series has only one image in it
					// and if the focus is on the image, which means we only want to delete the
					// image. if the series has more then one image, then series is not selected
					// automatically when selecting an image.
					if (1 == lImageCnt && bFocused)
					{
						listIndexStudy.erase(iterStudy);
						listIndexSeries.erase(iterSeries);
					}

					return S_OK;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

void CVsiDataListWnd::UpdateList()
{
	if (S_OK == LockData())
	{
		m_vecItem.clear();

		UpdateViewVector(&m_listRoot);

		// LVM_SETITEMCOUNT is thread safe
		ListView_SetItemCountEx(*this, m_vecItem.size(), LVSICF_NOSCROLL);

		LV_ITEM lvi = { 0 };
		lvi.stateMask = LVIS_FOCUSED | LVIS_SELECTED;

		int iCount = (int)m_vecItem.size();
		for (int i = 0; i < iCount; ++i)
		{
			CVsiItem *pTest = m_vecItem[i];

			lvi.state = 0;

			if (pTest->m_bSelected)
			{
				lvi.state |= LVIS_SELECTED;
			}
			if (pTest->m_bFocused)
			{
				lvi.state |= LVIS_FOCUSED;
			}

			SendMessage(LVM_SETITEMSTATE, i, (LPARAM)&lvi);
		}

		UnlockData();
	}
}

LRESULT CVsiDataListWnd::OnDestroy(UINT, WPARAM, LPARAM, BOOL& bHandled)
{
	Clear();

	m_ilHeader.Destroy();

	if (!m_ilState.IsNull())
	{
		ListView_SetImageList(*this, NULL, LVSIL_STATE);
		m_ilState.Destroy();
	}

	if (!m_ilSmall.IsNull())
	{
		ListView_SetImageList(*this, NULL, LVSIL_SMALL);
		m_ilSmall.Destroy();
	}

	m_pModeManager.Release();
	m_pStudyMgr.Release();
	m_pPdm.Release();
	m_pApp.Release();

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiDataListWnd::OnSettingChange(UINT, WPARAM, LPARAM, BOOL& bHandled)
{
	bHandled = FALSE;

	RedrawWindow(NULL, NULL, RDW_INVALIDATE);

	return 0;
}

LRESULT CVsiDataListWnd::OnButtonDown(UINT, WPARAM, LPARAM lp, BOOL& bHandled)
{
	// Prevents removal of selection

	LVHITTESTINFO lviti = { 0 };
	
	lviti.pt.x = GET_X_LPARAM(lp);
	lviti.pt.y = GET_Y_LPARAM(lp);

	ListView_HitTest(*this, &lviti);

	if (lviti.flags & LVHT_ONITEM)
	{
		bHandled = FALSE;
	}

	return 0;
}

LRESULT CVsiDataListWnd::OnMouseMove(UINT, WPARAM, LPARAM lp, BOOL &bHandled)
{
	bHandled = FALSE;

	if (!m_bTrackLeave)
	{
		m_bTrackLeave = TRUE;
	
		TRACKMOUSEEVENT tme = { sizeof(TRACKMOUSEEVENT), TME_LEAVE, m_hWnd, 0 };
		TrackMouseEvent(&tme);
	}

	LVHITTESTINFO lviti = { 0 };
	
	lviti.pt.x = GET_X_LPARAM(lp);
	lviti.pt.y = GET_Y_LPARAM(lp);

	ListView_HitTest(*this, &lviti);

	if (m_iLastHotItem != lviti.iItem)
	{
		RECT rc;

		if (-1 != m_iLastHotItem)
		{
			ListView_GetItemRect(*this, m_iLastHotItem, &rc, LVIR_BOUNDS);
			
			InvalidateRect(&rc);

			m_iLastHotItem = -1;
		}

		if (-1 != lviti.iItem)
		{
			m_iLastHotItem = lviti.iItem;
		
			ListView_GetItemRect(*this, m_iLastHotItem, &rc, LVIR_BOUNDS);
			
			InvalidateRect(&rc);
		}
	}

	if (0 <= m_iLastHotItem)
	{
		CRect rcHitTest;
		ListView_GetItemRect(*this, m_iLastHotItem, &rcHitTest, LVIR_ICON);
		rcHitTest.left -= 20;

		int iHotState = -1;
		if (rcHitTest.PtInRect(lviti.pt))
		{
			iHotState = m_iLastHotItem;
		}

		if (iHotState != m_iLastHotState)
		{
			m_iLastHotState = iHotState;

			RECT rc;
			ListView_GetItemRect(*this, m_iLastHotItem, &rc, LVIR_BOUNDS);
			
			InvalidateRect(&rc);
		}
	}

	return 0;
}

LRESULT CVsiDataListWnd::OnMouseLeave(UINT, WPARAM, LPARAM, BOOL&)
{
	if (m_bTrackLeave)
	{
		m_bTrackLeave = FALSE;

		if (-1 != m_iLastHotItem)
		{
			RECT rc;
			ListView_GetItemRect(*this, m_iLastHotItem, &rc, LVIR_BOUNDS);
			
			InvalidateRect(&rc);

			m_iLastHotItem = -1;
			m_iLastHotState = -1;
		}
	}

	return 0;
}

LRESULT CVsiDataListWnd::OnSetItemState(UINT, WPARAM wp, LPARAM lp, BOOL& bHandled)
{
	bHandled = FALSE;

	LVITEM *plvi = (LVITEM*)lp;

	if (0 < (plvi->stateMask & LVIS_STATEIMAGEMASK))
	{
		BYTE state = (BYTE)((plvi->state & LVIS_STATEIMAGEMASK) >> 12);

		if ((WPARAM)-1 == wp)
		{
			for (int i = 0; i < (int)m_vecItem.size(); ++i)
			{
				CVsiItem *pItem = m_vecItem[i];
				pItem->m_iStateImage = state;
			}
			
			Invalidate();
		}
		else
		{
			if (wp < m_vecItem.size())
			{
				CVsiItem *pItem = m_vecItem[wp];

				pItem->m_iStateImage = state;

				RECT rc;
				ListView_GetItemRect(*this, wp, &rc, LVIR_BOUNDS);
				InvalidateRect(&rc);
			}
		}
	}

	return 0;
}

LRESULT CVsiDataListWnd::OnGetItemState(UINT, WPARAM wp, LPARAM lp, BOOL& bHandled)
{
	LRESULT lRet(0);

	if (0 < (lp & LVIS_STATEIMAGEMASK))
	{
		if (wp < m_vecItem.size())
		{
			CVsiItem *pItem = m_vecItem[wp];

			lRet = DefWindowProc();
			lRet |= INDEXTOSTATEIMAGEMASK(pItem->m_iStateImage);
		}
	}
	else
	{
		bHandled = FALSE;
	}

	return lRet;
}

LRESULT CVsiDataListWnd::OnHdnBeginDrag(int, LPNMHDR pnmh, BOOL&)
{
	NMHEADER *pnmhr = (NMHEADER*)pnmh;

	// Don't allow dragging of the 1st column
	return (VSI_DATA_LIST_COL_NAME == pnmhr->iItem);
}

LRESULT CVsiDataListWnd::OnHdnEndDrag(int, LPNMHDR pnmh, BOOL&)
{
	NMHEADER *pnmhr = (NMHEADER*)pnmh;

	Invalidate();

	const MSG* pMsg = GetCurrentMessage();
	GetParent().SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);

	// Don't allow dragging of the 1st column
	return (VSI_DATA_LIST_COL_NAME == pnmhr->iItem);
}

LRESULT CVsiDataListWnd::OnHdnBeginTrack(int, LPNMHDR pnmh, BOOL&)
{
	NMHEADER *pnmhr = (NMHEADER*)pnmh;

	// Don't allow dragging the lock column divider
	return (VSI_DATA_LIST_COL_LOCKED == pnmhr->iItem);
}

LRESULT CVsiDataListWnd::OnHdnEndTrack(int, LPNMHDR pnmh, BOOL&)
{
	NMHEADER *pnmhr = (NMHEADER*)pnmh;

	const MSG* pMsg = GetCurrentMessage();
	GetParent().SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);

	// Don't allow dragging the lock column divider
	return (VSI_DATA_LIST_COL_LOCKED == pnmhr->iItem);
}

LRESULT CVsiDataListWnd::OnColumnClick(int, LPNMHDR pnmh, BOOL& bHandled)
{
	NMLISTVIEW *pnmlv = (NMLISTVIEW*)pnmh;

	UpdateSortIndicator(FALSE);

	if (m_sort.m_sortCol == m_mapIndexToCol.find(pnmlv->iSubItem)->second)
	{
		m_sort.m_bDescending = !m_sort.m_bDescending;

		switch (m_sort.m_sortCol)
		{
		case VSI_DATA_LIST_COL_NAME:
		case VSI_DATA_LIST_COL_DATE:
		case VSI_DATA_LIST_COL_SIZE:
			m_sort.m_bStudyDescending = m_sort.m_bDescending;
			m_sort.m_bSeriesDescending = m_sort.m_bDescending;
			m_sort.m_bImageDescending = m_sort.m_bDescending;
			break;
		case VSI_DATA_LIST_COL_LOCKED:
		case VSI_DATA_LIST_COL_OWNER:
			m_sort.m_bStudyDescending = m_sort.m_bDescending;
			break;
		case VSI_DATA_LIST_COL_ACQ_BY:
		case VSI_DATA_LIST_COL_ANIMAL_ID:
		case VSI_DATA_LIST_COL_PROTOCOL_ID:
			m_sort.m_bSeriesDescending = m_sort.m_bDescending;
			break;
		case VSI_DATA_LIST_COL_MODE:
			m_sort.m_bImageDescending = m_sort.m_bDescending;
			break;
		case VSI_DATA_LIST_COL_LENGTH:
			m_sort.m_bImageDescending = m_sort.m_bDescending;
			break;
		}
	}
	else
	{
		m_sort.m_sortCol = m_mapIndexToCol.find(pnmlv->iSubItem)->second;

		switch (m_sort.m_sortCol)
		{
		case VSI_DATA_LIST_COL_NAME:
		case VSI_DATA_LIST_COL_DATE:
		case VSI_DATA_LIST_COL_SIZE:
			m_sort.m_sortStudyCol = m_sort.m_sortCol;
			m_sort.m_bStudyDescending = m_sort.m_bDescending;
			m_sort.m_sortSeriesCol = m_sort.m_sortCol;
			m_sort.m_bSeriesDescending = m_sort.m_bDescending;
			m_sort.m_sortImageCol = m_sort.m_sortCol;
			m_sort.m_bImageDescending = m_sort.m_bDescending;
			break;
		case VSI_DATA_LIST_COL_LOCKED:
		case VSI_DATA_LIST_COL_OWNER:
			m_sort.m_sortStudyCol = m_sort.m_sortCol;
			m_sort.m_bStudyDescending = m_sort.m_bDescending;
			break;
		case VSI_DATA_LIST_COL_ACQ_BY:
		case VSI_DATA_LIST_COL_ANIMAL_ID:
		case VSI_DATA_LIST_COL_PROTOCOL_ID:
			m_sort.m_sortSeriesCol = m_sort.m_sortCol;
			m_sort.m_bSeriesDescending = m_sort.m_bDescending;
			break;
		case VSI_DATA_LIST_COL_MODE:
		case VSI_DATA_LIST_COL_LENGTH:
			m_sort.m_sortImageCol = m_sort.m_sortCol;
			m_sort.m_bImageDescending = m_sort.m_bDescending;
			break;
		}
	}

	UpdateSortIndicator(TRUE);

	UpdateList();

	// Notify parent
	NMHDR nmhdr = {
		*this,
		m_wIdCtl,
		(UINT)NM_VSI_DL_SORT_CHANGED
	};

	GetParent().SendMessage(WM_NOTIFY, m_wIdCtl, (LPARAM)&nmhdr);

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiDataListWnd::OnGetDisplayInfo(int, LPNMHDR pnmh, BOOL& bHandled)
{
	HRESULT hr(S_OK);
	bHandled = FALSE;

	try
	{
		NMLVDISPINFO *pnmv = (NMLVDISPINFO*)pnmh;
		LVITEM *pLvi = &(pnmv->item);

		if ((int)m_vecItem.size() <= pLvi->iItem)
		{
			VSI_LOG_MSG(VSI_LOG_WARNING, L"Index out of range");

			return 0;
		}

		CVsiItem *pItem = m_vecItem[pLvi->iItem];

		VSI_DATA_LIST_COL col = m_mapIndexToCol.find(pLvi->iSubItem)->second;

		if (pLvi->mask & LVIF_STATE)
		{
			if (VSI_DATA_LIST_COL_NAME == col)
			{
				if (pItem->m_iStateImage > 0)
				{
					pLvi->state = INDEXTOSTATEIMAGEMASK(pItem->m_iStateImage);
				}
				else if (m_iLevelMax > pItem->m_iLevel)
				{
					if (pItem->m_bExpanded)
					{
						if (pItem->m_listChild.size() > 0)
						{
							pLvi->state = INDEXTOSTATEIMAGEMASK(VSI_DATA_STATE_IMAGE_CLOSE);
						}
						else
						{
							pLvi->state = INDEXTOSTATEIMAGEMASK(0);
						}
					}
					else
					{
						pLvi->state = INDEXTOSTATEIMAGEMASK(VSI_DATA_STATE_IMAGE_OPEN);
					}
				}
				else
				{
					pLvi->state = INDEXTOSTATEIMAGEMASK(0);
				}
			}

			if (pItem->m_bSelected)
			{
				pLvi->state |= LVIS_SELECTED;
			}

			if (pItem->m_bFocused)
			{
				pLvi->state |= LVIS_FOCUSED;
			}
		}

		if (pLvi->mask & LVIF_INDENT)
		{
			pLvi->iIndent = pItem->m_iLevel - 1;
		}

		if (pLvi->mask & LVIF_IMAGE)
		{
			switch (col)
			{
			case VSI_DATA_LIST_COL_NAME:
				{
					if (VSI_DATA_LIST_STUDY == pItem->m_type)
					{
						BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
							VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
							VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
						if (bSecureMode)
						{
							CVsiStudy *pStudy = dynamic_cast<CVsiStudy*>(pItem);
							VSI_VERIFY(nullptr != pStudy, VSI_LOG_ERROR, NULL);

							CComVariant vOwner;
							hr = pStudy->m_pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
							if (S_OK == hr)
							{
								CComVariant vAccess;
								hr = pStudy->m_pStudy->GetProperty(VSI_PROP_STUDY_ACCESS, &vAccess);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								if (S_OK == hr && 0 == wcscmp(VSI_STUDY_ACCESS_PRIVATE, V_BSTR(&vAccess)))
								{
									pLvi->iImage = VSI_DATA_IMAGE_STUDY_ACCESS_PRIVATE;
								}
								else if (S_OK == hr && 0 == wcscmp(VSI_STUDY_ACCESS_GROUP, V_BSTR(&vAccess)))
								{
									pLvi->iImage = VSI_DATA_IMAGE_STUDY_ACCESS_GROUP;
								}
								else
								{
									pLvi->iImage = VSI_DATA_IMAGE_STUDY_ACCESS_ALL;
								}
							}
						}
					}
					else if (VSI_DATA_LIST_IMAGE == pItem->m_type)
					{
						CVsiImage *pImage = dynamic_cast<CVsiImage*>(pItem);
						VSI_VERIFY(nullptr != pImage, VSI_LOG_ERROR, NULL);

						pLvi->iImage = VSI_DATA_IMAGE_CINE;

						CComVariant vFrame;
						hr = pImage->m_pImage->GetProperty(VSI_PROP_IMAGE_FRAME_TYPE, &vFrame);
						if (S_OK == hr)
						{
							if (VARIANT_FALSE != V_BOOL(&vFrame))
							{
								pLvi->iImage = VSI_DATA_IMAGE_FRAME;
							}
						}
						if (VSI_DATA_IMAGE_FRAME != pLvi->iImage)
						{
							CComVariant vMode;
							hr = pImage->m_pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
							if (S_OK == hr)
							{
								if (NULL != wcsstr(V_BSTR(&vMode), L"3D"))
								{
									pLvi->iImage = VSI_DATA_IMAGE_3D;
								}
							}
						}
					}
				}
				break;

			case VSI_DATA_LIST_COL_LOCKED:
				{
					if (VSI_DATA_LIST_STUDY == pItem->m_type)
					{
						BOOL bLockAll = VsiGetBooleanValue<BOOL>(
							VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
							VSI_PARAMETER_SYS_LOCK_ALL, m_pPdm);

						CVsiStudy *pStudy = dynamic_cast<CVsiStudy*>(pItem);
						VSI_VERIFY(nullptr != pStudy, VSI_LOG_ERROR, NULL);

						CComVariant v;
						hr = pStudy->m_pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &v);
						if (S_OK == hr)
						{
							pLvi->iImage = (VARIANT_FALSE != V_BOOL(&v)) ?
								(bLockAll ? VSI_DATA_IMAGE_CHECKED_LOCK_ALL : VSI_DATA_IMAGE_CHECKED) : (bLockAll ? VSI_DATA_IMAGE_UNCHECKED_LOCK_ALL : VSI_DATA_IMAGE_UNCHECKED);
						}
						else
						{
							pLvi->iImage = bLockAll ? VSI_DATA_IMAGE_UNCHECKED_LOCK_ALL : VSI_DATA_IMAGE_UNCHECKED;
						}
					}
				}
				break;
			}
		}

		if (pLvi->mask & LVIF_TEXT)
		{
			pItem->GetDisplayText(col, pLvi->pszText, pLvi->cchTextMax);
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiDataListWnd::OnItemChanged(int, LPNMHDR pnmh, BOOL& bHandled)
{
	LPNMLISTVIEW pnmlv = (LPNMLISTVIEW)pnmh;

	if (0 == (pnmlv->uChanged & LVIF_STATE))
	{
		bHandled = FALSE;
		return 0;  // Not state
	}

	if (pnmlv->iItem < 0)  // For all
	{
		if ((LVIS_SELECTED & pnmlv->uNewState) && (0 == (LVIS_SELECTED & pnmlv->uOldState)))
		{
			for (auto iter = m_vecItem.begin(); iter != m_vecItem.end(); ++iter)
			{
				CVsiItem *pItem = *iter;
				pItem->m_bSelected = TRUE;
			}
		}
		else if ((LVIS_SELECTED & pnmlv->uOldState) && (0 == (LVIS_SELECTED & pnmlv->uNewState)))
		{
			for (auto iter = m_vecItem.begin(); iter != m_vecItem.end(); ++iter)
			{
				CVsiItem *pItem = *iter;
				pItem->m_bSelected = FALSE;
			}
		}
	}
	else
	{
		CVsiItem *pItem = m_vecItem[pnmlv->iItem];

		if ((LVIS_SELECTED & pnmlv->uNewState) && (0 == (LVIS_SELECTED & pnmlv->uOldState)))
		{
			pItem->m_bSelected = TRUE;

			// Selects all children if item is opened
			if (pItem->m_bExpanded)
			{
				bool bSelect = true;

				if (0 != m_vecFilterJobs.size())
				{
					// Select all children if they are all visible
					if (pItem->m_listChild.size() != pItem->GetVisibleChildrenCount(false))
					{
						bSelect = false;
					}
				}

				if (bSelect)
				{
					int i = pnmlv->iItem;
					for (CVsiListItem::size_type j = 0; j < pItem->GetVisibleChildrenCount(); ++j)
					{
						ListView_SetItemState(pnmh->hwndFrom, ++i, LVIS_SELECTED, LVIS_SELECTED);
					}
				}
			}

			// Selects parent if all the children are selected
			if (NULL != pItem->m_pParent && !pItem->m_pParent->m_bSelected)
			{
				bool bSelect = true;
				int i = pnmlv->iItem - 1;
				int iParent = i;
				auto iter = pItem->m_pParent->m_listChild.begin();
				for (; iter != pItem->m_pParent->m_listChild.end(); ++iter, --i)
				{
					CVsiItem *pTest = *iter;

					if (pTest == pItem)
					{
						iParent = i;
					}

					if (!pTest->m_bSelected || !pTest->m_bVisible)
					{
						bSelect = false;
						break;
					}
				}

				if (bSelect)
				{
					ListView_SetItemState(pnmh->hwndFrom, iParent, LVIS_SELECTED, LVIS_SELECTED);
				}
			}
		}
		else if ((LVIS_SELECTED & pnmlv->uOldState) && (0 == (LVIS_SELECTED & pnmlv->uNewState)))
		{
			pItem->m_bSelected = FALSE;

			// Un-selects all children if item is opened
			if (pItem->m_bExpanded)
			{
				BOOL bAllChildrenSelected = TRUE;

				auto iter = pItem->m_listChild.begin();
				for (; iter != pItem->m_listChild.end(); ++iter)
				{
					CVsiItem *pTest = *iter;
					if (!pTest->m_bSelected)
					{
						bAllChildrenSelected = FALSE;
						break;
					}
				}

				if (bAllChildrenSelected)
				{
					int i = pnmlv->iItem;
					for (CVsiListItem::size_type j = 0; j < pItem->GetVisibleChildrenCount(); ++j)
					{
						ListView_SetItemState(pnmh->hwndFrom, ++i, 0, LVIS_SELECTED);
					}
				}
			}

			// Un-selects parent if it is selected
			if (NULL != pItem->m_pParent && pItem->m_pParent->m_bSelected)
			{
				int i = pnmlv->iItem - 1;
				auto iter = pItem->m_pParent->m_listChild.begin();
				for (; iter != pItem->m_pParent->m_listChild.end(); ++iter, --i)
				{
					CVsiItem *pTest = *iter;
					if (pTest == pItem)
					{
						// Un-selects parent
						ListView_SetItemState(pnmh->hwndFrom, i, 0, LVIS_SELECTED | LVIS_FOCUSED);
						break;
					}
				}
			}
		}

		if ((LVIS_FOCUSED & pnmlv->uNewState) && (0 == (LVIS_FOCUSED & pnmlv->uOldState)))
		{
			pItem->m_bFocused = TRUE;
		}
		else if ((LVIS_FOCUSED & pnmlv->uOldState) && (0 == (LVIS_FOCUSED & pnmlv->uNewState)))
		{
			pItem->m_bFocused = FALSE;
		}
	}

	if (m_bNotifyStateChanged)
	{
		bHandled = FALSE;
	}

	return 0;
}

LRESULT CVsiDataListWnd::OnOdStateChanged(int, LPNMHDR pnmh, BOOL& bHandled)
{
	LPNMLVODSTATECHANGE pnmlv = (LPNMLVODSTATECHANGE)pnmh;

	for (int i = pnmlv->iFrom; i <= pnmlv->iTo; ++i)
	{
		CVsiItem *pItem = m_vecItem[i];

		pItem->m_bSelected = (LVIS_SELECTED & pnmlv->uNewState) > 0;

		if (pItem->m_bSelected && pItem->m_bExpanded)
		{
			// Selects all children
			for (CVsiListItem::size_type j = 1; j <= pItem->GetVisibleChildrenCount(); ++j)
			{
				ListView_SetItemState(pnmh->hwndFrom, i + j, LVIS_SELECTED, LVIS_SELECTED);
			}
		}
	}

	// Selects parent?
	CVsiItem *pItemFrom = m_vecItem[pnmlv->iFrom];
	CVsiItem *pItemTo = m_vecItem[pnmlv->iTo];
	if (NULL != pItemFrom->m_pParent)
	{
		BOOL bSelect = TRUE;
		int i = pnmlv->iFrom - 1;
		int iParent = i;
		auto iter = pItemFrom->m_pParent->m_listChild.begin();
		for (; iter != pItemFrom->m_pParent->m_listChild.end(); ++iter, --i)
		{
			CVsiItem *pTest = *iter;

			if (pTest == pItemFrom)
			{
				iParent = i;
			}

			if (!pTest->m_bSelected)
			{
				bSelect = FALSE;
				break;
			}
		}
		if (bSelect)
		{
			ListView_SetItemState(pnmh->hwndFrom, iParent, LVIS_SELECTED, LVIS_SELECTED);
		}
	}

	if (NULL != pItemTo->m_pParent && (pItemTo->m_pParent != pItemFrom->m_pParent))
	{
		BOOL bSelect = TRUE;
		int i = pnmlv->iTo - 1;
		int iParent = i;
		auto iter = pItemTo->m_pParent->m_listChild.begin();
		for (; iter != pItemTo->m_pParent->m_listChild.end(); ++iter, --i)
		{
			CVsiItem *pTest = *iter;

			if (pTest == pItemTo)
			{
				iParent = i;
			}

			if (!pTest->m_bSelected)
			{
				bSelect = FALSE;
				break;
			}
		}
		if (bSelect)
		{
			ListView_SetItemState(pnmh->hwndFrom, iParent, LVIS_SELECTED, LVIS_SELECTED);
		}
	}

	if (m_bNotifyStateChanged)
	{
		bHandled = FALSE;
	}

	return 0;
}

LRESULT CVsiDataListWnd::OnListItemActivate(int, LPNMHDR pnmh, BOOL& bHandled)
{
	UNREFERENCED_PARAMETER(bHandled);

	NMITEMACTIVATE *pnmia = (NMITEMACTIVATE*)pnmh;
	if (pnmia->iItem >= 0)
	{
		HRESULT hr = S_OK;
		CComPtr<IUnknown> pUnk;

		try
		{
			CVsiItem *pItem = m_vecItem[pnmia->iItem];

			if (pItem->m_iLevel < m_iLevelMax)
			{
				if (pItem->m_bExpanded && ((pnmia->iItem + 1) < (int)m_vecItem.size()))
				{
					pItem->Collapse();
				}
				else
				{
					CWaitCursor wait;

					pItem->Expand(pItem->m_bSelected);

					// Analyze the list of created items
					UpdateChildrenStatus(pItem, TRUE);
				}

				UpdateList();
			}
		}
		VSI_CATCH(hr);
	}

	return 0;
}

LRESULT CVsiDataListWnd::OnKeydown(int, LPNMHDR pnmh, BOOL& bHandled)
{
	LPNMLVKEYDOWN pnmk = (LPNMLVKEYDOWN)pnmh;

	if (VK_LEFT == pnmk->wVKey || VK_RIGHT == pnmk->wVKey)
	{
		if (GetKeyState(VK_CONTROL) >= 0)
		{
			// No Ctrl

			CVsiItem *pItemSel = NULL;
			for (auto iter = m_vecItem.begin(); iter != m_vecItem.end(); ++iter)
			{
				CVsiItem *pItem = *iter;
				if (pItem->m_bFocused)
				{
					if (pItemSel != NULL)
					{
						// More than 1 selected - ignore
						pItemSel = NULL;
						break;
					}
					pItemSel = pItem;
				}
			}

			if (pItemSel != NULL)
			{
				if (m_iLevelMax > pItemSel->m_iLevel)
				{
					if (VK_LEFT == pnmk->wVKey)
					{
						if (pItemSel->m_bExpanded)
						{
							pItemSel->Collapse();
							UpdateList();
							return 1;
						}
					}
					else if (VK_RIGHT == pnmk->wVKey)
					{
						if (!pItemSel->m_bExpanded)
						{
							CWaitCursor wait;
						
							pItemSel->Expand(pItemSel->m_bSelected);
							// Analyze the list of created items
							UpdateChildrenStatus(pItemSel);
							UpdateList();
							return 1;
						}
					}
				}
			}
		}
	}

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiDataListWnd::OnGetInfoTip(int, LPNMHDR pnmh, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		ULONG ulSysState = VsiGetBitfieldValue<ULONG>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_STATE, m_pPdm);

		if (ulSysState & VSI_SYS_STATE_ENGINEER_MODE)
		{
			LPNMLVGETINFOTIP pgit = (LPNMLVGETINFOTIP)pnmh;

			if (pgit->iItem >= 0 && pgit->iItem < (int)m_vecItem.size())
			{
				CVsiItem *pItem = m_vecItem[pgit->iItem];

				switch (pItem->m_type)
				{
				case VSI_DATA_LIST_STUDY:
					{
						CVsiStudy *pStudy = dynamic_cast<CVsiStudy*>(pItem);
						VSI_VERIFY(nullptr != pStudy, VSI_LOG_ERROR, NULL);

						CString strTip;

						{
							CComVariant vName;
							hr = pStudy->m_pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vName);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (0 != *V_BSTR(&vName)))
							{
								strTip += V_BSTR(&vName);
								strTip += L"\r\n";
							}
						}

						{
							CComVariant vDate;
							hr = pStudy->m_pStudy->GetProperty(VSI_PROP_STUDY_CREATED_DATE, &vDate);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							COleDateTime date(vDate);
							SYSTEMTIME stDate;
							date.GetAsSystemTime(stDate);

							SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

							WCHAR szDate[50], szTime[50];
							GetDateFormat(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &stDate, NULL, szDate, _countof(szDate));
							GetTimeFormat(LOCALE_USER_DEFAULT, 0, &stDate, NULL, szTime, _countof(szTime));

							strTip += L"Date Created:\t\t";
							strTip += szDate;
							strTip += L" ";
							strTip += szTime;
							strTip += L"\r\n";
						}

						{
							CComVariant vVersion;
							hr = pStudy->m_pStudy->GetProperty(VSI_PROP_STUDY_VERSION_CREATED, &vVersion);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (NULL != *V_BSTR(&vVersion)))
							{
								strTip += L"Version Created:\t";
								strTip += V_BSTR(&vVersion);
								strTip += L"\r\n";
							}
						}

						{
							CComVariant vVersion;
							hr = pStudy->m_pStudy->GetProperty(VSI_PROP_STUDY_VERSION_MODIFIED, &vVersion);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (NULL != *V_BSTR(&vVersion)))
							{
								strTip += L"Version Modified:\t";
								strTip += V_BSTR(&vVersion);
								strTip += L"\r\n";
							}
						}

						{
							CComVariant vVersion;
							hr = pStudy->m_pStudy->GetProperty(VSI_PROP_STUDY_VERSION_REQUIRED, &vVersion);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (NULL != *V_BSTR(&vVersion)))
							{
								strTip += L"Version Required:\t";
								strTip += V_BSTR(&vVersion);
								strTip += L"\r\n";
							}
						}

						{
							LONG lSeries(0);
							hr = pStudy->m_pStudy->GetSeriesCount(&lSeries, NULL);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							
							CString strSeries;
							strSeries.Format(L"Series:\t\t\t%d\r\n", lSeries);

							strTip += strSeries;
						}

						{
							CComHeapPtr<OLECHAR> pszPath;
							hr = pStudy->m_pStudy->GetDataPath(&pszPath);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							PathRemoveFileSpec(pszPath);

							__int64 iSize = VsiGetDirectorySize(pszPath);

							CString strSize;

							if (iSize < 1024)
							{
								strSize.Format(L"%d B", static_cast<int>(iSize));
							}
							else if (iSize < (1024 * 1024))
							{
								iSize /= 1024;
								strSize.Format(L"%d KB", static_cast<int>(iSize));
							}
							else if (iSize < (1024 * 1024 * 1024))
							{
								iSize /= 1024 * 1024;
								strSize.Format(L"%d MB", static_cast<int>(iSize));
							}
							else
							{
								iSize /= 1024 * 1024 * 1024;
								strSize.Format(L"%d GB", static_cast<int>(iSize));
							}

							strTip += L"Size:\t\t\t";
							strTip += strSize;
							strTip += L"\r\n";
						}

						lstrcpyn(pgit->pszText, strTip, pgit->cchTextMax);
					}
					break;

				case VSI_DATA_LIST_SERIES:
					{
						CVsiSeries *pSeries = dynamic_cast<CVsiSeries*>(pItem);
						VSI_VERIFY(nullptr != pSeries, VSI_LOG_ERROR, NULL);

						CString strTip;

						{
							CComVariant vName;
							hr = pSeries->m_pSeries->GetProperty(VSI_PROP_SERIES_NAME, &vName);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (0 != *V_BSTR(&vName)))
							{
								strTip += V_BSTR(&vName);
								strTip += L"\r\n";
							}
						}

						{
							CComVariant vDate;
							hr = pSeries->m_pSeries->GetProperty(VSI_PROP_SERIES_CREATED_DATE, &vDate);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							COleDateTime date(vDate);
							SYSTEMTIME stDate;
							date.GetAsSystemTime(stDate);

							SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

							WCHAR szDate[50], szTime[50];
							GetDateFormat(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &stDate, NULL, szDate, _countof(szDate));
							GetTimeFormat(LOCALE_USER_DEFAULT, 0, &stDate, NULL, szTime, _countof(szTime));

							strTip += L"Date Created:\t\t";
							strTip += szDate;
							strTip += L" ";
							strTip += szTime;
							strTip += L"\r\n";
						}

						{
							CComVariant vVersion;
							hr = pSeries->m_pSeries->GetProperty(VSI_PROP_SERIES_VERSION_CREATED, &vVersion);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (NULL != *V_BSTR(&vVersion)))
							{
								strTip += L"Version Created:\t";
								strTip += V_BSTR(&vVersion);
								strTip += L"\r\n";
							}
						}

						{
							CComVariant vVersion;
							hr = pSeries->m_pSeries->GetProperty(VSI_PROP_SERIES_VERSION_MODIFIED, &vVersion);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (NULL != *V_BSTR(&vVersion)))
							{
								strTip += L"Version Modified:\t";
								strTip += V_BSTR(&vVersion);
								strTip += L"\r\n";
							}
						}

						{
							CComVariant vVersion;
							hr = pSeries->m_pSeries->GetProperty(VSI_PROP_SERIES_VERSION_REQUIRED, &vVersion);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (NULL != *V_BSTR(&vVersion)))
							{
								strTip += L"Version Required:\t";
								strTip += V_BSTR(&vVersion);
								strTip += L"\r\n";
							}
						}

						{
							LONG lImages(0);
							hr = pSeries->m_pSeries->GetImageCount(&lImages, NULL);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							
							CString strImages;
							strImages.Format(L"Images:\t\t%d\r\n", lImages);

							strTip += strImages;
						}

						{
							CComHeapPtr<OLECHAR> pszPath;
							hr = pSeries->m_pSeries->GetDataPath(&pszPath);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							PathRemoveFileSpec(pszPath);

							__int64 iSize = VsiGetDirectorySize(pszPath);

							CString strSize;

							if (iSize < 1024)
							{
								strSize.Format(L"%d B", static_cast<int>(iSize));
							}
							else if (iSize < (1024 * 1024))
							{
								iSize /= 1024;
								strSize.Format(L"%d KB", static_cast<int>(iSize));
							}
							else if (iSize < (1024 * 1024 * 1024))
							{
								iSize /= 1024 * 1024;
								strSize.Format(L"%d MB", static_cast<int>(iSize));
							}
							else
							{
								iSize /= 1024 * 1024 * 1024;
								strSize.Format(L"%d GB", static_cast<int>(iSize));
							}

							strTip += L"Size:\t\t\t";
							strTip += strSize;
							strTip += L"\r\n";
						}

						lstrcpyn(pgit->pszText, strTip, pgit->cchTextMax);
					}
					break;

				case VSI_DATA_LIST_IMAGE:
					{
						CVsiImage *pImage = dynamic_cast<CVsiImage*>(pItem);
						VSI_VERIFY(nullptr != pImage, VSI_LOG_ERROR, NULL);

						CString strTip;

						{
							CComVariant vName;
							hr = pImage->m_pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vName);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (0 != *V_BSTR(&vName)))
							{
								strTip += V_BSTR(&vName);
								strTip += L"\r\n";
							}
						}

						{
							CComVariant vAcqDate;
							hr = pImage->m_pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vAcqDate);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							COleDateTime date(vAcqDate);
							SYSTEMTIME stAcqDate;
							date.GetAsSystemTime(stAcqDate);

							SystemTimeToTzSpecificLocalTime(NULL, &stAcqDate, &stAcqDate);

							WCHAR szDate[50], szTime[50];
							GetDateFormat(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &stAcqDate, NULL, szDate, _countof(szDate));
							GetTimeFormat(LOCALE_USER_DEFAULT, 0, &stAcqDate, NULL, szTime, _countof(szTime));

							strTip += L"Date Acquired:\t\t";
							strTip += szDate;
							strTip += L" ";
							strTip += szTime;
							strTip += L"\r\n";
						}

						{
							CComVariant vDate;
							hr = pImage->m_pImage->GetProperty(VSI_PROP_IMAGE_MODIFIED_DATE, &vDate);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							COleDateTime date(vDate);
							SYSTEMTIME stDate;
							date.GetAsSystemTime(stDate);

							SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

							WCHAR szDate[50], szTime[50];
							GetDateFormat(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &stDate, NULL, szDate, _countof(szDate));
							GetTimeFormat(LOCALE_USER_DEFAULT, 0, &stDate, NULL, szTime, _countof(szTime));

							strTip += L"Data Modified:\t\t";
							strTip += szDate;
							strTip += L" ";
							strTip += szTime;
							strTip += L"\r\n";
						}

						{
							CComVariant vVersion;
							hr = pImage->m_pImage->GetProperty(VSI_PROP_IMAGE_VERSION, &vVersion);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (NULL != *V_BSTR(&vVersion)))
							{
								strTip += L"Version:\t\t";
								strTip += V_BSTR(&vVersion);
								strTip += L"\r\n";
							}
						}

						{
							CComVariant vVersion;
							hr = pImage->m_pImage->GetProperty(VSI_PROP_IMAGE_VERSION_CREATED, &vVersion);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (NULL != *V_BSTR(&vVersion)))
							{
								strTip += L"Version Created:\t";
								strTip += V_BSTR(&vVersion);
								strTip += L"\r\n";
							}
						}

						{
							CComVariant vVersion;
							hr = pImage->m_pImage->GetProperty(VSI_PROP_IMAGE_VERSION_MODIFIED, &vVersion);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (NULL != *V_BSTR(&vVersion)))
							{
								strTip += L"Version Modified:\t";
								strTip += V_BSTR(&vVersion);
								strTip += L"\r\n";
							}
						}

						{
							CComVariant vVersion;
							hr = pImage->m_pImage->GetProperty(VSI_PROP_IMAGE_VERSION_REQUIRED, &vVersion);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if ((S_OK == hr) && (NULL != *V_BSTR(&vVersion)))
							{
								strTip += L"Version Required:\t";
								strTip += V_BSTR(&vVersion);
								strTip += L"\r\n";
							}
						}

						{
							CComHeapPtr<OLECHAR> pszPath;
							hr = pImage->m_pImage->GetDataPath(&pszPath);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							PathRemoveExtension(pszPath);

							CString strPath(pszPath);
							strPath += L"*";

							PathRemoveFileSpec(pszPath);

							__int64 iSize = 0;

							CVsiFileIterator iter(strPath);

							LPCWSTR pszFile;
							while (NULL != (pszFile = iter.Next()))
							{
								if (!iter.IsDirectory())
								{
									CString strFilePath(pszPath);
									strFilePath += L"\\";
									strFilePath += pszFile;

									LARGE_INTEGER liSize;
									hr = VsiGetFileSize(strFilePath, &liSize);
									if (S_OK == hr)
									{
										iSize += liSize.QuadPart;
									}
								}
							}

							CString strSize;

							if (iSize < 1024)
							{
								strSize.Format(L"%d B", static_cast<int>(iSize));
							}
							else if (iSize < (1024 * 1024))
							{
								iSize /= 1024;
								strSize.Format(L"%d KB", static_cast<int>(iSize));
							}
							else if (iSize < (1024 * 1024 * 1024))
							{
								iSize /= 1024 * 1024;
								strSize.Format(L"%d MB", static_cast<int>(iSize));
							}
							else
							{
								iSize /= 1024 * 1024 * 1024;
								strSize.Format(L"%d GB", static_cast<int>(iSize));
							}

							strTip += L"Size:\t\t\t";
							strTip += strSize;
							strTip += L"\r\n";
						}

						lstrcpyn(pgit->pszText, strTip, pgit->cchTextMax);
					}
					break;
				}
			}
		}
	}
	VSI_CATCH(hr);
  
	return 0;
}

LRESULT CVsiDataListWnd::OnClick(int, LPNMHDR pnmh, BOOL& bHandled)
{
	LPNMITEMACTIVATE pnmia = (LPNMITEMACTIVATE)pnmh;

	if (pnmia->iItem < 0)  // Clicks outside the filled area
		return 0;

	VSI_DATA_LIST_COL col = m_mapIndexToCol.find(pnmia->iSubItem)->second;

	if (VSI_DATA_LIST_COL_NAME == col)
	{
		CVsiItem *pItem = m_vecItem[pnmia->iItem];

		// Hit test
		LVHITTESTINFO lvhti = { 0 };
		lvhti.pt = pnmia->ptAction;
		ListView_HitTest(pnmh->hwndFrom, &lvhti);

		if (VSI_DATA_STATE_IMAGE_ERROR == pItem->m_iStateImage ||
			VSI_DATA_STATE_IMAGE_WARNING == pItem->m_iStateImage)
		{
			bHandled = FALSE;
		}
		else
		{
			CRect rcHitTest;
			ListView_GetItemRect(*this, lvhti.iItem, &rcHitTest, LVIR_ICON);
			rcHitTest.left -= 20;

			if (rcHitTest.PtInRect(pnmia->ptAction))
			{
				int iEnsureVisibleLastItem = -1;
		
				if (pItem->m_bExpanded)
				{
					pItem->Collapse();
				}
				else
				{
					CWaitCursor wait;

					pItem->Expand(pItem->m_bSelected);
				
					iEnsureVisibleLastItem = pnmia->iItem + pItem->GetVisibleChildrenCount();

					// Analyze the list of created items
					UpdateChildrenStatus(pItem, TRUE);
				}

				UpdateList();

				if (iEnsureVisibleLastItem >= 0)
				{
					// Make sure the last child is visible
					ListView_EnsureVisible(pnmh->hwndFrom, iEnsureVisibleLastItem, FALSE);
				
					// Make sure the expanded item is visible again
					ListView_EnsureVisible(pnmh->hwndFrom, pnmia->iItem, FALSE);
				}
			}
		}

		// Common Controls 6.0
		// Clicking on the state icon doesn't automatically select the item
		// Manually select it here
		if ((lvhti.flags != LVHT_ONITEM) && (lvhti.flags & LVHT_ONITEMSTATEICON))
		{
			// Ignores if SHIFT or CTRL key is down (too complex to handle properly)
			if (GetKeyState(VK_SHIFT) >= 0 && GetKeyState(VK_CONTROL) >= 0)
			{
				ListView_SetItemState(pnmh->hwndFrom, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
				ListView_SetItemState(pnmh->hwndFrom, pnmia->iItem,
					LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
			}
		}
	}

	NMHDR nmhdr = {
		*this,
		m_wIdCtl,
		(UINT)NM_VSI_SELECTION_CHANGED
	};

	GetParent().SendMessage(WM_NOTIFY, m_wIdCtl, (LPARAM)&nmhdr);

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiDataListWnd::OnCustomDraw(int, LPNMHDR pnmh, BOOL&)
{
	LPNMLVCUSTOMDRAW plvcd = (LPNMLVCUSTOMDRAW)pnmh;

	switch (plvcd->nmcd.dwDrawStage)
	{
	case CDDS_PREPAINT:
		{
			if (m_osvi.dwMajorVersion >= 6)
			{
				// Draw grid lines
				CRect rc, rcHeader;
				GetClientRect(&rc);

				CWindow wndHeader(ListView_GetHeader(plvcd->nmcd.hdr.hwndFrom));

				wndHeader.GetWindowRect(&rcHeader);
				rc.top += rcHeader.Height();

				int iItems = ListView_GetItemCount(plvcd->nmcd.hdr.hwndFrom);
				if (0 <iItems)
				{
					CRect rcLast;
					ListView_GetItemRect(plvcd->nmcd.hdr.hwndFrom, iItems - 1, &rcLast, LVIR_BOUNDS);

					if (rcLast.bottom > rc.top && rcLast.bottom < rc.bottom)
					{
						rc.top = rcLast.bottom;
					}
					else
					{
						rc.top = rc.bottom;
					}
				}
				
				if (rc.bottom != rc.top)
				{
					int iColumnCount = Header_GetItemCount(wndHeader);

					std::unique_ptr<int[]> piOrder(new int[iColumnCount]);

					ListView_GetColumnOrderArray(plvcd->nmcd.hdr.hwndFrom, iColumnCount, piOrder.get());

					POINT pt;
					ListView_GetOrigin(plvcd->nmcd.hdr.hwndFrom, &pt);
					int iX = -pt.x - 1;

					HPEN hpen = CreatePen(PS_SOLID, 0, VSI_COLOR_L2GRAY);
					HPEN hpenOld = (HPEN)SelectObject(plvcd->nmcd.hdc, hpen);

					LVCOLUMN lvc = { LVCF_WIDTH };

					for (int i = 0; i < iColumnCount; ++i)
					{
						BOOL bRet = ListView_GetColumn(
							plvcd->nmcd.hdr.hwndFrom,
							piOrder[i],
							&lvc);
						if (!bRet)
						{
							break;
						}
				
						iX += lvc.cx;

						MoveToEx(plvcd->nmcd.hdc, iX, rc.top, NULL);
						LineTo(plvcd->nmcd.hdc, iX, rc.bottom);
					}

					SelectObject(plvcd->nmcd.hdc, hpenOld);
					DeleteObject(hpen);
				}
			}

			if (0 == ListView_GetItemCount(plvcd->nmcd.hdr.hwndFrom) &&
				!m_strEmptyMsg.IsEmpty())
			{
				CRect rc;
				GetEmptyMessageRect(&rc);

				UINT uFormat = DT_NOPREFIX | DT_SINGLELINE | DT_WORD_ELLIPSIS;

				::SetTextColor(plvcd->nmcd.hdc, VSI_COLOR_D0GRAY);
				SetBkMode(plvcd->nmcd.hdc, TRANSPARENT);

				CString strProgress;
				m_FilterProgress.GetText(strProgress);

				DrawText(plvcd->nmcd.hdc, m_strEmptyMsg + strProgress, -1, &rc, uFormat);
			}
		}
		return CDRF_NOTIFYITEMDRAW;

	case CDDS_ITEMPREPAINT:
		{
			if (0 != plvcd->nmcd.uItemState)
			{
				// Draw background
				CRect rc;
				ListView_GetItemRect(
					plvcd->nmcd.hdr.hwndFrom,
					plvcd->nmcd.dwItemSpec,
					&rc,
					LVIR_SELECTBOUNDS);

				// Get state - not sure why plvcd->nmcd.uItemState is not correct
				UINT uItemState = ListView_GetItemState(
					plvcd->nmcd.hdr.hwndFrom,
					plvcd->nmcd.dwItemSpec,
					LVIS_FOCUSED | LVIS_SELECTED);

				HBRUSH hbrushOld = (HBRUSH)SelectObject(plvcd->nmcd.hdc, GetStockObject(DC_BRUSH));
				HPEN hpenOld = (HPEN)SelectObject(plvcd->nmcd.hdc, GetStockObject(DC_PEN));

				if (LVIS_SELECTED & uItemState)
				{
					if (*this == GetFocus())
					{
						SetDCBrushColor(plvcd->nmcd.hdc, VSI_COLOR_ITEM_FOCUSED_SELECTED);
						SetDCPenColor(
							plvcd->nmcd.hdc,
							(LVIS_FOCUSED & uItemState) ?
								VSI_COLOR_ITEM_FOCUSED_FOCUSED : VSI_COLOR_ITEM_FOCUSED_SELECTED);
					}
					else
					{
						SetDCBrushColor(plvcd->nmcd.hdc, VSI_COLOR_ITEM_SELECTED);
						SetDCPenColor(
							plvcd->nmcd.hdc,
							(LVIS_FOCUSED & uItemState) ? VSI_COLOR_ITEM_FOCUSED : VSI_COLOR_ITEM_SELECTED);
					}
				}
				else
				{
					COLORREF rgbBkg = ListView_GetBkColor(plvcd->nmcd.hdr.hwndFrom);

					if (plvcd->nmcd.dwItemSpec & 1)
					{
						rgbBkg = VSI_COLOR_LIST_ODD_ITEM;
					}
					else
					{
						rgbBkg = ListView_GetBkColor(plvcd->nmcd.hdr.hwndFrom);
					}
					
					if (m_iLastHotItem == (int)plvcd->nmcd.dwItemSpec)
					{
						rgbBkg = VsiThemeGetDisabledColor(VSI_COLOR_ITEM_SELECTED, rgbBkg);
					}

					SetDCBrushColor(plvcd->nmcd.hdc, rgbBkg);

					if (LVIS_FOCUSED & uItemState)
					{
						if (*this == GetFocus())
							SetDCPenColor(plvcd->nmcd.hdc, VSI_COLOR_ITEM_FOCUSED_FOCUSED);
						else
							SetDCPenColor(plvcd->nmcd.hdc, VSI_COLOR_ITEM_FOCUSED);
					}
					else
					{
						SetDCPenColor(plvcd->nmcd.hdc, rgbBkg);
					}
				}

				Rectangle(plvcd->nmcd.hdc, rc.left + 1, rc.top, rc.right - 1, rc.bottom);

				// Draw grid lines
				if (m_osvi.dwMajorVersion >= 6)
				{
					if (LVIS_FOCUSED & uItemState)
					{
						++rc.top;
						--rc.bottom;
					}

					if (LVIS_SELECTED & uItemState)
					{
						SetDCPenColor(plvcd->nmcd.hdc, VSI_COLOR_L2GRAY);
					}
					else
					{
						SetDCPenColor(plvcd->nmcd.hdc, VSI_COLOR_L2GRAY);
					}

					CWindow wndHeader(ListView_GetHeader(plvcd->nmcd.hdr.hwndFrom));
					int iColumnCount = Header_GetItemCount(wndHeader);

					std::unique_ptr<int[]> piOrder(new int[iColumnCount]);

					ListView_GetColumnOrderArray(plvcd->nmcd.hdr.hwndFrom, iColumnCount, piOrder.get());

					POINT pt;
					ListView_GetOrigin(plvcd->nmcd.hdr.hwndFrom, &pt);
					int iX = -pt.x - 1;
					
					LVCOLUMN lvc = { LVCF_WIDTH };

					for (int i = 0; i < iColumnCount; ++i)
					{
						BOOL bRet = ListView_GetColumn(
							plvcd->nmcd.hdr.hwndFrom,
							piOrder[i],
							&lvc);
						if (!bRet)
							break;
					
						iX += lvc.cx;

						MoveToEx(plvcd->nmcd.hdc, iX, rc.top, NULL);
						LineTo(plvcd->nmcd.hdc, iX, rc.bottom);
					}
				}

				SelectObject(plvcd->nmcd.hdc, hpenOld);
				SelectObject(plvcd->nmcd.hdc, hbrushOld);
			}
		}
		return CDRF_NOTIFYSUBITEMDRAW;

	case (CDDS_SUBITEM | CDDS_ITEMPREPAINT):
		{
			CVsiItem *pItem = m_vecItem[plvcd->nmcd.dwItemSpec];

			CRect rc;

			WCHAR szText[500];
			LVITEM lvi = { LVIF_IMAGE | LVIF_STATE | LVIF_TEXT | LVIF_INDENT };
			lvi.iItem = (int)plvcd->nmcd.dwItemSpec;
			lvi.iSubItem = plvcd->iSubItem;
			lvi.stateMask = INDEXTOSTATEIMAGEMASK(0x000F);
			lvi.pszText = szText;
			lvi.cchTextMax = _countof(szText);

			ListView_GetItem(plvcd->nmcd.hdr.hwndFrom, &lvi);

			// Vista version
			if (m_osvi.dwMajorVersion >= 6)
			{
				rc = plvcd->nmcd.rc;
			}
			else
			{
				ListView_GetSubItemRect(
					plvcd->nmcd.hdr.hwndFrom,
					(int)plvcd->nmcd.dwItemSpec,
					plvcd->iSubItem,
					LVIR_BOUNDS,
					&rc);

				rc.right = rc.left + plvcd->nmcd.rc.right - plvcd->nmcd.rc.left;

				if (!rc.IsRectEmpty())
				{
					if (0 == plvcd->iSubItem)
					{
						rc.left += lvi.iIndent * 15 + 18;
					}
				}
			}

			if (!rc.IsRectEmpty())
			{
				if (0 == plvcd->iSubItem && lvi.state > 0)
				{
					int iState = ((lvi.state & LVIS_STATEIMAGEMASK) >> 12) - 1;
					if (0 <= iState)
					{
						if (VSI_STATE_IMAGE_WIDTH < rc.Width())
						{
							if (iState < VSI_DATA_STATE_IMAGE_CLOSE_HOT)
							{
								if ((int)plvcd->nmcd.dwItemSpec == m_iLastHotState)
								{
									++iState;
								}
							}

							int iTop = rc.top + (rc.Height() - 16) / 2;
							ImageList_Draw(
								m_ilState,
								iState,
								plvcd->nmcd.hdc,
								rc.left - VSI_STATE_IMAGE_WIDTH - 5,
								iTop,
								ILD_TRANSPARENT);
						}
					}
				}

				SIZE sizeImage;
				m_ilSmall.GetIconSize(sizeImage);

				if (0 <= lvi.iImage)
				{
					if (0 == plvcd->iSubItem)
					{
						rc.left -= 4;
					}

					if (sizeImage.cx < rc.Width())
					{
						ImageList_Draw(
							m_ilSmall,
							lvi.iImage,
							plvcd->nmcd.hdc,
							rc.left,
							rc.top + (rc.Height() - sizeImage.cy) / 2,
							ILD_TRANSPARENT);
					}

					if (0 == plvcd->iSubItem)
					{
						rc.left += 2;
					}
				}

				if (0 == plvcd->iSubItem)
				{
					rc.left += sizeImage.cx - 6;
				}

				if (VSI_DATA_LIST_COL_LOCKED == m_mapIndexToCol[plvcd->iSubItem])
				{
					*szText = 0;
				}
			
				if (0 != *szText)
				{
					if (12 < rc.Width())
					{
						rc.left += 6;
						rc.right -= 6;

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

						bool bFontNormal = true;

						bool bFilter(false);
						if (0 < m_vecFilterJobs.size())
						{
							auto & job = m_vecFilterJobs.back();

							for (auto i = job.m_vecFilters.begin();
								i != job.m_vecFilters.end() && !bFilter;
								++i)
							{
								bFilter = pItem->HasText(szText, *i);
							}
						}

						if (pItem->m_iStateText > 0)
						{
							::SetTextColor(plvcd->nmcd.hdc, m_prgbState[pItem->m_iStateText]);
					
							bFontNormal = (VSI_DATA_STATE_IMAGE_NORMAL == pItem->m_iStateText) ||
								(VSI_DATA_STATE_IMAGE_ACTIVE_NORMAL == pItem->m_iStateText) ||
								(VSI_DATA_STATE_IMAGE_INACTIVE_NORMAL == pItem->m_iStateText);
						}
						else
						{
							::SetTextColor(plvcd->nmcd.hdc, m_prgbState[pItem->m_iStateImage]);
						}

						if (bFilter)
						{
							// Override text color
							::SetTextColor(plvcd->nmcd.hdc, m_prgbState[VSI_DATA_STATE_IMAGE_FILTER]);
						}
				
						if (NULL != m_hfontNormal && NULL != m_hfontReviewed)
						{
							SelectObject(plvcd->nmcd.hdc, bFontNormal ? m_hfontNormal : m_hfontReviewed);
						}

						SetBkMode(plvcd->nmcd.hdc, TRANSPARENT);
						DrawText(plvcd->nmcd.hdc, szText, -1, &rc, uFormat);
					}
				}
			}
		}
		return CDRF_SKIPDEFAULT;
	}

	return CDRF_DODEFAULT;
}

void CVsiDataListWnd::GetEmptyMessageRect(LPRECT prc)
{
	CRect rcHeader;
	GetClientRect(prc);
				
	CWindow wndHeader(ListView_GetHeader(*this));
	wndHeader.GetWindowRect(&rcHeader);
	prc->top += rcHeader.Height() + 2;
	prc->left += 4;
}

void CVsiDataListWnd::UpdateSortIndicator(BOOL bShow)
{
	if (VSI_DATA_LIST_COL_LOCKED == m_sort.m_sortCol)
	{
		return;
	}

	CWindow wndHeader(ListView_GetHeader(*this));

	auto iter = m_mapColToIndex.find(m_sort.m_sortCol);

	if (iter != m_mapColToIndex.end())
	{
		int iIndex = iter->second;

		HDITEM hdi = { HDI_FORMAT };
		Header_GetItem(wndHeader, iIndex, &hdi);

		if (bShow)
		{
			hdi.fmt |= m_sort.m_bDescending ? HDF_SORTDOWN : HDF_SORTUP;
			Header_SetItem(wndHeader, iIndex, &hdi);
		}
		else
		{
			hdi.fmt &= ~(HDF_SORTDOWN | HDF_SORTUP);
			Header_SetItem(wndHeader, iIndex, &hdi);
		}
	}
}

void CVsiDataListWnd::UpdateViewVector(CVsiListItem *pList)
{
	pList->sort(
		[](CVsiDataListWnd::CVsiItem *px, CVsiDataListWnd::CVsiItem *py)->bool
		{
			return *px < *py;
		}
	);

	for (auto iter = pList->begin(); iter != pList->end(); ++iter)
	{
		CVsiItem *pItem = *iter;

		if (pItem->m_bVisible)
		{
			m_vecItem.push_back(pItem);
		}

		if (pItem->m_bExpanded)
		{
			UpdateViewVector(&pItem->m_listChild);
		}
	}
}

void CVsiDataListWnd::Fill(IVsiEnumStudy *pEnum, int iLevelMax)
{
	Clear();

	m_iLevelMax = iLevelMax;

	CComPtr<IVsiStudy> pStudy;
	while (pEnum->Next(1, &pStudy, NULL) == S_OK)
	{
		AddStudy(pStudy);
		pStudy.Release();
	}

	UpdateList();
}

void CVsiDataListWnd::Fill(IVsiEnumSeries *pEnum, int iLevelMax)
{
	Clear();

	m_iLevelMax = iLevelMax;

	CComPtr<IVsiSeries> pSeries;
	while (pEnum->Next(1, &pSeries, NULL) == S_OK)
	{
		AddSeries(pSeries);
		pSeries.Release();
	}

	UpdateList();
}

void CVsiDataListWnd::Fill(IVsiEnumImage *pEnum)
{
	Clear();

	m_iLevelMax = 1;

	CComPtr<IVsiImage> pImage;
	while (pEnum->Next(1, &pImage, NULL) == S_OK)
	{
		AddImage(pImage);
		pImage.Release();
	}

	UpdateList();
}

void CVsiDataListWnd::AddStudy(IVsiStudy *pStudy, BOOL bSelect) throw(...)
{
	VSI_CHECK_ARG(pStudy, VSI_LOG_ERROR, NULL);

	CVsiStudy *p = new CVsiStudy;
	VSI_CHECK_MEM(p, VSI_LOG_ERROR, NULL);

	VSI_STUDY_ERROR dwErrorCode(VSI_STUDY_ERROR_NONE);
	pStudy->GetErrorCode(&dwErrorCode);
	if (VSI_STUDY_ERROR_VERSION_INCOMPATIBLE & dwErrorCode)
		p->m_iStateImage = VSI_DATA_STATE_IMAGE_WARNING;
	else if (VSI_STUDY_ERROR_LOAD_XML & dwErrorCode)
		p->m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
	p->m_pRoot = this;
	p->m_pStudy = pStudy;
	p->m_pSort = &m_sort;
	if (bSelect)
	{
		p->m_bSelected = TRUE;
	}

	if (0 < m_vecFilterJobs.size())
	{
		auto const & job = m_vecFilterJobs.back();
		Filter(p, job.m_vecFilters, job.m_filterType);
	}

	m_listRoot.push_back(p);

	UpdateItemStatus(p);
}

void CVsiDataListWnd::AddSeries(IVsiSeries *pSeries, BOOL bSelect) throw(...)
{
	VSI_CHECK_ARG(pSeries, VSI_LOG_ERROR, NULL);

	CVsiSeries *p = new CVsiSeries;
	VSI_CHECK_MEM(p, VSI_LOG_ERROR, NULL);

	VSI_SERIES_ERROR dwErrorCode(VSI_SERIES_ERROR_NONE);
	pSeries->GetErrorCode(&dwErrorCode);
	if (VSI_SERIES_ERROR_VERSION_INCOMPATIBLE & dwErrorCode)
		p->m_iStateImage = VSI_DATA_STATE_IMAGE_WARNING;
	else if (VSI_SERIES_ERROR_LOAD_XML & dwErrorCode)
		p->m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
	p->m_pRoot = this;
	p->m_pSeries = pSeries;
	p->m_pSort = &m_sort;
	if (bSelect)
	{
		p->m_bSelected = TRUE;
	}

	if (0 < m_vecFilterJobs.size())
	{
		auto const & job = m_vecFilterJobs.back();
		Filter(p, job.m_vecFilters, job.m_filterType);
	}

	m_listRoot.push_back(p);

	UpdateItemStatus(p);
}

void CVsiDataListWnd::AddImage(IVsiImage *pImage, BOOL bSelect) throw(...)
{
	VSI_CHECK_ARG(pImage, VSI_LOG_ERROR, NULL);

	CVsiImage *p = new CVsiImage;
	VSI_CHECK_MEM(p, VSI_LOG_ERROR, NULL);

	VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_NONE);
	pImage->GetErrorCode(&dwErrorCode);
	if (VSI_IMAGE_ERROR_VERSION_INCOMPATIBLE & dwErrorCode)
		p->m_iStateImage = VSI_DATA_STATE_IMAGE_WARNING;
	else if (VSI_IMAGE_ERROR_NONE != dwErrorCode)
		p->m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
	p->m_pRoot = this;
	p->m_pImage = pImage;
	p->m_pSort = &m_sort;
	if (bSelect)
	{
		p->m_bSelected = TRUE;
	}

	if (0 < m_vecFilterJobs.size())
	{
		auto const & job = m_vecFilterJobs.back();
		Filter(p, job.m_vecFilters, job.m_filterType);
	}

	m_listRoot.push_back(p);

	UpdateItemStatus(p);
}

HRESULT CVsiDataListWnd::GetSelectedStudies(
	IVsiEnumStudy **ppEnumStudies, int *piCount, CVsiListItem *pSetSelected)
{
	HRESULT hr = S_OK;

	typedef CComObject<
		CComEnum<
			IVsiEnumStudy,
			&__uuidof(IVsiEnumStudy),
			IVsiStudy*, _CopyInterface<IVsiStudy>
		>
	> CVsiEnumStudy;

	IVsiStudy **ppStudy = NULL;
	CVsiEnumStudy *pEnum = NULL;
	int i = 0;

	try
	{
		VSI_CHECK_ARG(ppEnumStudies, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(piCount, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(pSetSelected, VSI_LOG_ERROR, NULL);

		std::list<IVsiStudy*> listStudy;
		CVsiItem *pItem = NULL;

		for (auto iter = pSetSelected->begin(); iter != pSetSelected->end(); ++iter)
		{
			pItem = *iter;
			CVsiStudy *pStudy = dynamic_cast<CVsiStudy*>(pItem);
			VSI_VERIFY(nullptr != pStudy, VSI_LOG_ERROR, NULL);

			listStudy.push_back(pStudy->m_pStudy);
		}

		*piCount = (int)listStudy.size();

		if ((NULL != ppEnumStudies) && (*piCount > 0))
		{
			ppStudy = new IVsiStudy*[*piCount];
			VSI_CHECK_MEM(ppStudy, VSI_LOG_ERROR, NULL);

			for (auto iterStudy = listStudy.begin();
				iterStudy != listStudy.end();
				++iterStudy)
			{
				IVsiStudy *pStudy = *iterStudy;
				pStudy->AddRef();

				ppStudy[i++] = pStudy;
			}

			pEnum = new CVsiEnumStudy;
			VSI_CHECK_MEM(pEnum, VSI_LOG_ERROR, NULL);

			hr = pEnum->Init(ppStudy, ppStudy + i, NULL, AtlFlagTakeOwnership);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			ppStudy = NULL;  // pEnum is the owner

			hr = pEnum->QueryInterface(__uuidof(IVsiEnumStudy),	(void**)ppEnumStudies);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			pEnum = NULL;
		}
	}
	VSI_CATCH(hr);

	if (NULL != pEnum) delete pEnum;
	if (NULL != ppStudy)
	{
		for (; i > 0; i--)
		{
			if (NULL != ppStudy[i])
			{
				ppStudy[i]->Release();
			}
		}

		delete [] ppStudy;
	}

	return hr;
}

HRESULT CVsiDataListWnd::GetSelectedSeries(
	IVsiEnumSeries **ppEnumSeries, int *piCount, CVsiListItem *pSetSelected)
{
	HRESULT hr = S_OK;

	typedef CComObject<
		CComEnum<
			IVsiEnumSeries,
			&__uuidof(IVsiEnumSeries),
			IVsiSeries*, _CopyInterface<IVsiSeries>
		>
	> CVsiEnumSeries;

	IVsiSeries **ppSeries = NULL;
	CVsiEnumSeries *pEnum = NULL;
	int i = 0;

	try
	{
		VSI_CHECK_ARG(ppEnumSeries, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(piCount, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(pSetSelected, VSI_LOG_ERROR, NULL);

		std::list<IVsiSeries*> listSeries;
		CVsiItem *pItem = NULL;

		for (auto iter = pSetSelected->begin(); iter != pSetSelected->end(); ++iter)
		{
			pItem = *iter;
			CVsiSeries* pSeries = dynamic_cast<CVsiSeries*>(pItem);
			VSI_VERIFY(nullptr != pSeries, VSI_LOG_ERROR, NULL);

			listSeries.push_back(pSeries->m_pSeries);
		}

		*piCount = (int)listSeries.size();

		if ((NULL != ppEnumSeries) && (*piCount > 0))
		{
			ppSeries = new IVsiSeries*[*piCount];
			VSI_CHECK_MEM(ppSeries, VSI_LOG_ERROR, NULL);

			for (auto iterSeries = listSeries.begin();
				iterSeries != listSeries.end();
				++iterSeries)
			{
				IVsiSeries *pSeries = *iterSeries;
				pSeries->AddRef();

				ppSeries[i++] = pSeries;
			}

			pEnum = new CVsiEnumSeries;
			VSI_CHECK_MEM(pEnum, VSI_LOG_ERROR, NULL);

			hr = pEnum->Init(ppSeries, ppSeries + i, NULL, AtlFlagTakeOwnership);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			ppSeries = NULL;  // pEnum is the owner

			hr = pEnum->QueryInterface(__uuidof(IVsiEnumSeries),	(void**)ppEnumSeries);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			pEnum = NULL;
		}
	}
	VSI_CATCH(hr);

	if (NULL != pEnum) delete pEnum;
	if (NULL != ppSeries)
	{
		for (; i > 0; i--)
		{
			if (NULL != ppSeries[i])
				ppSeries[i]->Release();
		}

		delete [] ppSeries;
	}

	return hr;
}

HRESULT CVsiDataListWnd::GetSelectedImages(
	IVsiEnumImage **ppEnumImage, int *piCount, CVsiListItem *pSetSelected)
{
	HRESULT hr = S_OK;

	typedef CComObject<
		CComEnum<
			IVsiEnumImage,
			&__uuidof(IVsiEnumImage),
			IVsiImage*, _CopyInterface<IVsiImage>
		>
	> CVsiEnumImage;

	std::unique_ptr<IVsiImage*[]> ppImage;
	std::unique_ptr<CVsiEnumImage> pEnum;
	int i = 0;

	try
	{
		VSI_CHECK_ARG(ppEnumImage, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(piCount, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(pSetSelected, VSI_LOG_ERROR, NULL);

		std::list<IVsiImage*> listImage;
		CVsiItem *pItem = NULL;

		for (auto iter = pSetSelected->begin(); iter != pSetSelected->end(); ++iter)
		{
			pItem = *iter;
			CVsiImage *pImage = dynamic_cast<CVsiImage*>(pItem);
			VSI_VERIFY(nullptr != pImage, VSI_LOG_ERROR, NULL);

			listImage.push_back(pImage->m_pImage);
		}

		*piCount = (int)listImage.size();

		if ((NULL != ppEnumImage) && (*piCount > 0))
		{
			ppImage.reset(new IVsiImage*[*piCount]);

			for (auto iterImage = listImage.begin();
				iterImage != listImage.end();
				++iterImage)
			{
				IVsiImage *pImage = *iterImage;
				pImage->AddRef();

				ppImage[i++] = pImage;
			}

			pEnum.reset(new CVsiEnumImage);

			hr = pEnum->Init(ppImage.get(), ppImage.get() + i, NULL, AtlFlagTakeOwnership);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			ppImage.release();  // pEnum is the owner

			hr = pEnum->QueryInterface(__uuidof(IVsiEnumImage),	(void**)ppEnumImage);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			pEnum.release();
		}
	}
	VSI_CATCH(hr);

	if (ppImage)
	{
		for (; i > 0; i--)
		{
			if (NULL != ppImage[i])
			{
				ppImage[i]->Release();
			}
		}
	}

	return hr;
}

void CVsiDataListWnd::UpdateItemStatus(CVsiItem* pItem, BOOL bRecurChildren) throw(...)
{
	if (0 == (m_dwFlags & VSI_DATA_LIST_FLAG_ITEM_STATUS_CALLBACK))
	{
		return;
	}

	HRESULT hr = S_OK;
	CComHeapPtr<OLECHAR> pszFile;

	NM_DL_ITEM nt;
	nt.hdr.code = (UINT)NM_VSI_DL_GET_ITEM_STATUS;
	nt.hdr.hwndFrom = *this;
	nt.hdr.idFrom = m_wIdCtl;

	switch (pItem->m_type)
	{
	case VSI_DATA_LIST_STUDY:
		{
			CVsiStudy *pStudy = dynamic_cast<CVsiStudy*>(pItem);
			VSI_VERIFY(nullptr != pStudy, VSI_LOG_ERROR, NULL);

			// Notify parent that Series are being added
			pszFile.Free();
			hr = pStudy->m_pStudy->GetDataPath(&pszFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			nt.pUnkItem = pStudy->m_pStudy;
		}
		break;

	case VSI_DATA_LIST_SERIES:
		{
			CVsiSeries *pSeries = dynamic_cast<CVsiSeries*>(pItem);
			VSI_VERIFY(nullptr != pSeries, VSI_LOG_ERROR, NULL);

			// Notify parent that Series are being added
			pszFile.Free();
			hr = pSeries->m_pSeries->GetDataPath(&pszFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			nt.pUnkItem = pSeries->m_pSeries;
		}
		break;

	case VSI_DATA_LIST_IMAGE:
		{
			CVsiImage *pImage = dynamic_cast<CVsiImage*>(pItem);
			VSI_VERIFY(nullptr != pImage, VSI_LOG_ERROR, NULL);

			// Notify parent that Image are being added
			pszFile.Free();
			hr = pImage->m_pImage->GetDataPath(&pszFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			nt.pUnkItem = pImage->m_pImage;
		}
		break;

	default:
		_ASSERT(0);
	}

	nt.type = pItem->m_type;
	nt.pszItemPath = pszFile;

	LRESULT lRes = GetParent().SendMessage(WM_NOTIFY, nt.hdr.idFrom, (LPARAM)&nt);

	// Sets status
	if (VSI_DATA_LIST_ITEM_STATUS_HIDDEN & lRes)
	{
		pItem->m_bVisible = FALSE;
	}

	if (VSI_DATA_STATE_IMAGE_ERROR != pItem->m_iStateImage &&
		VSI_DATA_STATE_IMAGE_WARNING != pItem->m_iStateImage)
	{
		pItem->m_iStateImage = 0;
		pItem->m_iStateText = 0;
	}

	if (VSI_DATA_LIST_ITEM_STATUS_ACTIVE_SINGLE & lRes)
	{
		pItem->m_iStateImage = VSI_DATA_STATE_IMAGE_SINGLE;
	}
	else
	{
		if (VSI_DATA_LIST_ITEM_STATUS_ACTIVE_LEFT & lRes)
			pItem->m_iStateImage = VSI_DATA_STATE_IMAGE_LEFT;
		else if (VSI_DATA_LIST_ITEM_STATUS_ACTIVE_RIGHT & lRes)
			pItem->m_iStateImage = VSI_DATA_STATE_IMAGE_RIGHT;
	}

	if (VSI_DATA_LIST_ITEM_STATUS_ACTIVE & lRes)
	{
		if (VSI_DATA_LIST_ITEM_STATUS_REVIEWED & lRes)
			pItem->m_iStateText = VSI_DATA_STATE_IMAGE_ACTIVE_REVIEWED;
		else
			pItem->m_iStateText = VSI_DATA_STATE_IMAGE_ACTIVE_NORMAL;
	}
	else if (VSI_DATA_LIST_ITEM_STATUS_INACTIVE & lRes)
	{
		if (VSI_DATA_LIST_ITEM_STATUS_REVIEWED & lRes)
			pItem->m_iStateText = VSI_DATA_STATE_IMAGE_INACTIVE_REVIEWED;
		else
			pItem->m_iStateText = VSI_DATA_STATE_IMAGE_INACTIVE_NORMAL;
	}
	else if (0 == pItem->m_iStateImage)
	{
		if (VSI_DATA_LIST_ITEM_STATUS_REVIEWED & lRes)
			pItem->m_iStateText = VSI_DATA_STATE_IMAGE_REVIEWED;
		else
			pItem->m_iStateText = VSI_DATA_STATE_IMAGE_NORMAL;
	}

	if (bRecurChildren)
	{
		if (pItem->m_bExpanded)
		{
			for (auto iter = pItem->m_listChild.begin();
				iter != pItem->m_listChild.end();
				++iter)
			{
				CVsiItem *pItemNext = *iter;
				UpdateItemStatus(pItemNext, bRecurChildren);
			}
		}
	}
}

void CVsiDataListWnd::UpdateChildrenStatus(CVsiItem* pItem, BOOL bRecurChildren)
{
	if (0 == (m_dwFlags & VSI_DATA_LIST_FLAG_ITEM_STATUS_CALLBACK))
	{
		return;
	}

	for (auto iter = pItem->m_listChild.begin(); iter != pItem->m_listChild.end(); ++iter)
	{
		CVsiItem *pChild = *iter;

		UpdateItemStatus(pChild, bRecurChildren);
	}
}

void CVsiDataListWnd::UpdateItemLocationMap(
	CVsiMapItemLocation &map,
	LPCWSTR pszParent,
	CVsiListItem &list) throw(...)
{
	HRESULT hr;
	CComHeapPtr<OLECHAR> pszId;

	for (auto iter = list.begin(); iter != list.end(); ++iter)
	{
		CVsiItem *pItem = *iter;

		pszId.Free();
		hr = pItem->GetId(&pszId);
		if (S_OK == hr)
		{
			CString strNs(pszParent);
			if (!strNs.IsEmpty())
			{
				strNs += L"/";
			}
			strNs += pszId;

			CVsiItemLocation loc(&list, iter);
			map[strNs] = loc;

			if (pItem->m_bExpanded)
			{
				UpdateItemLocationMap(map, strNs, pItem->m_listChild);
			}
		}
		else
		{
			// Corrupted item - ignores
		}
	}
}

CVsiDataListWnd::CVsiItem* CVsiDataListWnd::GetItem(
	LPCOLESTR pszNamespace, bool bVisibleOnly, BOOL *pbCancel)
{
	CVsiItem *pItemRet = NULL;
	LPCWSTR pszNext = pszNamespace;
	CString strId;
	CVsiListItem *pList = &m_listRoot;
	int i = 0;

	while ((NULL != pszNext) && (0 != *pszNext))
	{
		LPCWSTR pszSpt = wcschr(pszNext, L'/');
		if (pszSpt != NULL)
		{
			strId.SetString(pszNext, (int)(pszSpt - pszNext));
			pszNext = pszSpt + 1;
		}
		else
		{
			strId = pszNext;
			pszNext = NULL;
		}

		for (auto iter = pList->begin(); iter != pList->end(); ++iter)
		{
			CVsiItem *pItem = *iter;
			CComHeapPtr<OLECHAR> pszId;
			pItem->GetId(&pszId);

			if (pszId != NULL && strId == pszId)
			{
				if (NULL == pszNext)
				{
					// Reach the end
					pItemRet = pItem;
					break;
				}
				else
				{
					// Found a node, expand it
					if (!pItem->m_bExpanded && bVisibleOnly)
						return NULL;

					BOOL bWasExpanded = pItem->m_bExpanded;
					pItem->Expand(FALSE, pbCancel);
					pItem->m_bExpanded = bWasExpanded;

					pList = &pItem->m_listChild;
				}
				break;
			}

			if (pbCancel && *pbCancel)
			{
				break;
			}

			++i;
		}
	}

	return pItemRet;
}

void CVsiDataListWnd::GetNextFilterJob(CVsiFilterJob &job)
{
	// Search all items by default
	VSI_SEARCH_SCOPE scope(VSI_SEARCH_ALL);

	if (0 < m_vecFilterJobs.size())
	{
		auto const &lastJob = m_vecFilterJobs.back();

		if (job.m_filterType == lastJob.m_filterType)
		{
			// Have old and new filters, see whether we can speed up the search using
			// the existing results
			std::vector<CString> vecFiltersOldDiff;
			std::vector<CString> vecFiltersNewDiff;

			// Check is a superset or subset
			for (auto i = lastJob.m_vecFilters.begin(); i != lastJob.m_vecFilters.end(); ++i)
			{
				auto find = std::find(job.m_vecFilters.begin(), job.m_vecFilters.end(), *i);
				if (find == job.m_vecFilters.end())
				{
					vecFiltersOldDiff.push_back(*i);
				}
			}
			for (auto i = job.m_vecFilters.begin(); i != job.m_vecFilters.end(); ++i)
			{
				auto find = std::find(lastJob.m_vecFilters.begin(), lastJob.m_vecFilters.end(), *i);
				if (find == lastJob.m_vecFilters.end())
				{
					vecFiltersNewDiff.push_back(*i);
				}
			}

			if (0 == vecFiltersOldDiff.size() && 1 == vecFiltersNewDiff.size())
			{
				if (VSI_FILTER_TYPE_OR == job.m_filterType)
				{
					scope = VSI_SEARCH_HIDDEN;
				}
				else  // VSI_FILTER_TYPE_AND
				{
					scope = VSI_SEARCH_VISIBLE;
				}
			}
			else if (1 == vecFiltersOldDiff.size() && 0 == vecFiltersNewDiff.size())
			{
				if (VSI_FILTER_TYPE_OR == job.m_filterType)
				{
					scope = VSI_SEARCH_VISIBLE;
				}
				else  // VSI_FILTER_TYPE_AND
				{
					scope = VSI_SEARCH_HIDDEN;
				}
			}

			if (1 == vecFiltersOldDiff.size() && 1 == vecFiltersNewDiff.size())
			{
				// Can only handle 1 difference today
				auto const & i = vecFiltersOldDiff.front();
				auto const & j = vecFiltersNewDiff.front();

				if (NULL != StrStrI(i, j))
				{
					scope = VSI_SEARCH_HIDDEN;
				}
				else if (NULL != StrStrI(j, i))
				{
					scope = VSI_SEARCH_VISIBLE;
				}
			}

			if (VSI_SEARCH_ALL != scope)
			{
				job.m_scope = scope;
				job.m_pItemFilterStop = lastJob.m_pItemFilterStop;

				m_vecFilterJobs.pop_back();
			}
		}
	}

	if (VSI_SEARCH_ALL == scope)
	{
		m_vecFilterJobs.clear();
		job.m_scope = VSI_SEARCH_ALL;
		job.m_pItemFilterStop = nullptr;
	}
}

bool CVsiDataListWnd::Filter(
	CVsiListItem& listItem,
	CVsiFilterJob &job,
	BOOL *pbStop)
{
	bool bFoundAny = false;

	for (auto iter = listItem.begin(); iter != listItem.end(); ++iter)
	{
		CVsiItem *pItem = *iter;

		// Expand if filtering, collapse otherwise
		if (0 < job.m_vecFilters.size())
		{
			// Expand node
			pItem->Expand();

			bool bFound(false);
			
			if (0 < pItem->m_listChild.size())
			{
				bFound = Filter(pItem->m_listChild, job, pbStop);
			}

			if (bFound)
			{
				// If any of the child is visible, the parent must be visible too
				pItem->m_bVisible = TRUE;

				bFoundAny = true;
			}
			else
			{
				// None of the children are visible, search itself
				bool bSearch(true);

				switch (job.m_scope)
				{
				case VSI_SEARCH_HIDDEN:
					bSearch = (FALSE == pItem->m_bVisible);
					break;
				case VSI_SEARCH_VISIBLE:
					bSearch = (FALSE != pItem->m_bVisible);
					break;
				}

				if (bSearch)
				{
					if (Filter(pItem, job.m_vecFilters, job.m_filterType))
					{
						bFoundAny = true;
					}
				}

				if (job.m_pItemFilterStop == pItem)
				{
					// Reach stop marker, get next job
					GetNextFilterJob(job);
				}
			}
		}
		else
		{
			// Not filtering
			pItem->Collapse();

			pItem->m_bVisible = TRUE;
		}

		// Remove selection (collapse may change selection if children have selection)
		pItem->m_bSelected = FALSE;
		pItem->m_bFocused = FALSE;

		if (*pbStop)
		{
			// Stop requested

			if (VSI_SEARCH_ALL == job.m_scope)
			{
				// Search deeper than before, set new stop marker
				job.m_scope = VSI_SEARCH_NONE;

				job.m_pItemFilterStop = pItem;
			}

			break;
		}
	}

	return bFoundAny;
}

bool CVsiDataListWnd::Filter(CVsiItem *pItem, const std::vector<CString> &vecFilters, VSI_FILTER_TYPE type)
{
	bool bFoundAny(false);

	switch (type)
	{
	case VSI_FILTER_TYPE_OR:
		{
			pItem->m_bVisible = FALSE;

			for (auto iter = vecFilters.begin(); iter != vecFilters.end(); ++iter)
			{
				if (pItem->HasText(*iter, true))
				{
					pItem->m_bVisible = TRUE;

					bFoundAny = true;

					break;
				}
			}
		}
		break;
	case VSI_FILTER_TYPE_AND:
		{
			bFoundAny = true;

			pItem->m_bVisible = TRUE;

			for (auto iter = vecFilters.begin(); iter != vecFilters.end(); ++iter)
			{
				if (!pItem->HasText(*iter, true))
				{
					pItem->m_bVisible = FALSE;

					bFoundAny = false;

					break;
				}
			}
		}
		break;
	}

	return bFoundAny;
}

bool CVsiDataListWnd::FilterParent(
	CVsiItem *pItem,
	const std::vector<CString> &vecFilters,
	VSI_FILTER_TYPE type)
{
	bool bVisible(false);

	for (auto iter = pItem->m_listChild.begin();
		iter != pItem->m_listChild.end();
		++iter)
	{
		CVsiItem *pItemTest = *iter;
		if (pItemTest->m_bVisible)
		{
			bVisible = true;
			break;
		}
	}

	if (!bVisible)
	{
		if (!Filter(pItem, vecFilters, type))
		{
			pItem->m_bVisible = false;

			if (NULL != pItem->m_pParent)
			{
				FilterParent(pItem->m_pParent, vecFilters, type);
			}
		}
	}

	return bVisible;
}

void CVsiDataListWnd::ExpandAll(CVsiListItem& listItem, BOOL &bStop)
{
	VSI_LOG_MSG(VSI_LOG_INFO, L"Search - Expand all items begin");

	// Expand all is only used by search (for now) and there is no needed for session link
	// Remove the session link flag temporary to speed things up
	DWORD dwFlagsOld = m_dwFlags;

	if (m_dwFlags & VSI_DATA_LIST_FLAG_NO_SESSION_LINK)
	{
		m_dwFlags &= ~VSI_DATA_LIST_FLAG_NO_SESSION_LINK;
	}

	m_FilterProgress.SetTotal(listItem.size());

	CRect rc;
	GetEmptyMessageRect(&rc);

	Concurrency::structured_task_group tasks;

	auto task = Concurrency::make_task([&]() 
	{
		CVsiDataListWnd *pThis(this);
		BOOL &bStop1(bStop);
		LPCRECT prc = &rc;
		Concurrency::structured_task_group& tasks1(tasks);
		Concurrency::parallel_for_each(listItem.begin(), listItem.end(), [&](CVsiItem *pItem)
		{
			pItem->Expand(FALSE, &bStop1);

			if (0 < pItem->m_listChild.size())
			{
				for (auto i = pItem->m_listChild.begin(); i != pItem->m_listChild.end() && !bStop1; ++i)
				{
					(*i)->Expand(FALSE, &bStop1);
				}
			}

			pThis->m_FilterProgress.IncrementDone();
			pThis->InvalidateRect(prc);

			if (bStop1)
			{
				tasks1.cancel();
			}
		});
	});

	tasks.run_and_wait(task);

	m_FilterProgress.SetTotal(0);

	m_dwFlags = dwFlagsOld;

	VSI_LOG_MSG(VSI_LOG_INFO, L"Search - Expand all items end");
}

