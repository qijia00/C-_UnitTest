/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiDataListWnd.h
**
**	Description:
**		Declaration of the CVsiDataListWnd
**
*******************************************************************************/

#pragma once

#include <ppl.h>
#include <atltime.h>
#include <ATLComTime.h>
#include <VsiWtl.h>
#include <VsiAppModule.h>
#include <VsiModeModule.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiWindow.h>
#include "resource.h"       // main symbols


class ATL_NO_VTABLE CVsiDataListWnd :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiDataListWnd, &CLSID_VsiDataListWnd>,
	public CWindowImpl<CVsiDataListWnd, CVsiWindow>,
	public IVsiDataListWnd
{
	typedef CWindowImpl<CVsiDataListWnd, CVsiWindow> baseClass;

private:
	class CVsiSort;

public:
	class CVsiItem;
	class CVsiStudy;
	class CVsiSeries;
	class CVsiImage;

	enum
	{
		VSI_DATA_STATE_IMAGE_UNCHECKED = 0,				// Image (0 base)
		VSI_DATA_STATE_IMAGE_CHECKED = 1,				// Image
		VSI_DATA_STATE_IMAGE_OPEN = 3,					// State (1 base)
		VSI_DATA_STATE_IMAGE_OPEN_HOT = 4,				// State
		VSI_DATA_STATE_IMAGE_CLOSE = 5,					// State
		VSI_DATA_STATE_IMAGE_CLOSE_HOT = 6,				// State
		VSI_DATA_STATE_IMAGE_ERROR = 7,					// State
		VSI_DATA_STATE_IMAGE_WARNING = 8,				// State
		VSI_DATA_STATE_IMAGE_SINGLE = 9,				// State
		VSI_DATA_STATE_IMAGE_LEFT = 10,					// State
		VSI_DATA_STATE_IMAGE_RIGHT = 11,				// State
		VSI_DATA_STATE_IMAGE_NORMAL = 12,				// Dummy
		VSI_DATA_STATE_IMAGE_REVIEWED = 13,				// Dummy
		VSI_DATA_STATE_IMAGE_ACTIVE_NORMAL = 14,		// Dummy
		VSI_DATA_STATE_IMAGE_ACTIVE_REVIEWED = 15,		// Dummy
		VSI_DATA_STATE_IMAGE_INACTIVE_NORMAL = 16,		// Dummy
		VSI_DATA_STATE_IMAGE_INACTIVE_REVIEWED = 17,	// Dummy
		VSI_DATA_STATE_IMAGE_FILTER = 18,				// Dummy
		VSI_DATA_STATE_IMAGE_MAX = 19,					// Max
	};

	enum
	{
		VSI_DATA_IMAGE_UNCHECKED = 0,				// Image (0 base)
		VSI_DATA_IMAGE_CHECKED,
		VSI_DATA_IMAGE_UNCHECKED_LOCK_ALL,
		VSI_DATA_IMAGE_CHECKED_LOCK_ALL,
		VSI_DATA_IMAGE_STUDY_ACCESS_PRIVATE,
		VSI_DATA_IMAGE_STUDY_ACCESS_GROUP,
		VSI_DATA_IMAGE_STUDY_ACCESS_ALL,
		VSI_DATA_IMAGE_FRAME,
		VSI_DATA_IMAGE_CINE,
		VSI_DATA_IMAGE_3D,
	};

	COLORREF m_prgbState[VSI_DATA_STATE_IMAGE_MAX];

	class CVsiListItem : public std::list<CVsiItem*>
	{
	public:
		~CVsiListItem()
		{
			clear();
		}
		CVsiListItem& operator=(const CVsiListItem& right)
		{
			if (this == &right)
				return *this;

			std::list<CVsiItem*>::operator=(right);

			for (auto iter = begin(); iter != end();	++iter)
			{
				CVsiItem *pItem = *iter;
				pItem->AddRef();
			}

			return *this;
		}
		void push_back(CVsiItem *pItem)
		{
			for (auto iter = begin(); iter != end();	++iter)
			{
				if(pItem == *iter)
				{
					return;
				}
			}

			std::list<CVsiItem*>::push_back(pItem);
			pItem->AddRef();
		}
		CVsiListItem::iterator erase(CVsiListItem::iterator iter)
		{
			CVsiItem *pItem = *iter;

			iterator iterNext = std::list<CVsiItem*>::erase(iter);

			pItem->Release();

			return iterNext;
		}
		void clear()
		{
			for (auto iter = begin(); iter != end();	++iter)
			{
				CVsiItem *pItem = *iter;
				pItem->Release();
			}

			std::list<CVsiItem*>::clear();
		}
	};

	class CVsiItem
	{
	public:
		LONG m_ulRef;
		VSI_DATA_LIST_TYPE m_type;
		CVsiDataListWnd *m_pRoot;
		CVsiSort *m_pSort;
		BOOL m_bSelected : 1;
		BOOL m_bFocused : 1;
		BOOL m_bExpanded : 1;
		BOOL m_bVisible : 1;
		BYTE m_iStateImage : 4;
		BYTE m_iStateText : 5;
		BYTE m_iLevel : 3;

		CVsiItem *m_pParent;
		CVsiListItem m_listChild;

	public:
		CVsiItem() :
			m_ulRef(0),
			m_type(VSI_DATA_LIST_NONE),
			m_pRoot(NULL),
			m_pSort(NULL),
			m_bSelected(FALSE),
			m_bFocused(FALSE),
			m_bExpanded(FALSE),
			m_bVisible(TRUE),
			m_iStateImage(0),
			m_iStateText(0),
			m_iLevel(1),
			m_pParent(NULL)
		{
		}
		virtual ~CVsiItem()
		{
			Clear();
		}
		// Copy constructor
		CVsiItem::CVsiItem(const CVsiItem& right)
		{
			m_ulRef = 0;
			*this = right;
		}
		virtual CVsiItem& operator=(const CVsiItem& right)
		{
			if (this == &right)
				return *this;

			m_type = right.m_type;
			m_pRoot = right.m_pRoot;
			m_pSort = right.m_pSort;
			m_bSelected = right.m_bSelected;
			m_bFocused = right.m_bFocused;
			m_bExpanded = right.m_bExpanded;
			m_bVisible = right.m_bVisible;
			m_iStateImage = right.m_iStateImage;
			m_iStateText = right.m_iStateText;
			m_iLevel = right.m_iLevel;
			m_pParent = right.m_pParent;
			m_listChild = right.m_listChild;

			return *this;
		}
		virtual bool operator<(const CVsiItem&) const
		{
			return false;
		}

		ULONG AddRef()
		{
			return ++m_ulRef;
		}

		ULONG Release()
		{
			_ASSERT(m_ulRef > 0);

			if (0 == --m_ulRef)
			{
				delete this;
				return 0;
			}
			else
			{
				return m_ulRef;
			}
		}

		virtual void Reload()
		{
		}
		virtual void Clear()
		{
			m_listChild.clear();
		}
		virtual BOOL IsParentOf(const CVsiItem& item)
		{
			CVsiListItem::iterator iter;
			BOOL bFound = FALSE;
			CComHeapPtr<OLECHAR> pszId;
			CComHeapPtr<OLECHAR> pszId1;

			item.m_pParent->GetId(&pszId);
			GetId(&pszId1);
			if (lstrcmp(pszId, pszId1) == 0)
			{
				bFound = TRUE;
			}
			return bFound;
		}
		virtual HRESULT GetId(LPOLESTR *ppszId)
		{
			UNREFERENCED_PARAMETER(ppszId);
			return E_NOTIMPL;
		}
		virtual BOOL GetIsExported()
		{
			return FALSE;
		}
		virtual HRESULT Collapse(bool bKeepData = false)
		{
			m_bExpanded = FALSE;

			// Any child selected?
			for (auto iter = m_listChild.begin(); iter != m_listChild.end(); ++iter)
			{
				CVsiItem *pItem = *iter;

				pItem->Collapse(bKeepData);

				if (pItem->m_bSelected)
					m_bSelected = TRUE;
				if (pItem->m_bFocused)
					m_bFocused = TRUE;

				pItem->m_bSelected = FALSE;
				pItem->m_bFocused = FALSE;
			}

			if (!bKeepData && 0 == (m_pRoot->m_dwFlags & VSI_DATA_LIST_FLAG_KEEP_CHILDREN_ON_COLLAPSE))
			{
				// Removes children
				Clear();
			}

			return S_OK;
		}
		virtual HRESULT Expand(BOOL bSelectAllChildren = FALSE, BOOL *pbCancel = NULL)
		{
			UNREFERENCED_PARAMETER(bSelectAllChildren);
			UNREFERENCED_PARAMETER(pbCancel);
			return E_NOTIMPL;
		}
		virtual HRESULT Delete(BOOL bDelete)
		{
			UNREFERENCED_PARAMETER(bDelete);
			return E_NOTIMPL;
		}
		virtual void GetDisplayText(VSI_DATA_LIST_COL col, LPWSTR pszText, int iText)
		{
			UNREFERENCED_PARAMETER(col);
			UNREFERENCED_PARAMETER(iText);
			*pszText = 0;
		}

		virtual bool HasText(LPCWSTR pszText, bool bDeepSearch)
		{
			bool bHasText(false);
			UNREFERENCED_PARAMETER(bDeepSearch);

			Concurrency::structured_task_group tasks;
			tasks.run_and_wait([&]
			{
				Concurrency::parallel_for(0, static_cast<int>(VSI_DATA_LIST_COL_LAST), [&](int i)
				{
					if (VSI_DATA_LIST_COL_LOCKED != i && VSI_DATA_LIST_COL_SIZE != i)
					{
						WCHAR szTest[500];
						*szTest = 0;
						GetDisplayText(static_cast<VSI_DATA_LIST_COL>(i), szTest, _countof(szTest));
						if (0 != *szTest)
						{
							if (NULL != StrStrI(szTest, pszText))
							{
								bHasText = true;
								tasks.cancel();
							}
						}
					}
				});
			});

			return bHasText;
		}

		bool HasText(LPCWSTR pszCol, LPCWSTR pszText)
		{
			return (NULL != StrStrI(pszCol, pszText));
		}

		void Sort()
		{
			m_listChild.sort(
				[](CVsiDataListWnd::CVsiItem *px, CVsiDataListWnd::CVsiItem *py)->bool
				{
					return *px < *py;
				});
		}
		
		CVsiListItem::size_type GetVisibleChildrenCount(bool bRecurs = true)
		{
			CVsiListItem::size_type count = 0;

			for (auto iter = m_listChild.begin();
				iter != m_listChild.end();
				++iter)
			{
				CVsiItem *pItem = *iter;
				if (pItem->m_bVisible)
				{
					++count;
					if (bRecurs && pItem->m_bExpanded)
					{
						count += pItem->GetVisibleChildrenCount();
					}
				}
			}

			return count;
		}
	};

	class CVsiStudy : public CVsiItem
	{
	public:
		CComPtr<IVsiStudy> m_pStudy;

	public:
		CVsiStudy()
		{
			m_type = VSI_DATA_LIST_STUDY;
		}
		virtual ~CVsiStudy()
		{
		}
		// Copy constructor
		CVsiStudy::CVsiStudy(const CVsiStudy& right)
		{
			m_ulRef = 0;
			*this = right;
		}
		virtual CVsiStudy& operator=(const CVsiItem& right)
		{
			if (this == &right)
				return *this;

			// copy base member
			CVsiItem::operator= (right);

			m_pStudy = ((CVsiStudy*)(&right))->m_pStudy;

			return *this;
		}
		virtual bool operator<(const CVsiItem& r) const
		{
			const CVsiStudy &right = (const CVsiStudy &)r;

			VSI_PROP_STUDY prop;

			switch (m_pSort->m_sortStudyCol)
			{
			case VSI_DATA_LIST_COL_NAME:
				prop = VSI_PROP_STUDY_NAME;
				break;
			case VSI_DATA_LIST_COL_LOCKED:
				prop = VSI_PROP_STUDY_LOCKED;
				break;
			case VSI_DATA_LIST_COL_OWNER:
				prop = VSI_PROP_STUDY_OWNER;
				break;
			default:
			case VSI_DATA_LIST_COL_DATE:
				prop = VSI_PROP_STUDY_CREATED_DATE;
				break;
			}

			int iCmp = 0;

			m_pStudy->Compare(right.m_pStudy, prop, &iCmp);
			
			if (m_pSort->m_bStudyDescending)
				iCmp *= -1;

			return (iCmp < 0);
		}
		virtual void Clear()
		{
			HRESULT hr(S_OK);

			try
			{
				CVsiItem::Clear();
				hr = m_pStudy->RemoveAllSeries();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			VSI_CATCH(hr);
		}
		virtual HRESULT GetId(LPOLESTR *ppszId)
		{
			CComVariant vId;
			HRESULT hr = m_pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
			if (S_OK == hr)
				*ppszId = AtlAllocTaskWideString((LPCWSTR)V_BSTR(&vId));

			return (NULL != *ppszId) ? S_OK : E_FAIL;
		}
		virtual BOOL GetIsExported()
		{
			CComVariant vbExported;
			HRESULT hr = m_pStudy->GetProperty(VSI_PROP_STUDY_EXPORTED, &vbExported);
			if (S_OK == hr)
				return (VARIANT_FALSE != V_BOOL(&vbExported));

			return FALSE;
		}
		virtual HRESULT Expand(BOOL bSelectAllChildren = FALSE, BOOL *pbCancel = NULL)
		{
			HRESULT hr = S_OK;

			try
			{
				if (0 == m_listChild.size())
				{
					DWORD dwFlags = VSI_DATA_TYPE_SERIES_LIST;

					if (m_pRoot->m_dwFlags & VSI_DATA_LIST_FLAG_NO_SESSION_LINK)
					{
						dwFlags |= VSI_DATA_TYPE_NO_SESSION_LINK;
					}

					// Fill m_listChild
					CComQIPtr<IVsiPersistStudy> pPersist(m_pStudy);
					hr = pPersist->LoadStudyData(NULL, dwFlags, pbCancel);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiEnumSeries> pEnum;
					hr = m_pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_MIN, FALSE, &pEnum, pbCancel);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						CComPtr<IVsiSeries> pSeries;
						while (pEnum->Next(1, &pSeries, NULL) == S_OK)
						{
							AddItem(pSeries);
							pSeries.Release();

							if (pbCancel && *pbCancel)
							{
								Clear();
								break;
							}
						}

						if (0 < m_pRoot->m_vecFilterJobs.size())
						{
							BOOL bStop(FALSE);

							auto & job = m_pRoot->m_vecFilterJobs.back();

							m_pRoot->Filter(m_listChild, job, &bStop);
						}
					}
				}

				if (NULL == pbCancel || FALSE == *pbCancel)
				{
					m_bExpanded = TRUE;

					if (bSelectAllChildren)
					{
						for (auto iter = m_listChild.begin();
							iter != m_listChild.end();
							++iter)
						{
							CVsiItem *pItem = *iter;
							pItem->m_bSelected = bSelectAllChildren;
						}
					}

					// Expand series if there is only 1 series under this study
					if (1 == m_listChild.size())
					{
						auto iter = m_listChild.begin();
						CVsiItem *pItem = *iter;

						if (m_pRoot->m_iLevelMax > pItem->m_iLevel)
						{
							CVsiSeries *pSeries = dynamic_cast<CVsiSeries*>(pItem);
							VSI_VERIFY(nullptr != pSeries, VSI_LOG_ERROR, NULL);

							hr = pSeries->Expand(bSelectAllChildren, pbCancel);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
					}
				}
				else
				{
					Clear();
				}
			}
			VSI_CATCH_(err)
			{
				hr = err;
				if (FAILED(hr))
					VSI_ERROR_LOG(err);

				m_bExpanded = FALSE;
				Clear();
				m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
			}

			return hr;
		}
		virtual HRESULT Delete(BOOL bDelete)
		{
			HRESULT hr = S_OK;

			try
			{
				// If the study has not been exported and the user has not given
				// permission to delete un-exported studies then skip this one.
				if (!GetIsExported() && !bDelete)
					return hr;

				CComHeapPtr<OLECHAR> pszFile;
				hr = m_pStudy->GetDataPath(&pszFile);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				LPCWSTR pszSpt = wcsrchr(pszFile, L'\\');
				CString strPath(pszFile, (int)(pszSpt - pszFile));

				CString strDest(strPath);
				strDest.Append(L".tmp");

				BOOL bRet = MoveFile(strPath, strDest);
				if (bRet)
				{
					VsiRemoveDirectory(strDest);
				}
			}
			VSI_CATCH_(err)
			{
				hr = err;
				if (FAILED(hr))
					VSI_ERROR_LOG(err);

				m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
			}

			return hr;
		}
		virtual void GetDisplayText(VSI_DATA_LIST_COL col, LPWSTR pszText, int iText)
		{
			HRESULT hr;
			*pszText = 0;

			switch (col)
			{
			case VSI_DATA_LIST_COL_NAME:
				{
					CComVariant v;
					hr = m_pStudy->GetProperty(VSI_PROP_STUDY_NAME, &v);
					if (S_OK == hr && 0 != *V_BSTR(&v))
					{
						lstrcpyn(pszText, V_BSTR(&v), iText);
					}
				}
				break;
			case VSI_DATA_LIST_COL_LOCKED:
				{
					CComVariant v;
					hr = m_pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &v);
					if (S_OK == hr)
					{
						lstrcpyn(pszText, (VARIANT_FALSE != V_BOOL(&v)) ? L"Locked" : L"", iText);
					}
				}
				break;
			case VSI_DATA_LIST_COL_OWNER:
				{
					CComVariant v;
					hr = m_pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &v);
					if (S_OK == hr && 0 != *V_BSTR(&v))
					{
						lstrcpyn(pszText, V_BSTR(&v), iText);
					}
				}
				break;
			case VSI_DATA_LIST_COL_DATE:
				{
					CComVariant v;
					hr = m_pStudy->GetProperty(VSI_PROP_STUDY_CREATED_DATE, &v);
					if (S_OK == hr)
					{
						COleDateTime date(v);

						SYSTEMTIME stCreate;
						date.GetAsSystemTime(stCreate);

						// The time is stored as UTC. Convert it to local time for display.
						BOOL bRet = SystemTimeToTzSpecificLocalTime(NULL, &stCreate, &stCreate);
						if (bRet)
						{
							GetDateFormat(
								LOCALE_USER_DEFAULT,
								DATE_SHORTDATE,
								&stCreate,
								NULL,
								pszText,
								iText);
						}
					}
				}
				break;
			case VSI_DATA_LIST_COL_SIZE:
				{
					CComHeapPtr<OLECHAR> pszPath;
					hr = m_pStudy->GetDataPath(&pszPath);
					if (S_OK == hr)
					{
						NM_DL_ITEM nt = { 0 };
						nt.hdr.code = (UINT)NM_VSI_DL_GET_ITEM_STATUS;
						nt.hdr.hwndFrom = m_pRoot->m_hWnd;
						nt.hdr.idFrom = m_pRoot->m_wIdCtl;

						LPWSTR pszSpt = wcsrchr(pszPath, L'\\');
						CString strPath(pszPath, (int)(pszSpt - pszPath));
						INT64 iTotalSize = VsiGetDirectorySize(strPath);

						DWORD dwFlags = VSI_DATA_TYPE_SERIES_LIST;

						if (m_pRoot->m_dwFlags & VSI_DATA_LIST_FLAG_NO_SESSION_LINK)
						{
							dwFlags |= VSI_DATA_TYPE_NO_SESSION_LINK;
						}

						CComQIPtr<IVsiPersistStudy> pPersist(m_pStudy);
						hr = pPersist->LoadStudyData(NULL, dwFlags, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IVsiEnumSeries> pEnum;
						hr = m_pStudy->GetSeriesEnumerator(VSI_PROP_SERIES_MIN, FALSE, &pEnum, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (S_OK == hr)
						{
							CComPtr<IVsiSeries> pSeries;
							while (pEnum->Next(1, &pSeries, NULL) == S_OK)
							{
								pszPath.Free();
								pSeries->GetDataPath(&pszPath);

								nt.type = VSI_DATA_LIST_SERIES;
								nt.pUnkItem = pSeries;
								nt.pszItemPath = pszPath;
								LRESULT lRes = m_pRoot->GetParent().SendMessage(WM_NOTIFY, nt.hdr.idFrom, (LPARAM)&nt);
								// Checks if it's not removed
								if (VSI_DATA_LIST_ITEM_STATUS_HIDDEN == (int)lRes)
								{
									pszSpt = wcsrchr(pszPath, L'\\');
									CString strPathSeries(pszPath, (int)(pszSpt - pszPath));
									UINT64 iSeriesSize = VsiGetDirectorySize(strPathSeries);
									iTotalSize -= iSeriesSize;
								}

								pSeries.Release();
							}
						}

						WCHAR szSpace[64];
						double dbSpace = (double)(iTotalSize /  VSI_ONE_MEGABYTE);
						dbSpace = (dbSpace < 0.1) ? 0.1 : dbSpace;
						swprintf_s(szSpace, _countof(szSpace), L"%.1f", dbSpace);
						VsiGetDoubleFormat(szSpace, szSpace, _countof(szSpace), 1);
						swprintf_s(pszText, iText, L"%s MB", szSpace);
					}
				}
				break;
			}
		}
		virtual bool HasText(LPCWSTR pszText, bool bDeepSearch)
		{
			bool bHasText = CVsiItem::HasText(pszText, bDeepSearch);
			if (!bHasText)
			{
				VSI_PROP_STUDY pProps[] = {
					VSI_PROP_STUDY_GRANTING_INSTITUTION,
					VSI_PROP_STUDY_NOTES
				};

				for (int i = 0; i < _countof(pProps); ++i)
				{
					CComVariant vText;
					HRESULT hr = m_pStudy->GetProperty(pProps[i], &vText);
					if (S_OK == hr)
					{
						bHasText = (NULL != StrStrI(V_BSTR(&vText), pszText));
						if (bHasText)
						{
							break;
						}
					}
				}
			}

			return bHasText;
		}

		void AddItem(IVsiSeries *pSeries) throw(...)
		{
			CVsiSeries *p = new CVsiSeries;
			VSI_CHECK_MEM(p, VSI_LOG_ERROR, NULL);

			VSI_SERIES_ERROR dwErrorCode(VSI_SERIES_ERROR_NONE);
			pSeries->GetErrorCode(&dwErrorCode);
			if (VSI_SERIES_ERROR_VERSION_INCOMPATIBLE & dwErrorCode)
			{
				p->m_iStateImage = VSI_DATA_STATE_IMAGE_WARNING;
			}
			else if (VSI_SERIES_ERROR_NONE != dwErrorCode)
			{
				p->m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
			}
			p->m_pRoot = m_pRoot;
			p->m_pSeries = pSeries;
			p->m_pSort = m_pSort;
			p->m_pParent = this;
			p->m_iLevel = m_iLevel + 1;
			p->m_bSelected = m_bSelected;

			if (0 < m_pRoot->m_vecFilterJobs.size())
			{
				auto const & job = m_pRoot->m_vecFilterJobs.back();
				m_pRoot->Filter(p, job.m_vecFilters, job.m_filterType);
			}

			m_listChild.push_back(p);
		}
	};

	class CVsiSeries : public CVsiItem
	{
	public:
		CComPtr<IVsiSeries> m_pSeries;

	public:
		CVsiSeries()
		{
			m_type = VSI_DATA_LIST_SERIES;
		}
		virtual ~CVsiSeries()
		{
		}
		// Copy constructor
		CVsiSeries::CVsiSeries(const CVsiSeries& right)
		{
			m_ulRef = 0;
			*this = right;
		}
		virtual CVsiSeries& operator=(const CVsiItem& right)
		{
			if (this == &right)
				return *this;

			// copy base member
			CVsiItem::operator= (right);

			m_pSeries = ((CVsiSeries*)(&right))->m_pSeries;

			return *this;
		}
		virtual bool operator<(const CVsiItem& r) const
		{
			const CVsiSeries &right = (const CVsiSeries &)r;

			VSI_PROP_SERIES prop;

			switch (m_pSort->m_sortSeriesCol)
			{
			case VSI_DATA_LIST_COL_NAME:
				prop = VSI_PROP_SERIES_NAME;
				break;
			case VSI_DATA_LIST_COL_ACQ_BY:
				prop = VSI_PROP_SERIES_ACQUIRED_BY;
				break;
			case VSI_DATA_LIST_COL_ANIMAL_ID:
				prop = VSI_PROP_SERIES_ANIMAL_ID;
				break;
			case VSI_DATA_LIST_COL_PROTOCOL_ID:
				prop = VSI_PROP_SERIES_CUSTOM4;
				break;
			default:
			case VSI_DATA_LIST_COL_DATE:
				prop = VSI_PROP_SERIES_CREATED_DATE;
				break;
			}

			int iCmp = 0;

			m_pSeries->Compare(right.m_pSeries, prop, &iCmp);

			if (m_pSort->m_bSeriesDescending)
			{
				iCmp *= -1;
			}

			return (iCmp < 0);
		}
		virtual void Clear()
		{
			HRESULT hr(S_OK);

			try
			{
				CVsiItem::Clear();
				hr = m_pSeries->RemoveAllImages();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			VSI_CATCH(hr);
		}
		virtual HRESULT GetId(LPOLESTR *ppszId)
		{
			CComVariant vId;
			HRESULT hr = m_pSeries->GetProperty(VSI_PROP_SERIES_ID, &vId);
			if (S_OK == hr)
				*ppszId = AtlAllocTaskWideString((LPCWSTR)V_BSTR(&vId));

			return (NULL != *ppszId) ? S_OK : E_FAIL;
		}
		virtual BOOL GetIsExported()
		{
			CComVariant vbExported;
			HRESULT hr = m_pSeries->GetProperty(VSI_PROP_SERIES_EXPORTED, &vbExported);
			if (S_OK == hr)
				return (VARIANT_FALSE != V_BOOL(&vbExported));

			return FALSE;
		}
		virtual HRESULT Expand(BOOL bSelectAllChildren = FALSE, BOOL *pbCancel = NULL)
		{
			HRESULT hr = S_OK;

			try
			{
				if (0 == m_listChild.size())
				{
					DWORD dwFlags = VSI_DATA_TYPE_IMAGE_LIST;

					if (m_pRoot->m_dwFlags & VSI_DATA_LIST_FLAG_NO_SESSION_LINK)
					{
						dwFlags |= VSI_DATA_TYPE_NO_SESSION_LINK;
					}

					// Fill m_listChild
					CComQIPtr<IVsiPersistSeries> pPersist(m_pSeries);
					hr = pPersist->LoadSeriesData(NULL, dwFlags, pbCancel);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiEnumImage> pEnum;
					hr = m_pSeries->GetImageEnumerator(
						VSI_PROP_IMAGE_MIN,
						TRUE,
						&pEnum, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						CComPtr<IVsiImage> pImage;
						while (pEnum->Next(1, &pImage, NULL) == S_OK)
						{
							AddItem(pImage);
							pImage.Release();

							if (pbCancel && *pbCancel)
							{
								Clear();
								break;
							}
						}

						if (0 < m_pRoot->m_vecFilterJobs.size())
						{
							BOOL bStop(FALSE);

							auto & job = m_pRoot->m_vecFilterJobs.back();

							m_pRoot->Filter(m_listChild, job, &bStop);
						}
					}
				}
				
				if (NULL == pbCancel || FALSE == *pbCancel)
				{
					if (bSelectAllChildren)
					{
						for (auto iter = m_listChild.begin();
							iter != m_listChild.end();
							++iter)
						{
							CVsiItem *pItem = *iter;
							pItem->m_bSelected = bSelectAllChildren;
						}
					}

					m_bExpanded = TRUE;
				}
				else
				{
					Clear();
				}
			}
			VSI_CATCH_(err)
			{
				hr = err;
				if (FAILED(hr))
					VSI_ERROR_LOG(err);

				m_bExpanded = FALSE;
				Clear();
				m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
			}

			return hr;
		}
		virtual HRESULT Delete(BOOL bDelete)
		{
			HRESULT hr = S_OK;

			try
			{
				// If the study has not been exported and the user has not given
				// permission to delete un-exported studies then skip this one.
				if (!GetIsExported() && !bDelete)
					return hr;

				CComHeapPtr<OLECHAR> pszFile;
				hr = m_pSeries->GetDataPath(&pszFile);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				LPCWSTR pszSpt = wcsrchr(pszFile, L'\\');
				CString strPath(pszFile, (int)(pszSpt - pszFile));

				CString strDest(strPath);
				strDest.Append(L".tmp");

				BOOL bRet = MoveFile(strPath, strDest);
				if (bRet)
				{
					VsiRemoveDirectory(strDest);
				}
				else
				{
					// TODO: Error
				}
			}
			VSI_CATCH_(err)
			{
				hr = err;
				if (FAILED(hr))
					VSI_ERROR_LOG(err);

				m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
			}

			return hr;
		}
		virtual void GetDisplayText(VSI_DATA_LIST_COL col, LPWSTR pszText, int iText)
		{
			*pszText = 0;

			HRESULT hr;

			switch (col)
			{
			case VSI_DATA_LIST_COL_NAME:
				{
					CComVariant v;
					hr = m_pSeries->GetProperty(VSI_PROP_SERIES_NAME, &v);
					if (S_OK == hr && 0 != *V_BSTR(&v))
					{
						lstrcpyn(pszText, V_BSTR(&v), iText);
					}
				}
				break;
			case VSI_DATA_LIST_COL_DATE:
				{
					SYSTEMTIME stDate;
					CComVariant vDate;
					hr = m_pSeries->GetProperty(VSI_PROP_SERIES_CREATED_DATE, &vDate);
					if (S_OK == hr)
					{
						COleDateTime date(vDate);
						date.GetAsSystemTime(stDate);

						// The time is stored as UTC. Convert it to local time for display.
						BOOL bRet = SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);
						if (bRet)
						{
							GetDateFormat(
								LOCALE_USER_DEFAULT,
								DATE_SHORTDATE,
								&stDate,
								NULL,
								pszText,
								iText);
						}
					}
				}
				break;
			case VSI_DATA_LIST_COL_ACQ_BY:
				{
					CComVariant v;
					hr = m_pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &v);
					if (S_OK == hr && 0 != *V_BSTR(&v))
					{
						lstrcpyn(pszText, V_BSTR(&v), iText);
					}
				}
				break;
			case VSI_DATA_LIST_COL_SIZE:
				{
					CComHeapPtr<OLECHAR> pszPath;
					hr = m_pSeries->GetDataPath(&pszPath);
					if (SUCCEEDED(hr))
					{
						LPWSTR pszSpt = wcsrchr(pszPath, L'\\');
						CString strPath(pszPath, (int)(pszSpt - pszPath));
						INT64 iTotalSize = VsiGetDirectorySize(strPath);
						TCHAR str[64];
						double dbSpace = (double)(iTotalSize /  VSI_ONE_MEGABYTE);
						dbSpace = (dbSpace < 0.1) ? 0.1 : dbSpace;
						_stprintf_s(str, _countof(str), L"%.1f MB", dbSpace);
						lstrcpyn(pszText, str, iText);
					}
				}
				break;
			case VSI_DATA_LIST_COL_ANIMAL_ID:
				{
					CComVariant v;
					hr = m_pSeries->GetProperty(VSI_PROP_SERIES_ANIMAL_ID, &v);
					if (S_OK == hr && 0 != *V_BSTR(&v))
					{
						lstrcpyn(pszText, V_BSTR(&v), iText);
					}
				}
				break;
			case VSI_DATA_LIST_COL_PROTOCOL_ID:
				{
					CComVariant v;
					hr = m_pSeries->GetProperty(VSI_PROP_SERIES_CUSTOM4, &v);
					if (S_OK == hr && 0 != *V_BSTR(&v))
					{
						lstrcpyn(pszText, V_BSTR(&v), iText);
					}
				}
				break;
			}
		}
		virtual bool HasText(LPCWSTR pszText, bool bDeepSearch)
		{
			bool bHasText = CVsiItem::HasText(pszText, bDeepSearch);
			if (!bHasText)
			{
				VSI_PROP_SERIES pProps[] = {
					VSI_PROP_SERIES_ANIMAL_ID,
					VSI_PROP_SERIES_APPLICATION,
					VSI_PROP_SERIES_MSMNT_PACKAGE,
					VSI_PROP_SERIES_ANIMAL_COLOR,
					VSI_PROP_SERIES_STRAIN,
					VSI_PROP_SERIES_SOURCE,
					VSI_PROP_SERIES_WEIGHT,
					VSI_PROP_SERIES_TYPE,
					VSI_PROP_SERIES_HEART_RATE,
					VSI_PROP_SERIES_BODY_TEMP,
					VSI_PROP_SERIES_NOTES,
					VSI_PROP_SERIES_CUSTOM1,
					VSI_PROP_SERIES_CUSTOM2,
					VSI_PROP_SERIES_CUSTOM3,
					VSI_PROP_SERIES_CUSTOM4,
					VSI_PROP_SERIES_CUSTOM5,
					VSI_PROP_SERIES_CUSTOM6,
					VSI_PROP_SERIES_CUSTOM7,
					VSI_PROP_SERIES_CUSTOM8,
				};

				for (int i = 0; i < _countof(pProps); ++i)
				{
					CComVariant vText;
					HRESULT hr = m_pSeries->GetProperty(pProps[i], &vText);
					if (S_OK == hr)
					{
						bHasText = (NULL != StrStrI(V_BSTR(&vText), pszText));
						if (bHasText)
						{
							break;
						}
					}
				}

				if (!bHasText)
				{
					VSI_PROP_SERIES pProps[] = {
						VSI_PROP_SERIES_DATE_OF_BIRTH,
						VSI_PROP_SERIES_DATE_MATED,
						VSI_PROP_SERIES_DATE_PLUGGED
					};

					for (int i = 0; i < _countof(pProps); ++i)
					{
						CComVariant vDate;
						HRESULT hr = m_pSeries->GetProperty(pProps[i], &vDate);
						if (S_OK == hr)
						{
							WCHAR szDate[100];
							SYSTEMTIME st;

							COleDateTime date(vDate);
							date.GetAsSystemTime(st);
							GetDateFormat(
								LOCALE_USER_DEFAULT,
								DATE_SHORTDATE,
								&st,
								NULL,
								szDate,
								_countof(szDate));
							bHasText = (NULL != StrStrI(szDate, pszText));
							if (bHasText)
							{
								break;
							}
						}
					}
				}
			}

			return bHasText;
		}

		void AddItem(IVsiImage *pImage) throw(...)
		{
			CVsiImage *p = new CVsiImage;
			VSI_CHECK_MEM(p, VSI_LOG_ERROR, NULL);

			VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_NONE);
			pImage->GetErrorCode(&dwErrorCode);
			if (VSI_IMAGE_ERROR_VERSION_INCOMPATIBLE & dwErrorCode)
			{
				p->m_iStateImage = VSI_DATA_STATE_IMAGE_WARNING;
			}
			else if (VSI_IMAGE_ERROR_NONE != dwErrorCode)
			{
				p->m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
			}
			p->m_pRoot = m_pRoot;
			p->m_pImage = pImage;
			p->m_pSort = m_pSort;
			p->m_pParent = this;
			p->m_iLevel = m_iLevel + 1;
			p->m_bSelected = m_bSelected;

			if (0 < m_pRoot->m_vecFilterJobs.size())
			{
				auto const & job = m_pRoot->m_vecFilterJobs.back();
				m_pRoot->Filter(p, job.m_vecFilters, job.m_filterType);
			}

			m_listChild.push_back(p);
		}
	};

	class CVsiImage : public CVsiItem
	{
	public:
		CComPtr<IVsiImage> m_pImage;

	public:
		CVsiImage()
		{
			m_type = VSI_DATA_LIST_IMAGE;
		}
		// Copy constructor
		CVsiImage::CVsiImage(const CVsiImage& right)
		{
			*this = right;
		}
		CVsiImage& operator=(const CVsiItem& right)
		{
			if (this == &right)
				return *this;

			// copy base member
			CVsiItem::operator= (right);

			m_pImage = ((CVsiImage*)(&right))->m_pImage;

			return *this;
		}
		// Images sorting support
		bool operator<(const CVsiItem& r) const
		{
			const CVsiImage &right = (const CVsiImage &)r;

			VSI_PROP_IMAGE prop;

			switch (m_pSort->m_sortImageCol)
			{
			case VSI_DATA_LIST_COL_NAME:
				prop = VSI_PROP_IMAGE_NAME;
				break;
			default:
			case VSI_DATA_LIST_COL_DATE:
				prop = VSI_PROP_IMAGE_ACQUIRED_DATE;
				break;
			case VSI_DATA_LIST_COL_MODE:
				prop = VSI_PROP_IMAGE_MODE_DISPLAY;
				break;
			case VSI_DATA_LIST_COL_LENGTH:
				prop = VSI_PROP_IMAGE_LENGTH;
				break;
			}

			int iCmp = 0;
			m_pImage->Compare(right.m_pImage, prop, &iCmp);

			if (m_pSort->m_bImageDescending)
				iCmp *= -1;

			return (iCmp < 0);
		}
		virtual void Reload()
		{
			CComQIPtr<IVsiPersistImage> pPersist(m_pImage);
			if (NULL != pPersist)
			{
				pPersist->LoadImageData(NULL, NULL, NULL, 0);
			}
		}
		virtual HRESULT GetId(LPOLESTR *ppszId)
		{
			CComVariant vId;
			HRESULT hr = m_pImage->GetProperty(VSI_PROP_IMAGE_ID, &vId);
			if (S_OK == hr)
				*ppszId = AtlAllocTaskWideString((LPCWSTR)V_BSTR(&vId));

			return (NULL != *ppszId) ? S_OK : E_FAIL;
		}
		virtual HRESULT Expand(BOOL bSelectAllChildren = FALSE, BOOL *pbCancel = NULL)
		{
			UNREFERENCED_PARAMETER(bSelectAllChildren);
			UNREFERENCED_PARAMETER(pbCancel);
			return S_FALSE;
		}
		virtual void GetDisplayText(VSI_DATA_LIST_COL col, LPWSTR pszText, int iText)
		{
			*pszText = 0;
			HRESULT hr;

			switch (col)
			{
			case VSI_DATA_LIST_COL_NAME:
				{
					CComVariant v;
					hr = m_pImage->GetProperty(VSI_PROP_IMAGE_NAME, &v);
					if (S_OK == hr && 0 != *V_BSTR(&v))
					{
						lstrcpyn(pszText, V_BSTR(&v), iText);
					}
				}
				break;
			case VSI_DATA_LIST_COL_DATE:
				{
					SYSTEMTIME stDate;
					CComVariant vDate;
					hr = m_pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vDate);
					if (S_OK == hr)
					{
						COleDateTime date(vDate);
						date.GetAsSystemTime(stDate);

						// The time is stored as UTC. Convert it to local time for display.
						BOOL bRet = SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);
						if (bRet)
						{
							GetTimeFormat(
								LOCALE_USER_DEFAULT,
								0,
								&stDate,
								NULL,
								pszText,
								iText);
						}
					}
				}
				break;
			case VSI_DATA_LIST_COL_ACQ_BY:
				{
					CVsiSeries *pSeries = dynamic_cast<CVsiSeries*>(m_pParent);
					if (NULL != pSeries)
					{
						CComVariant v;
						hr = pSeries->m_pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &v);
						if (S_OK == hr && 0 != *V_BSTR(&v))
						{
							lstrcpyn(pszText, V_BSTR(&v), iText);
						}
					}
				}
				break;
			case VSI_DATA_LIST_COL_MODE:
				{
					CComVariant v;
					hr = m_pImage->GetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &v);
					if (hr == S_OK)
					{
						BOOL bIsStrain(FALSE);

						CComVariant vVevoStrain;
						hr = m_pImage->GetProperty(VSI_PROP_IMAGE_VEVOSTRAIN, &vVevoStrain);
						if (hr == S_OK)
						{
							if (V_BOOL(&vVevoStrain) == VARIANT_TRUE)
							{
								CString strMode(V_BSTR(&v));
								strMode += L" (VevoStrain)";
								v = CComVariant(strMode);
								bIsStrain = TRUE;
							}
						}
						
						if (!bIsStrain)
						{
							CComVariant vSonoVevo;
							hr = m_pImage->GetProperty(VSI_PROP_IMAGE_VEVOCQ, &vSonoVevo);
							if (hr == S_OK)
							{
								if (V_BOOL(&vSonoVevo) == VARIANT_TRUE)
								{
									CString strMode(V_BSTR(&v));
									strMode += L" (VevoCQ)";
									v = CComVariant(strMode);
								}
							}
						}

						//I don't think we need this code? BUG 12364
						//CComVariant vAdvContMip;
						//hr = m_pImage->GetProperty(VSI_PROP_IMAGE_ADVCONTMIP, &vAdvContMip);
						//if(hr == S_OK)
						//{
						//	if(V_BOOL(&vAdvContMip) == VARIANT_TRUE)
						//	{
						//		CString strMode(V_BSTR(&v));
						//		int iMipStr = strMode.Find(L"(MIP)", 0);
						//		if(iMipStr == -1)
						//		{
						//			// Only copy measurement label once
						//			strMode += L" (MIP)";
						//			v = CComVariant(strMode);
						//		}
						//	}
						//}
					
						lstrcpyn(pszText, V_BSTR(&v), iText);
					}
				}
				break;
			case VSI_DATA_LIST_COL_LENGTH:
				{
					CComVariant vLength;
					hr = m_pImage->GetProperty(VSI_PROP_IMAGE_LENGTH, &vLength);
					if (S_OK == hr)
					{
						CComVariant vMode;
						hr = m_pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
						if (S_OK == hr)
						{
							CComVariant vType;
							hr = m_pRoot->m_pModeManager->GetModePropertyFromInternalName(
								V_BSTR(&vMode), L"lengthType", &vType);
							if (S_OK == hr)
							{
								_ASSERT(VT_I4 == V_VT(&vType));

								switch (V_I4(&vType))
								{
								case VSI_MODE_LENGTH_FRAME:
									_snwprintf_s(pszText, iText, _TRUNCATE, (V_R8(&vLength) > 1.0) ? L"%d Frames" : L"%d Frame", (int)V_R8(&vLength));
									break;
								case VSI_MODE_LENGTH_TIME:
									{
										_snwprintf_s(pszText, iText, _TRUNCATE, L"%.2f", V_R8(&vLength));
										VsiGetDoubleFormat(pszText, pszText, iText, 2);
										wcscat_s(pszText, iText, L" Seconds");
									}
									break;

								case VSI_MODE_LENGTH_DISTANCE:
									{
										_snwprintf_s(pszText, iText, _TRUNCATE, L"%.2f", V_R8(&vLength));
										VsiGetDoubleFormat(pszText, pszText, iText, 2);
										wcscat_s(pszText, iText, L" mm");
									}
									break;

								default:
									_ASSERT(0);
								}
							}
							else
							{
								m_iStateImage = VSI_DATA_STATE_IMAGE_ERROR;
							}
						}
					}
				}
				break;
			case VSI_DATA_LIST_COL_ANIMAL_ID:
				{
					CVsiSeries *pSeries = dynamic_cast<CVsiSeries*>(m_pParent);
					if (NULL != pSeries)
					{
						CComVariant v;
						hr = pSeries->m_pSeries->GetProperty(VSI_PROP_SERIES_ANIMAL_ID, &v);
						if (S_OK == hr && 0 != *V_BSTR(&v))
						{
							lstrcpyn(pszText, V_BSTR(&v), iText);
						}
					}
				}
				break;
			case VSI_DATA_LIST_COL_PROTOCOL_ID:
				{
					CVsiSeries *pSeries = dynamic_cast<CVsiSeries*>(m_pParent);
					if (NULL != pSeries)
					{
						CComVariant v;
						hr = pSeries->m_pSeries->GetProperty(VSI_PROP_SERIES_CUSTOM4, &v);
						if (S_OK == hr && 0 != *V_BSTR(&v))
						{
							lstrcpyn(pszText, V_BSTR(&v), iText);
						}
					}
				}
				break;
			}
		}

	private:
		int CompareDates(const CVsiImage& right) const
		{	
			int iCmp = 0;
			HRESULT hr1, hr2;

			CComVariant v1, v2;
			hr1 = m_pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &v1);
			hr2 = right.m_pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &v2);
			if (S_OK == hr1 && S_OK == hr2)
			{
				COleDateTime cDate1(v1), cDate2(v2);
				if (cDate1 < cDate2)
				{
					iCmp = -1;
				}
				else if (cDate2 < cDate1)
				{
					iCmp = 1;
				}
				else  // Same - checks creation date
				{
					CComVariant vCreate1, vCreate2;
					m_pImage->GetProperty(VSI_PROP_IMAGE_CREATED_DATE, &vCreate1);
					right.m_pImage->GetProperty(VSI_PROP_IMAGE_CREATED_DATE, &vCreate2);
					cDate1 = vCreate1;
					cDate2 = vCreate2;
					if (cDate1 < cDate2)
					{
						iCmp = -1;
					}
					else if (cDate2 < cDate1)
					{
						iCmp = 1;
					}
				}
			}
			else if (S_OK == hr1)
			{
				iCmp = 1;
			}
			else if (S_OK == hr2)
			{
				iCmp = -1;
			}

			return iCmp;
		}
	};

private:
	CComQIPtr<IVsiApp> m_pApp;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiModeManager> m_pModeManager;
	CComPtr<IVsiStudyManager> m_pStudyMgr;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;

	CCriticalSection m_cs;

	OSVERSIONINFO m_osvi;
	
	DWORD m_dwFlags;

	CString m_strEmptyMsg;

	CImageList m_ilHeader;
	CImageList m_ilState;
	CImageList m_ilSmall;
	
	HFONT m_hfontNormal;
	HFONT m_hfontReviewed;
	
	WORD m_wIdCtl;

	BOOL m_bTrackLeave;
	int m_iLastHotItem;
	int m_iLastHotState;

	class CVsiSort
	{
	public:
		VSI_DATA_LIST_COL m_sortStudyCol;
		BOOL m_bStudyDescending;
		VSI_DATA_LIST_COL m_sortSeriesCol;
		BOOL m_bSeriesDescending;
		VSI_DATA_LIST_COL m_sortImageCol;
		BOOL m_bImageDescending;
		VSI_DATA_LIST_COL m_sortCol;
		BOOL m_bDescending;

		CVsiSort() :
			m_sortStudyCol(VSI_DATA_LIST_COL_DATE),
			m_bStudyDescending(TRUE),
			m_sortSeriesCol(VSI_DATA_LIST_COL_DATE),
			m_bSeriesDescending(TRUE),
			m_sortImageCol(VSI_DATA_LIST_COL_DATE),
			m_bImageDescending(TRUE),
			m_sortCol(VSI_DATA_LIST_COL_DATE),
			m_bDescending(TRUE)
		{
		}
	} m_sort;

	// Filers
	typedef enum {
		VSI_FILTER_TYPE_NONE = 0,
		VSI_FILTER_TYPE_OR,
		VSI_FILTER_TYPE_AND,
	} VSI_FILTER_TYPE;
	
	typedef enum
	{
		VSI_SEARCH_NONE = 0,
		VSI_SEARCH_ALL,
		VSI_SEARCH_HIDDEN,
		VSI_SEARCH_VISIBLE,
	} VSI_SEARCH_SCOPE;

	class CVsiFilterJob
	{
	public:
		std::vector<CString> m_vecFilters;
		VSI_FILTER_TYPE m_filterType;
		VSI_SEARCH_SCOPE m_scope;
		CVsiItem *m_pItemFilterStop;

		CVsiFilterJob() :
			m_filterType(VSI_FILTER_TYPE_NONE),
			m_scope(VSI_SEARCH_NONE),
			m_pItemFilterStop(nullptr)
		{
		}
		CVsiFilterJob::CVsiFilterJob(const CVsiFilterJob& right)
		{
			*this = right;
		}
		virtual CVsiFilterJob& operator=(const CVsiFilterJob& right)
		{
			if (this == &right)
				return *this;

			m_vecFilters = right.m_vecFilters;
			m_filterType = right.m_filterType;
			m_scope = right.m_scope;
			m_pItemFilterStop = right.m_pItemFilterStop;

			return *this;
		}
	};

	std::vector<CVsiFilterJob> m_vecFilterJobs;

	class CVsiFilterProgress
	{
	public:
		ULONG m_ulTotal;
		volatile ULONG m_ulDone;

		CVsiFilterProgress() :
			m_ulTotal(0),
			m_ulDone(0)
		{
		}

		void SetTotal(ULONG ulTotal)
		{
			m_ulTotal = ulTotal;
			m_ulDone = 0;
		}
		void IncrementDone()
		{
			InterlockedIncrement(&m_ulDone);
		}

		void GetText(CString &strOut)
		{
			if (m_ulTotal > 0)
			{
				DWORD i = m_ulDone * 100 / m_ulTotal;
				if (0 == i)
				{
					// Don't show 0%
					i = 1;
				}

				strOut.Format(L"%d%%", i);
			}
		}
	} m_FilterProgress;


	// Master data record
	CVsiListItem m_listRoot;

	std::map<int, VSI_DATA_LIST_COL> m_mapIndexToCol;  // Maps subitem to column id

	std::map<VSI_DATA_LIST_COL, int> m_mapColToIndex;  // Maps column id to subitem

	int m_iLevelMax;
	BOOL m_bNotifyStateChanged;

	std::vector<CVsiItem*> m_vecItem;  // Sorted items

	VSI_DATA_LIST_TYPE m_typeRoot; // VSI_DATA_LIST_STUDY | VSI_DATA_LIST_SERIES

	// Item location map
	class CVsiItemLocation
	{
	public:
		CVsiListItem *m_pListItem;
		CVsiListItem::iterator m_iter;

	public:
		CVsiItemLocation() :
			m_pListItem(NULL)
		{
		}
		CVsiItemLocation(CVsiListItem *pList, CVsiListItem::iterator iter) :
			m_pListItem(pList),
			m_iter(iter)
		{
		}
		CVsiItemLocation::CVsiItemLocation(const CVsiItemLocation& right)
		{
			*this = right;
		}
		virtual CVsiItemLocation& operator=(const CVsiItemLocation& right)
		{
			if (this == &right)
				return *this;

			m_pListItem = right.m_pListItem;
			m_iter = right.m_iter;

			return *this;
		}
	};

	typedef std::map<CString, CVsiItemLocation> CVsiMapItemLocation;

public:
	CVsiDataListWnd();
	~CVsiDataListWnd();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSIDATALISTWND)

	BEGIN_COM_MAP(CVsiDataListWnd)
		COM_INTERFACE_ENTRY(IVsiDataListWnd)
		COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct()
	{
		return CoCreateFreeThreadedMarshaler(
			GetControllingUnknown(), &m_pUnkMarshaler.p);
	}

	void FinalRelease()
	{
		m_pUnkMarshaler.Release();
	}

	CComPtr<IUnknown> m_pUnkMarshaler;

	VSI_DECLARE_USE_MESSAGE_LOG()

public:
	// IVsiDataListWnd
	STDMETHOD(Initialize)(
		HWND hwnd,
		DWORD dwFlags,
		VSI_DATA_LIST_TYPE typeRoot,
		IVsiApp *pApp,
		const VSI_LVCOLUMN *pColumns,
		int iColumns);
	STDMETHOD(Uninitialize)();
	STDMETHOD(LockData)();
	STDMETHOD(UnlockData)();
	STDMETHOD(SetColumns)(const VSI_LVCOLUMN *pColumns, int iColumns);
	STDMETHOD(SetTextColor)(VSI_DATA_LIST_ITEM_STATUS dwState, COLORREF rgbState);
	STDMETHOD(SetFont)(HFONT hfontNormal, HFONT hfontReviewed);
	STDMETHOD(SetEmptyMessage)(LPCOLESTR pszMessage);
	STDMETHOD(GetSortSettings)(VSI_DATA_LIST_TYPE type, VSI_DATA_LIST_COL *pCol, BOOL *pbDescending);
	STDMETHOD(SetSortSettings)(VSI_DATA_LIST_TYPE type, VSI_DATA_LIST_COL col, BOOL bDescending);
	STDMETHOD(Clear)();
	STDMETHOD(Fill)(IUnknown *pUnkEnum, int iLevelMax);
	STDMETHOD(Filter)(LPCWSTR pszSearch, BOOL *pbStop);
	STDMETHOD(GetItemCount)(int *piCount);
	STDMETHOD(GetSelectedItem)(int *piIndex);
	STDMETHOD(SetSelectedItem)(int iIndex, BOOL bEnsureVisible);
	STDMETHOD(SelectAll)();
	STDMETHOD(GetSelection)(VSI_DATA_LIST_COLLECTION *pSelection);
	STDMETHOD(SetSelection)(LPCOLESTR pszNamespace, BOOL bEnsureVisible);
	STDMETHOD(GetLastSelectedItem)(VSI_DATA_LIST_TYPE *pType, IUnknown **ppUnk);
	STDMETHOD(UpdateItem)(LPCOLESTR pszNamespace, BOOL bRecurChildren);
	STDMETHOD(UpdateItemStatus)(LPCOLESTR pszNamespace, BOOL bRecurChildren);
	STDMETHOD(UpdateList)(IXMLDOMDocument *pXmlDoc);
	STDMETHOD(GetItem)(int iIndex, VSI_DATA_LIST_TYPE *pType, IUnknown **ppUnk);
	STDMETHOD(GetItemStatus)(LPCOLESTR pszNamespace, VSI_DATA_LIST_ITEM_STATUS *pdwStatus);
	STDMETHOD(AddItem)(IUnknown *pUnkItem, VSI_DATA_LIST_SELECTION dwSelect);
	STDMETHOD(RemoveItem)(LPCOLESTR pszNamespace, VSI_DATA_LIST_SELECTION dwSelect);
	STDMETHOD(GetItems)(VSI_DATA_LIST_COLLECTION *pItems);

protected:
	BEGIN_MSG_MAP(CVsiDataListWnd)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
		MESSAGE_HANDLER(WM_LBUTTONDOWN, OnButtonDown);
		MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnButtonDown);
		MESSAGE_HANDLER(WM_RBUTTONDOWN, OnButtonDown);
		MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnButtonDown);
		MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove);
		MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)

		MESSAGE_HANDLER(LVM_SETITEMSTATE, OnSetItemState)
		MESSAGE_HANDLER(LVM_GETITEMSTATE, OnGetItemState)

		NOTIFY_CODE_HANDLER(HDN_BEGINDRAG, OnHdnBeginDrag)
		NOTIFY_CODE_HANDLER(HDN_ENDDRAG, OnHdnEndDrag)
		NOTIFY_CODE_HANDLER(HDN_BEGINTRACK, OnHdnBeginTrack)
		NOTIFY_CODE_HANDLER(HDN_ENDTRACK, OnHdnEndTrack)

		REFLECTED_NOTIFY_CODE_HANDLER(LVN_COLUMNCLICK, OnColumnClick)
		REFLECTED_NOTIFY_CODE_HANDLER(LVN_GETDISPINFO, OnGetDisplayInfo)
		REFLECTED_NOTIFY_CODE_HANDLER(LVN_ITEMCHANGED, OnItemChanged)
		REFLECTED_NOTIFY_CODE_HANDLER(LVN_ODSTATECHANGED, OnOdStateChanged)
		REFLECTED_NOTIFY_CODE_HANDLER(LVN_ITEMACTIVATE, OnListItemActivate)
		REFLECTED_NOTIFY_CODE_HANDLER(LVN_KEYDOWN, OnKeydown)
		REFLECTED_NOTIFY_CODE_HANDLER(LVN_GETINFOTIP , OnGetInfoTip)
		REFLECTED_NOTIFY_CODE_HANDLER(NM_CLICK, OnClick)
		REFLECTED_NOTIFY_CODE_HANDLER(NM_DBLCLK, OnClick)
		REFLECTED_NOTIFY_CODE_HANDLER(NM_CUSTOMDRAW, OnCustomDraw)
	END_MSG_MAP()

	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnSettingChange(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnButtonDown(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnMouseMove(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnMouseLeave(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnSetItemState(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnGetItemState(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnHdnBeginDrag(int, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnHdnEndDrag(int, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnHdnBeginTrack(int, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnHdnEndTrack(int, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnColumnClick(int, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnGetDisplayInfo(int, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnItemChanged(int, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnOdStateChanged(int, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnListItemActivate(int, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnKeydown(int, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnGetInfoTip(int, LPNMHDR, BOOL&);
	LRESULT OnClick(int, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnCustomDraw(int, LPNMHDR, BOOL&);

private:
	void GetEmptyMessageRect(LPRECT prc);
	void UpdateList();
	void UpdateSortIndicator(BOOL bShow);
	void UpdateViewVector(CVsiListItem *pList);

	void Fill(IVsiEnumStudy *pEnum, int iLevelMax) throw(...);
	void Fill(IVsiEnumSeries *pEnum, int iLevelMax) throw(...);
	void Fill(IVsiEnumImage *pEnum) throw(...);

	void AddStudy(IVsiStudy *pStudy, BOOL bSelect = FALSE) throw(...);
	void AddSeries(IVsiSeries *pSeries, BOOL bSelect = FALSE) throw(...);
	void AddImage(IVsiImage *pImage, BOOL bSelect = FALSE) throw(...);

	HRESULT GetSelectedStudies(
		IVsiEnumStudy **ppEnumStudies,
		int *piCount,
		CVsiListItem *pListSelected);
	HRESULT GetSelectedSeries(
		IVsiEnumSeries **ppEnumSeries,
		int *piCount,
		CVsiListItem *pListSelected);
	HRESULT GetSelectedImages(
		IVsiEnumImage **ppEnumImage,
		int *piCount,
		CVsiListItem *pListSelected);

	HRESULT RemoveDuplicateChildren(
		CVsiListItem &listParent,
		CVsiListItem &listChild);
	HRESULT RemoveDuplicateParent(
		CVsiListItem &listParent,
		CVsiListItem &listChild);
	HRESULT MoveToChildren(
		CVsiListItem &listParent,
		CVsiListItem &listChild,
		bool bVisibleOnly = true);

	HRESULT AdjustForActiveSeries(
		CVsiListItem &listIndexStudy,
		CVsiListItem &listIndexSeries,
		CVsiListItem &listIndexImage,
		IVsiSeries *pSeriesActive);
	HRESULT RemoveDuplicates(CVsiListItem &setIndex);

	void UpdateItemStatus(CVsiItem* pItem, BOOL bRecurChildren = FALSE) throw(...);
	void UpdateChildrenStatus(CVsiItem* pItem, BOOL bRecurChildren = FALSE) throw(...);
	void UpdateItemLocationMap(
		CVsiMapItemLocation &map,
		LPCWSTR pszParent,
		CVsiListItem &list) throw(...);
	CVsiItem* GetItem(LPCOLESTR pszNamespace, bool bVisibleOnly, BOOL *pbCancel);

	void GetNextFilterJob(CVsiFilterJob &job);
	bool Filter(
		CVsiListItem& listItem,
		CVsiFilterJob &job,
		BOOL *pbStop);
	bool Filter(CVsiItem *pItem, const std::vector<CString> &vecFilters, VSI_FILTER_TYPE type);
	bool FilterParent(CVsiItem *pItem, const std::vector<CString> &vecFilters, VSI_FILTER_TYPE type);

	void ExpandAll(CVsiListItem& listItem, BOOL &bStop);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiDataListWnd), CVsiDataListWnd)
