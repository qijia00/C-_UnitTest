/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudy.cpp
**
**	Description:
**		Implementation of CVsiStudy
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiBuildNumber.h>
#include <VsiCommUtl.h>
#include <VsiSaxUtils.h>
#include <VsiServiceKey.h>
#include <VsiLicenseUtils.h>
#include "VsiStudyXml.h"
#include "VsiStudy.h"


// Tag info
typedef struct tagVSI_STUDY_TAG
{
	DWORD dwId;
	VARTYPE type;
	BOOL bOptional;
	LPCWSTR pszName;
} VSI_STUDY_TAG;

// Must follow the order defined in IDL
static const VSI_STUDY_TAG g_pszXmlTag[] =
{
	{ VSI_PROP_STUDY_ID, VT_BSTR, FALSE, VSI_STUDY_XML_ATT_ID },
	{ VSI_PROP_STUDY_ID_CREATED, VT_BSTR, FALSE, VSI_STUDY_XML_ATT_ID_CREATED },
	{ VSI_PROP_STUDY_ID_COPIED, VT_BSTR, TRUE, VSI_STUDY_XML_ATT_ID_COPIED },
	{ VSI_PROP_STUDY_NS, VT_BSTR, TRUE, NULL },
	{ VSI_PROP_STUDY_NAME, VT_BSTR, FALSE, VSI_STUDY_XML_ATT_NAME },
	{ VSI_PROP_STUDY_LOCKED, VT_BOOL, TRUE, VSI_STUDY_XML_ATT_LOCK },
	{ VSI_PROP_STUDY_CREATED_DATE, VT_DATE, FALSE, VSI_STUDY_XML_ATT_CREATED_DATE },
	{ VSI_PROP_STUDY_OWNER, VT_BSTR, FALSE, VSI_STUDY_XML_ATT_OWNER },
	{ VSI_PROP_STUDY_EXPORTED, VT_BOOL, TRUE, VSI_STUDY_XML_ATT_EXPORTED },
	{ VSI_PROP_STUDY_NOTES, VT_BSTR, TRUE, VSI_STUDY_XML_ATT_NOTES },
	{ VSI_PROP_STUDY_GRANTING_INSTITUTION, VT_BSTR, TRUE, VSI_STUDY_XML_ATT_GRANTING_INSTITUTION },
	{ VSI_PROP_STUDY_VERSION_REQUIRED, VT_BSTR, TRUE, VSI_STUDY_XML_ATT_VERSION_REQUIRED },
	{ VSI_PROP_STUDY_VERSION_CREATED, VT_BSTR, TRUE, VSI_STUDY_XML_ATT_VERSION_CREATED },
	{ VSI_PROP_STUDY_VERSION_MODIFIED, VT_BSTR, TRUE, VSI_STUDY_XML_ATT_VERSION_MODIFIED },
	{ VSI_PROP_STUDY_ACCESS, VT_BSTR, TRUE, VSI_STUDY_XML_ATT_ACCESS },
	{ VSI_PROP_STUDY_LAST_MODIFIED_BY, VT_BSTR, TRUE, VSI_STUDY_XML_ATT_LAST_MODIFIED_BY },
	{ VSI_PROP_STUDY_LAST_MODIFIED_DATE, VT_DATE, TRUE, VSI_STUDY_XML_ATT_CREATED_LAST_MODIFIED_DATE },
};


CVsiStudy::CVsiStudy() :
	m_dwModified(0),
	m_dwErrorCode(VSI_STUDY_ERROR_NONE)
{

}

CVsiStudy::~CVsiStudy()
{
}

STDMETHODIMP CVsiStudy::Clone(IVsiStudy **ppStudy)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		CComPtr<IStream> pStm;
		CComPtr<IVsiStudy> pStudy;
		CComVariant v;

		hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IPersistStreamInit> pPersistDst(pStudy);
		VSI_CHECK_INTERFACE(pPersistDst, VSI_LOG_ERROR, NULL);
		CComQIPtr<IPersistStreamInit> pPersistSrc(this);
		VSI_CHECK_INTERFACE(pPersistDst, VSI_LOG_ERROR, NULL);

		// Parpare memory
		hr = CreateStreamOnHGlobal(NULL, TRUE, &pStm);
		VSI_CHECK_INTERFACE(pPersistDst, VSI_LOG_ERROR, NULL);

		// Copy parameter
		hr = pPersistSrc->Save(pStm, FALSE);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Move to beginning
		LARGE_INTEGER li;
		li.QuadPart = 0;
		hr = pStm->Seek(li, STREAM_SEEK_SET, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pPersistDst->InitNew();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pPersistDst->Load(pStm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		*ppStudy = pStudy.Detach();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::GetErrorCode(VSI_STUDY_ERROR *pdwErrorCode)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pdwErrorCode, VSI_LOG_ERROR, L"pdwErrorCode");

		*pdwErrorCode = m_dwErrorCode;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::GetDataPath(LPOLESTR *ppszPath)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(ppszPath, VSI_LOG_ERROR, L"ppszPath");

		*ppszPath = AtlAllocTaskWideString((LPCWSTR)m_strFile);
		VSI_CHECK_MEM(*ppszPath, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::SetDataPath(LPCOLESTR pszPath)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszPath, VSI_LOG_ERROR, L"pszPath");
		m_strFile = pszPath;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::Compare(IVsiStudy *pStudy, VSI_PROP_STUDY dwProp, int *piCmp)
{
	HRESULT hr(S_OK);
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pStudy, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(piCmp, VSI_LOG_ERROR, NULL);

		*piCmp = 0;

		HRESULT hr1, hr2;
		CComVariant v1, v2;

		switch (dwProp)
		{
		case VSI_PROP_STUDY_LOCKED:
			{
				hr1 = GetProperty(VSI_PROP_STUDY_LOCKED, &v1);
				hr2 = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &v2);
				if (S_OK == hr1 && S_OK == hr2)
				{
					if (VARIANT_FALSE == V_BOOL(&v1) && VARIANT_FALSE != V_BOOL(&v2))
					{
						*piCmp = -1;
					}
					else if (VARIANT_FALSE != V_BOOL(&v1) && VARIANT_FALSE == V_BOOL(&v2))
					{
						*piCmp = 1;
					}
				}
				else if (S_OK == hr1)
				{
					*piCmp = 1;
				}
				else if (S_OK == hr2)
				{
					*piCmp = -1;
				}
			}
			break;

		case VSI_PROP_STUDY_NS:
		case VSI_PROP_STUDY_NAME:
		case VSI_PROP_STUDY_OWNER:
			{
				hr1 = GetProperty(dwProp, &v1);
				hr2 = pStudy->GetProperty(dwProp, &v2);
				if (S_OK == hr1 && S_OK == hr2)
				{
					*piCmp = lstrcmp(V_BSTR(&v1), V_BSTR(&v2));
				}
				else if (S_OK == hr1)
				{
					*piCmp = 1;
				}
				else if (S_OK == hr2)
				{
					*piCmp = -1;
				}
			}
			break;

		case VSI_PROP_STUDY_CREATED_DATE:
			// Handle below
			break;

		default:
			_ASSERT(0);
		}  // switch

		// Same
		if (0 == *piCmp)
		{
			// Compare date
			hr1 = GetProperty(VSI_PROP_STUDY_CREATED_DATE, &v1);
			hr2 = pStudy->GetProperty(VSI_PROP_STUDY_CREATED_DATE, &v2);
			if (S_OK == hr1 && S_OK == hr2)
			{
				if (v1 < v2)
					*piCmp = -1;
				else if (v2 < v1)
					*piCmp = 1;
			}
			else if (S_OK == hr1)
			{
				*piCmp = 1;
			}
			else if (S_OK == hr2)
			{
				*piCmp = -1;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::GetProperty(VSI_PROP_STUDY dwProp, VARIANT* pvValue)
{
	CCritSecLock csl(m_cs);

	CVsiMapPropIter iterProp = m_mapProp.find(dwProp);
	if (iterProp == m_mapProp.end())
		return S_FALSE;

	return VariantCopy(pvValue, &iterProp->second);
}

STDMETHODIMP CVsiStudy::SetProperty(VSI_PROP_STUDY dwProp, const VARIANT *pvValue)
{
	CCritSecLock csl(m_cs);

	m_mapProp.erase(dwProp);

	if (NULL != pvValue && VT_EMPTY != V_VT(pvValue))
		m_mapProp.insert(CVsiMapProp::value_type(dwProp, *pvValue));

	return S_OK;
}

STDMETHODIMP CVsiStudy::AddSeries(IVsiSeries *pSeries)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pSeries, VSI_LOG_ERROR, L"pSeries");

		hr = pSeries->SetStudy(this);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vId;
		hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vId);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		_ASSERT(VT_BSTR == V_VT(&vId));

		if (0 == m_mapSeries.size() && 0 < m_strFile.Length())
		{
			hr = LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST, nullptr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		m_mapSeries.erase(CString(V_BSTR(&vId)));
		m_mapSeries[CString(V_BSTR(&vId))] = pSeries;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::RemoveSeries(LPCOLESTR pszId)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszId, VSI_LOG_ERROR, L"pszId");

		m_mapSeries.erase(pszId);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::RemoveAllSeries()
{
	CCritSecLock csl(m_cs);

	m_mapSeries.clear();

	return S_OK;
}

STDMETHODIMP CVsiStudy::GetSeriesCount(LONG *plCount, BOOL *pbCancel)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(plCount, VSI_LOG_ERROR, L"plCount");

		if (0 == m_mapSeries.size())
		{
			hr = LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST | VSI_DATA_TYPE_GET_COUNT_ONLY, pbCancel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			*plCount = (LONG)m_mapSeries.size();

			m_mapSeries.clear();
		}
		else
		{
			*plCount = (LONG)m_mapSeries.size();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::GetSeriesEnumerator(
	VSI_PROP_SERIES sortBy,
	BOOL bDescending,
	IVsiEnumSeries **ppEnum,
	BOOL *pbCancel)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

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
	bool bLoadSeries(false);

	try
	{
		VSI_CHECK_ARG(ppEnum, VSI_LOG_ERROR, L"ppEnum");

		*ppEnum = NULL;

		if (0 == m_mapSeries.size())
		{
			hr = LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST, pbCancel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			bLoadSeries = true;
		}

		if (0 == m_mapSeries.size())
		{
			return S_FALSE;
		}

		if (pbCancel && *pbCancel)
		{
			if (bLoadSeries)
			{
				m_mapSeries.clear();
			}
			return S_FALSE;
		}

		ppSeries = new IVsiSeries*[m_mapSeries.size()];
		VSI_CHECK_MEM(ppSeries, VSI_LOG_ERROR, NULL);

		if (VSI_PROP_SERIES_MIN != sortBy)
		{
			CVsiSeries::CVsiSort sort;
			sort.m_sortBy = sortBy;
			sort.m_bDescending = bDescending;

			CVsiListSeries listSeries;

			for (auto iterMap = m_mapSeries.begin();
				iterMap != m_mapSeries.end();
				++iterMap)
			{
				IVsiSeries *pSeries = iterMap->second;

				CVsiSeries series(pSeries, &sort);
				listSeries.push_back(series);
			}

			if (pbCancel && *pbCancel)
			{
				return S_FALSE;
			}

			listSeries.sort();

			if (pbCancel && *pbCancel)
			{
				return S_FALSE;
			}

			for (auto iter = listSeries.begin();
				iter != listSeries.end();
				++iter)
			{
				CVsiSeries &series = *iter;

				CComPtr<IVsiSeries> pSeries(series.m_pSeries);
				ppSeries[i++] = pSeries.Detach();
			}
		}
		else
		{
			for (auto iter = m_mapSeries.cbegin();
				iter != m_mapSeries.cend();
				iter++)
			{
				CComPtr<IVsiSeries> pSeries(iter->second);
				ppSeries[i++] = pSeries.Detach();
			}
		}

		pEnum = new CVsiEnumSeries;
		VSI_CHECK_MEM(pEnum, VSI_LOG_ERROR, NULL);

		hr = pEnum->Init(ppSeries, ppSeries + i, NULL, AtlFlagTakeOwnership);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		ppSeries = NULL;  // pEnum is the owner

		hr = pEnum->QueryInterface(
			__uuidof(IVsiEnumSeries),
			(void**)ppEnum
		);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		pEnum = NULL;
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

	if (bLoadSeries)
	{
		m_mapSeries.clear();
	}

	return hr;
}

STDMETHODIMP CVsiStudy::GetSeries(LPCOLESTR pszId, IVsiSeries **ppSeries)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszId, VSI_LOG_ERROR, L"pszId");
		VSI_CHECK_ARG(ppSeries, VSI_LOG_ERROR, L"ppSeries");

		BOOL bLoadSeries(FALSE);

		if (0 == m_mapSeries.size())
		{
			hr = LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST, nullptr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			bLoadSeries = TRUE;
		}

		auto iter = m_mapSeries.find(pszId);
		if (iter != m_mapSeries.end())
		{
			*ppSeries = iter->second;
			(*ppSeries)->AddRef();

			hr = S_OK;
		}

		if (bLoadSeries)
		{
			m_mapSeries.clear();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::GetItem(LPCOLESTR pszNs, IUnknown **ppUnk)
{
	HRESULT hr(S_FALSE);
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");
		VSI_CHECK_ARG(ppUnk, VSI_LOG_ERROR, L"ppUnk");

		*ppUnk = NULL;

		BOOL bLoadSeries(FALSE);

		if (0 == m_mapSeries.size())
		{
			hr = LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST, nullptr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			bLoadSeries = TRUE;
		}

		CComPtr<IVsiSeries> pSeries;

		CString strId;
		LPCWSTR pszNext = pszNs;
		LPCWSTR pszSpt;

		if (0 != *pszNext)
		{
			pszSpt = wcschr(pszNext, L'/');
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

			auto iter = m_mapSeries.find(strId);
			if (iter != m_mapSeries.end())
			{
				pSeries = iter->second;

				if (NULL != pszNext && 0 != *pszNext)
				{
					hr = pSeries->GetItem(pszNext, ppUnk);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
				else
				{
					*ppUnk = pSeries.Detach();
					hr = S_OK;
				}
			}
		}
		
		if (bLoadSeries)
		{
			m_mapSeries.clear();
		}

		hr = (NULL != *ppUnk) ? S_OK : S_FALSE;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::LoadStudyData(LPCOLESTR pszFile, DWORD dwFlags, BOOL *pbCancel)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		if (NULL != pszFile)
		{
			m_strFile = pszFile;

			m_mapProp.clear();

			m_dwErrorCode |= VSI_STUDY_ERROR_LOAD_XML;

			CComPtr<ISAXXMLReader> pReader;
			hr = VsiCreateXmlSaxReader(&pReader);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pReader->putContentHandler(static_cast<ISAXContentHandler*>(this));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CVsiSAXErrorHandler err;
			hr = pReader->putErrorHandler(&err);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pReader->parseURL(pszFile);
			if (FAILED(hr))
			{
				CComVariant vName;
				if (S_OK != GetProperty(VSI_PROP_STUDY_NAME, &vName))
				{
					CComVariant vError(CString(MAKEINTRESOURCE(IDS_STUDY_CORRUPTED_STUDY)));
					hr = SetProperty(VSI_PROP_STUDY_NAME, &vError);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				// Set Id and NS in the case of a failure
				CComVariant vId;
				if (S_OK == GetProperty(VSI_PROP_STUDY_ID, &vId))
				{
					hr = SetProperty(VSI_PROP_STUDY_NS, &vId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
				else
				{
					CComVariant vPath(m_strFile);

					hr = SetProperty(VSI_PROP_STUDY_ID, &vPath);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = SetProperty(VSI_PROP_STUDY_NS, &vPath);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				WCHAR szMsg[300];
				err.FormatErrorMsg(szMsg, _countof(szMsg));
				VSI_LOG_MSG(VSI_LOG_ERROR | VSI_LOG_NO_BOX, VsiFormatMsg(L"Read file '%s' failed", pszFile));
				VSI_FAIL(VSI_LOG_ERROR | VSI_LOG_NO_BOX, VsiFormatMsg(L"SAXErrorHandler - %s", szMsg));
			}

			// Sets namespace
			{
				CComVariant vId;
				if (S_OK == GetProperty(VSI_PROP_STUDY_ID, &vId))
				{
					hr = SetProperty(VSI_PROP_STUDY_NS, &vId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}

			// Patch VSI_PROP_STUDY_ACCESS - new in Sierra VI
			{
				CComVariant vAccess;
				hr = GetProperty(VSI_PROP_STUDY_ACCESS, &vAccess);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_FALSE == hr)
				{
					vAccess = VSI_STUDY_ACCESS_PUBLIC;
					hr = SetProperty(VSI_PROP_STUDY_ACCESS, &vAccess);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}

			m_dwErrorCode &= ~VSI_STUDY_ERROR_LOAD_XML;

			// Checks version
			CComVariant vVersionRequired;
			HRESULT hrTest = GetProperty(VSI_PROP_STUDY_VERSION_REQUIRED, &vVersionRequired);
			VSI_CHECK_HR(hrTest, VSI_LOG_ERROR, NULL);
			if (S_OK == hrTest)
			{
				CString strSoftware;
				strSoftware.Format(
					L"%u.%u.%u.%u",
					VSI_SOFTWARE_MAJOR,
					VSI_SOFTWARE_MIDDLE,
					VSI_SOFTWARE_MINOR,
					VSI_BUILD_NUMBER);

				int iRet = VsiVersionCompare(strSoftware, V_BSTR(&vVersionRequired));
				if (iRet < 0)
				{
					m_dwErrorCode |= VSI_STUDY_ERROR_VERSION_INCOMPATIBLE;
				}
			}
		}

		if (VSI_DATA_TYPE_SERIES_LIST & dwFlags)
		{
			// Skip if already loaded
			if (0 == m_mapSeries.size())
			{
				m_dwErrorCode |= VSI_STUDY_ERROR_LOAD_SERIES;

				// Get session
				CComPtr<IVsiSession> pSession;

				if (0 == (VSI_DATA_TYPE_NO_SESSION_LINK & dwFlags) &&
					0 == (VSI_DATA_TYPE_GET_COUNT_ONLY & dwFlags))
				{
					CComQIPtr<IVsiServiceProvider> pServiceProvider(g_pApp);
					VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

					hr = pServiceProvider->QueryService(
						VSI_SERVICE_SESSION_MGR,
						__uuidof(IVsiSession),
						(IUnknown**)&pSession);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				// Load series
				m_mapSeries.clear();

				if (NULL == pszFile)
				{
					VSI_VERIFY(0 < m_strFile.Length(), VSI_LOG_ERROR, NULL);
					pszFile = m_strFile;
				}

				LPCWSTR pszSpt = wcsrchr(pszFile, L'\\');
				CString strPath(pszFile, (int)(pszSpt - pszFile));

				// Find series
				CString strFilter;
				strFilter.Format(L"%s\\*", (LPCWSTR)strPath);

				BOOL bError = FALSE;
				CString strSeriesFile;
				LPWSTR pszDir;
				CVsiDirectoryIterator dirIter(strFilter);
				while (NULL != (pszDir = dirIter.Next()))
				{
					strSeriesFile.Format(L"%s\\%s\\%s", (LPCWSTR)strPath, pszDir, CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME)));

					if (0 == _waccess(strSeriesFile, 0))
					{
						if (0 == (VSI_DATA_TYPE_GET_COUNT_ONLY & dwFlags))
						{
							CComPtr<IVsiSeries> pSeries;
							hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							hr = pSeries->SetStudy(static_cast<IVsiStudy*>(this));
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							CComVariant vId;
							CComQIPtr<IVsiPersistSeries> pPersist(pSeries);
							hr = pPersist->LoadSeriesData(strSeriesFile, dwFlags, NULL);
							if (FAILED(hr))
							{
								bError = TRUE;
								VSI_LOG_MSG(VSI_LOG_WARNING,
									VsiFormatMsg(L"Load series failed - '%s'",
									(LPCWSTR)strSeriesFile));
								hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vId);
								if (S_OK != hr)
								{
									// Cannot even get the id, use file path as key
									vId = strSeriesFile;
								}
							}
							else
							{
								hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vId);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								if (S_OK != hr)
								{
									// Cannot even get the id, use file path as key
									vId = strSeriesFile;
								}
							}

							if (NULL != pSession)
							{
								CComVariant vNs;
								hr = pSeries->GetProperty(VSI_PROP_SERIES_NS, &vNs);
								if (S_OK == hr)
								{
									int iSlot(-1);
									if (S_OK == pSession->GetIsSeriesLoaded(V_BSTR(&vNs), &iSlot))
									{
										pSeries.Release();
										hr = pSession->GetSeries(iSlot, &pSeries);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
									}
								}
							}

							m_mapSeries[CString(V_BSTR(&vId))] = pSeries;
						}
						else
						{
							m_mapSeries[pszDir] = NULL;
						}
					}

					if (pbCancel && (FALSE != *pbCancel))
					{
						m_mapSeries.clear();
						break;
					}
				}

				if (!bError)
				{
					m_dwErrorCode &= ~VSI_STUDY_ERROR_LOAD_SERIES;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::SaveStudyData(LPCOLESTR pszFile, DWORD dwFlags)
{
	UNREFERENCED_PARAMETER(dwFlags);

	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);
	HANDLE hTransaction = INVALID_HANDLE_VALUE;

	try
	{
		if (NULL != pszFile)
			m_strFile = pszFile;

		hTransaction = VsiTxF::VsiCreateTransaction(
			NULL,
			0,		
			TRANSACTION_DO_NOT_PROMOTE,	
			0,		
			0,		
			INFINITE,
			L"Study xml save transaction");
		VSI_WIN32_VERIFY(INVALID_HANDLE_VALUE != hTransaction, VSI_LOG_ERROR, L"Failure creating transaction");

		// Create a backup of the original file
		WCHAR szBackupFileName[VSI_MAX_PATH];
		errno_t err = wcscpy_s(szBackupFileName, m_strFile);
		VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");
		VsiPathRemoveExtension(szBackupFileName);
		err = wcscat_s(szBackupFileName, L".bak");
		VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");

		if (0 == _waccess(m_strFile, 0))
		{
			BOOL bRet = VsiTxF::VsiCopyFileTransacted(
				m_strFile,
				szBackupFileName,
				NULL,
				NULL,
				NULL,
				0,
				hTransaction);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, NULL);
		}

		// Create a temporary file using the SaveStudyData, then use the transacted 
		// file methods to move the temp file back to the original.
		// Use commutl to create unique temp file.
		WCHAR szTempFileName[VSI_MAX_PATH];
		err = wcscpy_s(szTempFileName, m_strFile);
		VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");
		VsiPathRemoveExtension(szTempFileName);

		WCHAR szUniqueName[VSI_MAX_PATH];
		hr = VsiCreateUniqueName(L"", L"", szUniqueName, VSI_MAX_PATH);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting unique name");

		err = wcscat_s(szTempFileName, szUniqueName);
		VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");
		err = wcscat_s(szTempFileName, L".xml");
		VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");

		// Saves operator list
		hr = SaveStudyDataInternal(szTempFileName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		BOOL bRet = VsiTxF::VsiMoveFileTransacted(
			szTempFileName,
			m_strFile,	
			NULL,	
			NULL,	
			MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING,
			hTransaction);
		VSI_VERIFY(bRet, VSI_LOG_ERROR, L"Move file failed");

		bRet = VsiTxF::VsiCommitTransaction(hTransaction);
		VSI_VERIFY(bRet, VSI_LOG_ERROR, L"Create transaction failed");
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

	return hr;
}

HRESULT CVsiStudy::SaveStudyDataInternal(LPCOLESTR pszFile)
{
	HRESULT hr = S_OK;

	try
	{
		// Add version required
		CComVariant vVersionRequired(VSI_STUDY_VERSION_REQUIRED_VEVO2100);
		hr = SetProperty(VSI_PROP_STUDY_VERSION_REQUIRED, &vVersionRequired);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	
		// Add version modified
		CString strVersion;
		strVersion.Format(
			L"%u.%u.%u.%u",
			VSI_SOFTWARE_MAJOR, VSI_SOFTWARE_MIDDLE, VSI_SOFTWARE_MINOR, VSI_BUILD_NUMBER);		

		CComVariant vVersionModified(strVersion.GetString());
		hr = SetProperty(VSI_PROP_STUDY_VERSION_MODIFIED, &vVersionModified);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Save
		CComPtr<IMXWriter> pWriter;
		hr = VsiCreateXmlSaxWriter(&pWriter);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pWriter->put_indent(VARIANT_TRUE);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pWriter->put_encoding(CComBSTR(L"utf-8"));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CVsiMXFileWriterStream out;
		hr = out.Open(pszFile);
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

		CVsiMapPropIter iter;
		DWORD dwTagXmlIndex = 0;
		for (iter = m_mapProp.begin(); iter != m_mapProp.end(); ++iter)
		{
			dwTagXmlIndex = iter->first - VSI_PROP_STUDY_MIN - 1;
			_ASSERT(g_pszXmlTag[dwTagXmlIndex].dwId == iter->first);

			if (NULL == g_pszXmlTag[dwTagXmlIndex].pszName)
				continue;  // No persist

			if (VT_EMPTY == V_VT(&(iter->second)))
				continue;

			CComVariant v;

			hr = VsiXmlVariantToVariantXml(&(iter->second), &v);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComBSTR bstrName(g_pszXmlTag[dwTagXmlIndex].pszName);
			hr = pmxa->addAttribute(bstrNull, bstrName, bstrName, bstrNull, V_BSTR(&v));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		hr = pch->startElement(szNull, 0, szNull, 0,
			VSI_STUDY_XML_ELM_STUDY, VSI_STATIC_STRING_COUNT(VSI_STUDY_XML_ELM_STUDY), pa.p);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pch->endElement(szNull, 0, szNull, 0,
			VSI_STUDY_XML_ELM_STUDY, VSI_STATIC_STRING_COUNT(VSI_STUDY_XML_ELM_STUDY));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pch->endDocument();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		out.Close();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::startElement(
	const wchar_t *pwszNamespaceUri,
	int cchNamespaceUri,
	const wchar_t *pwszLocalName,
	int cchLocalName,
	const wchar_t *pwszQName,
	int cchQName,
	ISAXAttributes *pAttributes)
{
	UNREFERENCED_PARAMETER(pwszNamespaceUri);
	UNREFERENCED_PARAMETER(cchNamespaceUri);
	UNREFERENCED_PARAMETER(pwszQName);
	UNREFERENCED_PARAMETER(cchQName);

	HRESULT hr = S_OK;
	BOOL bMatched;
	LPCWSTR pszValue;
	int iValue;
	CComVariant v;
	CString strValue;

	try
	{
		// study?
		bMatched = VsiXmlIsEqualElement(
			pwszLocalName,
			cchLocalName,
			VSI_STUDY_XML_ELM_STUDY,
			VSI_STATIC_STRING_COUNT(VSI_STUDY_XML_ELM_STUDY)
		);
		if (bMatched)
		{
			for (int i = 0; i < _countof(g_pszXmlTag); ++i)
			{
				if (VT_EMPTY == g_pszXmlTag[i].type)
					continue;  // Not used

				if (NULL == g_pszXmlTag[i].pszName)
					continue;  // No persist

				hr = pAttributes->getValueFromName(
					L"", 0,
					g_pszXmlTag[i].pszName, lstrlen(g_pszXmlTag[i].pszName),
					&pszValue, &iValue);
				if (hr != S_OK)
				{
					// Checks if tag is optional
					VSI_VERIFY(g_pszXmlTag[i].bOptional, VSI_LOG_ERROR | VSI_LOG_NO_BOX, L"Attribute missing");
					hr = S_OK;
					continue;
				}

				VsiReadProperty(
					g_pszXmlTag[i].dwId,
					g_pszXmlTag[i].type,
					pszValue, iValue,
					m_mapProp);
			}

			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::InitNew()
{
	HRESULT hr = S_OK;

	try
	{
		m_strFile.Empty();
		m_dwErrorCode = VSI_STUDY_ERROR_NONE;
		m_mapProp.clear();
		m_mapSeries.clear();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::GetClassID(CLSID *pClassID)
{
	if (pClassID == NULL)
		return E_POINTER;

	*pClassID = GetObjectCLSID();

	return S_OK;
}

STDMETHODIMP CVsiStudy::IsDirty()
{
	return (m_dwModified > 0) ? S_OK : S_FALSE;
}

STDMETHODIMP CVsiStudy::Load(LPSTREAM pStm)
{
	HRESULT hr = S_OK;
	ULONG ulRead;

	try
	{
		// Series not loaded

		hr = m_strFile.ReadFromStream(pStm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pStm->Read(&m_dwModified, sizeof(m_dwModified), &ulRead);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pStm->Read(&m_dwErrorCode, sizeof(m_dwErrorCode), &ulRead);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		DWORD dwCount;
		hr = pStm->Read(&dwCount, sizeof(dwCount), &ulRead);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		for (DWORD i = 0; i < dwCount; ++i)
		{
			DWORD dwProp;
			CComVariant v;

			hr = pStm->Read(&dwProp, sizeof(dwProp), &ulRead);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = v.ReadFromStream(pStm);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_mapProp.insert(CVsiMapProp::value_type(dwProp, v));
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::Save(LPSTREAM pStm, BOOL fClearDirty)
{
	UNREFERENCED_PARAMETER(fClearDirty);

	HRESULT hr = S_OK;
	ULONG ulWritten;

	try
	{
		// Series not saved

		hr = m_strFile.WriteToStream(pStm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pStm->Write(&m_dwModified, sizeof(m_dwModified), &ulWritten);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pStm->Write(&m_dwErrorCode, sizeof(m_dwErrorCode), &ulWritten);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		DWORD dwCount = (DWORD)m_mapProp.size();
		hr = pStm->Write(&dwCount, sizeof(dwCount), &ulWritten);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CVsiMapPropIter iter;
		for (iter = m_mapProp.begin(); iter != m_mapProp.end(); ++iter)
		{
			DWORD dwProp = iter->first;
			CComVariant &v = iter->second;

			hr = pStm->Write(&dwProp, sizeof(dwProp), &ulWritten);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = v.WriteToStream(pStm);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudy::GetSizeMax(ULARGE_INTEGER* pcbSize)
{
	UNREFERENCED_PARAMETER(pcbSize);

	return E_NOTIMPL;
}

