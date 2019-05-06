/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiSeries.cpp
**
**	Description:
**		Implementation of CVsiSeries
**
*******************************************************************************/

#include "stdafx.h"
#include <ATLComTime.h>
#include <VsiBuildNumber.h>
#include <IVsiServiceProviderImpl.h>
#include <VsiCommUtl.h>
#include <VsiServiceKey.h>
#include <VsiModeModule.h>
#include <VsiAnalRsltXmlTags.h>
#include <VsiLicenseUtils.h>
#include "VsiStudyXml.h"
#include "VsiSeries.h"


// Tag info
typedef struct tagVSI_SERIES_TAG
{
	DWORD dwId;
	VARTYPE type;
	BOOL bOptional;
	LPCWSTR pszName;
} VSI_SERIES_TAG;

// Must follow the order defined in IDL
static const VSI_SERIES_TAG g_pszXmlTag[] =
{
	{ VSI_PROP_SERIES_ID, VT_BSTR, FALSE, VSI_SERIES_XML_ATT_ID },
	{ VSI_PROP_SERIES_ID_STUDY, VT_BSTR, FALSE, VSI_SERIES_XML_ATT_ID_STUDY },
	{ VSI_PROP_SERIES_NS, VT_BSTR, TRUE, NULL },
	{ VSI_PROP_SERIES_NAME, VT_BSTR, FALSE, VSI_SERIES_XML_ATT_NAME },
	{ VSI_PROP_SERIES_CREATED_DATE, VT_DATE, FALSE, VSI_SERIES_XML_ATT_CREATED_DATE },
	{ VSI_PROP_SERIES_ACQUIRED_BY, VT_BSTR, FALSE, VSI_SERIES_XML_ATT_ACQUIRED_BY },
	{ VSI_PROP_SERIES_EXPORTED, VT_BOOL, TRUE, VSI_SERIES_XML_ATT_EXPORTED },
	{ VSI_PROP_SERIES_VERSION_REQUIRED, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_VERSION_REQUIRED },
	{ VSI_PROP_SERIES_VERSION_CREATED, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_VERSION_CREATED },
	{ VSI_PROP_SERIES_VERSION_MODIFIED, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_VERSION_MODIFIED },
	{ VSI_PROP_SERIES_NOTES, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_NOTES },
	{ VSI_PROP_SERIES_APPLICATION, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_APPLICATION },
	{ VSI_PROP_SERIES_MSMNT_PACKAGE, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_MSMNT_PACKAGE },
	{ VSI_PROP_SERIES_ANIMAL_ID, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_ANIMAL_ID },
	{ VSI_PROP_SERIES_ANIMAL_COLOR, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_ANIMAL_COLOR },
	{ VSI_PROP_SERIES_STRAIN, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_STRAIN },
	{ VSI_PROP_SERIES_SOURCE, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_SOURCE },
	{ VSI_PROP_SERIES_WEIGHT, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_WEIGHT },
	{ VSI_PROP_SERIES_TYPE, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_TYPE },
	{ VSI_PROP_SERIES_DATE_OF_BIRTH, VT_DATE, TRUE, VSI_SERIES_XML_ATT_DATE_OF_BIRTH },
	{ VSI_PROP_SERIES_HEART_RATE, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_HEART_RATE },
	{ VSI_PROP_SERIES_BODY_TEMP, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_BODY_TEMP },
	{ VSI_PROP_SERIES_SEX, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_SEX },
	{ VSI_PROP_SERIES_PREGNANT, VT_BOOL, TRUE, VSI_SERIES_XML_ATT_PREGNANT },
	{ VSI_PROP_SERIES_DATE_MATED, VT_DATE, TRUE, VSI_SERIES_XML_ATT_DATE_MATED },
	{ VSI_PROP_SERIES_DATE_PLUGGED, VT_DATE, TRUE, VSI_SERIES_XML_ATT_DATE_PLUGGED },
	{ VSI_PROP_SERIES_CUSTOM1, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM1 },
	{ VSI_PROP_SERIES_CUSTOM2, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM2 },
	{ VSI_PROP_SERIES_CUSTOM3, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM3 },
	{ VSI_PROP_SERIES_CUSTOM4, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM4 },
	{ VSI_PROP_SERIES_CUSTOM5, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM5 },
	{ VSI_PROP_SERIES_CUSTOM6, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM6 },
	{ VSI_PROP_SERIES_CUSTOM7, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM7 },
	{ VSI_PROP_SERIES_CUSTOM8, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM8 },
	{ VSI_PROP_SERIES_CUSTOM9, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM9 },
	{ VSI_PROP_SERIES_CUSTOM10, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM10 },
	{ VSI_PROP_SERIES_CUSTOM11, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM11 },
	{ VSI_PROP_SERIES_CUSTOM12, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM12 },
	{ VSI_PROP_SERIES_CUSTOM13, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM13 },
	{ VSI_PROP_SERIES_CUSTOM14, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_CUSTOM14 },
	{ VSI_PROP_SERIES_COLOR_INDEX, VT_I4, TRUE, VSI_SERIES_XML_ATT_COLOR_INDEX },
	{ VSI_PROP_SERIES_LAST_MODIFIED_BY, VT_BSTR, TRUE, VSI_SERIES_XML_ATT_LAST_MODIFIED_BY },
	{ VSI_PROP_SERIES_LAST_MODIFIED_DATE, VT_DATE, TRUE, VSI_SERIES_XML_ATT_CREATED_LAST_MODIFIED_DATE },
};


CVsiSeries::CVsiSeries() :
	m_pStudy(NULL),
	m_dwModified(0),
	m_dwErrorCode(VSI_SERIES_ERROR_NONE)
{
}

CVsiSeries::~CVsiSeries()
{
}

STDMETHODIMP CVsiSeries::Clone(IVsiSeries **ppSeries)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		CComPtr<IStream> pStm;
		CComPtr<IVsiSeries> pSeries;
		CComVariant v;

		hr = pSeries.CoCreateInstance(__uuidof(VsiSeries));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IPersistStreamInit> pPersistDst(pSeries);
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

		*ppSeries = pSeries.Detach();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::GetErrorCode(VSI_SERIES_ERROR *pdwErrorCode)
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

STDMETHODIMP CVsiSeries::GetDataPath(LPOLESTR *ppszPath)
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

STDMETHODIMP CVsiSeries::SetDataPath(LPCOLESTR pszPath)
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

STDMETHODIMP CVsiSeries::GetStudy(IVsiStudy **ppStudy)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(ppStudy, VSI_LOG_ERROR, L"ppStudy");

		*ppStudy = m_pStudy;
		if (m_pStudy != NULL)
			m_pStudy->AddRef();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::SetStudy(IVsiStudy *pStudy)
{
	HRESULT hr;
	CCritSecLock csl(m_cs);

	try
	{
		m_pStudy = pStudy;

		// Updates namespace
		if (m_pStudy != NULL)
		{
			UpdateNamespace();
		}
		else
		{
			// Removes namespace
			hr = SetProperty(VSI_PROP_SERIES_NS, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return S_OK;
}

STDMETHODIMP CVsiSeries::Compare(IVsiSeries *pSeries, VSI_PROP_SERIES dwProp, int *piCmp)
{
	HRESULT hr(S_OK);
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(piCmp, VSI_LOG_ERROR, NULL);
		*piCmp = 0;

		HRESULT hr1, hr2;
		CComVariant v1, v2;

		switch (dwProp)
		{
		case VSI_PROP_SERIES_NAME:
		case VSI_PROP_SERIES_ACQUIRED_BY:
		case VSI_PROP_SERIES_ANIMAL_ID:
		case VSI_PROP_SERIES_CUSTOM4:
			{
				hr1 = GetProperty(dwProp, &v1);
				hr2 = pSeries->GetProperty(dwProp, &v2);
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

		case VSI_PROP_SERIES_CREATED_DATE:
			// Handle below
			break;

		default:
			_ASSERT(0);
		}  // switch

		if (*piCmp == 0)
		{
			// If names are the same then compare by date
			hr1 = GetProperty(VSI_PROP_SERIES_CREATED_DATE, &v1);
			hr2 = pSeries->GetProperty(VSI_PROP_SERIES_CREATED_DATE, &v2);
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

STDMETHODIMP CVsiSeries::GetProperty(VSI_PROP_SERIES dwProp, VARIANT *pvValue)
{
	CCritSecLock csl(m_cs);

	CVsiMapPropIter iterProp = m_mapProp.find(dwProp);
	if (iterProp == m_mapProp.end())
	{
		if (VSI_PROP_SERIES_NS == dwProp)
		{
			// Namespace not defined yet - try construct it now
			UpdateNamespace();

			iterProp = m_mapProp.find(dwProp);
			if (iterProp == m_mapProp.end())
				return S_FALSE;
		}
		else
		{
			return S_FALSE;
		}
	}

	return VariantCopy(pvValue, &iterProp->second);
}

STDMETHODIMP CVsiSeries::SetProperty(VSI_PROP_SERIES dwProp, const VARIANT *pvValue)
{
	CCritSecLock csl(m_cs);

	m_mapProp.erase(dwProp);

	if (NULL != pvValue && VT_EMPTY != V_VT(pvValue))
		m_mapProp.insert(CVsiMapProp::value_type(dwProp, *pvValue));

	return S_OK;
}

STDMETHODIMP CVsiSeries::GetPropertyLabel(
	VSI_PROP_SERIES dwProp,
	LPOLESTR *ppszLabel)
{
	HRESULT hr(S_OK);
	CCritSecLock csl(m_cs);

	try
	{
		const VSI_SERIES_TAG &tag = g_pszXmlTag[dwProp - VSI_PROP_SERIES_MIN - 1];

		auto iterProp = m_mapPropLabel.find(tag.pszName);
		if (iterProp == m_mapPropLabel.end())
			return S_FALSE;

		const CString &strLabel = iterProp->second;

		*ppszLabel = AtlAllocTaskWideString(strLabel.GetString());
		VSI_CHECK_MEM(*ppszLabel, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::AddImage(IVsiImage *pImage)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pImage, VSI_LOG_ERROR, L"pImage");

		hr = pImage->SetSeries(this);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vId;
		hr = pImage->GetProperty(VSI_PROP_IMAGE_ID, &vId);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		_ASSERT(VT_BSTR == V_VT(&vId));

		if (0 == m_mapImage.size() && 0 < m_strFile.Length())
		{
			hr = LoadSeriesData(NULL, VSI_DATA_TYPE_IMAGE_LIST, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		m_mapImage.erase(CString(V_BSTR(&vId)));
		m_mapImage.insert(CVsiMapImage::value_type(
			CString(V_BSTR(&vId)), pImage));
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::RemoveImage(LPCOLESTR pszId)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszId, VSI_LOG_ERROR, L"pszId");

		m_mapImage.erase(pszId);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::RemoveAllImages()
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		m_mapImage.clear();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::GetImageCount(LONG *plCount, BOOL *pbCancel)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(plCount, VSI_LOG_ERROR, L"plCount");

		if (0 == m_mapImage.size())
		{
			if (0 < m_strFile.Length())
			{
				hr = LoadSeriesData(NULL, VSI_DATA_TYPE_IMAGE_LIST | VSI_DATA_TYPE_GET_COUNT_ONLY, pbCancel);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				*plCount = (LONG)m_mapImage.size();
				
				m_mapImage.clear();
			}
			else
			{
				*plCount = 0;
			}
		}
		else
		{
			*plCount = (LONG)m_mapImage.size();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::GetImageEnumerator(
	VSI_PROP_IMAGE sortBy,
	BOOL bDescending,
	IVsiEnumImage **ppEnum,
	BOOL *pbCancel)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	typedef CComObject<
		CComEnum<
			IVsiEnumImage,
			&__uuidof(IVsiEnumImage),
			IVsiImage*, _CopyInterface<IVsiImage>
		>
	> CVsiEnumImage;

	std::unique_ptr<IVsiImage *[]> ppImage;
	std::unique_ptr<CVsiEnumImage> pEnum;
	int i = 0;
	bool bLoadImages(false);

	try
	{
		VSI_CHECK_ARG(ppEnum, VSI_LOG_ERROR, L"ppEnum");

		*ppEnum = NULL;

		if (0 == m_mapImage.size())
		{
			hr = LoadSeriesData(NULL, VSI_DATA_TYPE_IMAGE_LIST, pbCancel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			bLoadImages = true;
		}

		if (pbCancel && *pbCancel)
		{
			if (bLoadImages)
			{
				m_mapImage.clear();
			}

			return S_FALSE;
		}

		ppImage.reset(new IVsiImage*[max(1, m_mapImage.size())]);
		VSI_CHECK_MEM(ppImage, VSI_LOG_ERROR, NULL);

		if (VSI_PROP_IMAGE_MIN != sortBy)
		{
			CComQIPtr<IVsiServiceProvider> pServiceProvider(g_pApp);
			VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

			// Gets Mode Manager
			CComPtr<IVsiModeManager> pModeMgr;
			hr = pServiceProvider->QueryService(
				VSI_SERVICE_MODE_MANAGER,
				__uuidof(IVsiModeManager),
				(IUnknown**)&pModeMgr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CVsiImage::CVsiSort sort;
			sort.m_sortBy = sortBy;
			sort.m_bDescending = bDescending;
			sort.m_pModeMgr = pModeMgr;

			CVsiListImage listImage;

			for (auto iterMap = m_mapImage.begin();	iterMap != m_mapImage.end(); ++iterMap)
			{
				IVsiImage *pImage = iterMap->second;

				CVsiImage image(pImage, &sort);
				listImage.push_back(image);
			}

			if (pbCancel && *pbCancel)
			{
				return S_FALSE;
			}

			listImage.sort();

			if (pbCancel && *pbCancel)
			{
				return S_FALSE;
			}

			for (auto iter = listImage.begin();	iter != listImage.end(); ++iter)
			{
				CVsiImage &image = *iter;

				CComPtr<IVsiImage> pImage(image.m_pImage);
				ppImage[i++] = pImage.Detach();
			}
		}
		else
		{
			for (auto iterMap = m_mapImage.begin();
				iterMap != m_mapImage.end();
				++iterMap)
			{
				CComPtr<IVsiImage> pImage(iterMap->second);
				ppImage[i++] = pImage.Detach();
			}
		}

		if (pbCancel && *pbCancel)
		{
			return S_FALSE;
		}

		pEnum.reset(new CVsiEnumImage);
		VSI_CHECK_MEM(pEnum, VSI_LOG_ERROR, NULL);

		hr = pEnum->Init(ppImage.get(), ppImage.get() + i, NULL, AtlFlagTakeOwnership);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		ppImage.release();  // pEnum is the owner

		hr = pEnum->QueryInterface(__uuidof(IVsiEnumImage), (void**)ppEnum);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		pEnum.release();  // ppEnum is the owner
	}
	VSI_CATCH(hr);

	if (NULL != ppImage)
	{
		for (; i > 0; i--)
		{
			if (NULL != ppImage[i])
			{
				ppImage[i]->Release();
			}
		}
	}

	if (bLoadImages)
	{
		m_mapImage.clear();
	}

	return hr;
}

/// <summary>
/// Get image using index
/// </summary>
/// <param name="iIndex">Index of the image, 0 is the image at the bottom of the study browser</param>
STDMETHODIMP CVsiSeries::GetImageFromIndex(
	int iIndex,
	VSI_PROP_IMAGE sortBy,
	BOOL bDescending,
	IVsiImage **ppImage)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		if (0 == m_mapImage.size())
		{
			if (0 < m_strFile.Length())
			{
				hr = LoadSeriesData(NULL, VSI_DATA_TYPE_IMAGE_LIST, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

		VSI_CHECK_ARG((iIndex >= 0) && (iIndex < (int)m_mapImage.size()),
			VSI_LOG_ERROR, VsiFormatMsg(L"iIndex = %d", iIndex));

		CComQIPtr<IVsiServiceProvider> pServiceProvider(g_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Gets Mode Manager
		CComPtr<IVsiModeManager> pModeMgr;
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_MODE_MANAGER,
			__uuidof(IVsiModeManager),
			(IUnknown**)&pModeMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CVsiImage::CVsiSort sort;
		sort.m_sortBy = sortBy;
		sort.m_bDescending = bDescending;
		sort.m_pModeMgr = pModeMgr;

		CVsiListImage listImage;

		for (auto iterMap = m_mapImage.begin();
			iterMap != m_mapImage.end();
			++iterMap)
		{
			IVsiImage *pImage = iterMap->second;

			CVsiImage image(pImage, &sort);
			listImage.push_back(image);
		}

		listImage.sort();

		int i = 0;
		CVsiListImage::reverse_iterator iter;
		for (iter = listImage.rbegin();
			iter != listImage.rend();
			iter++)
		{
			if (i++ == iIndex)
			{
				CVsiImage &image = *iter;

				CComPtr<IVsiImage> pImage(image.m_pImage);  // AddRef
				*ppImage = pImage.Detach();
				break;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}
STDMETHODIMP CVsiSeries::GetImage(LPCOLESTR pszId, IVsiImage **ppImage)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszId, VSI_LOG_ERROR, L"pszId");
		VSI_CHECK_ARG(ppImage, VSI_LOG_ERROR, L"ppImage");

		BOOL bLoadImages(FALSE);

		if (0 == m_mapImage.size() && 0 < m_strFile.Length())
		{
			hr = LoadSeriesData(NULL, VSI_DATA_TYPE_IMAGE_LIST, nullptr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			bLoadImages = TRUE;
		}

		auto iter = m_mapImage.find(pszId);
		if (iter != m_mapImage.end())
		{
			*ppImage = iter->second;
			(*ppImage)->AddRef();

			hr = S_OK;
		}

		if (bLoadImages)
		{
			m_mapImage.clear();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::GetItem(LPCOLESTR pszNs, IUnknown **ppUnk)
{
	HRESULT hr(S_FALSE);
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");
		VSI_CHECK_ARG(ppUnk, VSI_LOG_ERROR, L"ppUnk");

		*ppUnk = NULL;

		BOOL bLoadImages(FALSE);

		if (0 == m_mapImage.size() && 0 < m_strFile.Length())
		{
			hr = LoadSeriesData(NULL, VSI_DATA_TYPE_IMAGE_LIST, nullptr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			bLoadImages = TRUE;
		}

		CComPtr<IVsiImage> pImage;

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

			auto iter = m_mapImage.find(strId);
			if (iter != m_mapImage.end())
			{
				pImage = iter->second;

				*ppUnk = pImage.Detach();
				hr = S_OK;
			}
		}

		if (bLoadImages)
		{
			m_mapImage.clear();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::LoadSeriesData(LPCOLESTR pszFile, DWORD dwFlags, BOOL *pbCancel)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		if (NULL != pszFile)
		{
			m_strFile = pszFile;

			m_mapProp.clear();
			m_mapPropLabel.clear();

			m_dwErrorCode |= VSI_SERIES_ERROR_LOAD_XML;

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
				if (S_OK != GetProperty(VSI_PROP_SERIES_NAME, &vName))
				{
					CComVariant vError(CString(MAKEINTRESOURCE(IDS_SERIES_CORRUPTED_SERIES)));
					SetProperty(VSI_PROP_SERIES_NAME, &vError);
				}

				//set Id and NS in the case of a failure
				CComVariant vId;
				if (S_OK != GetProperty(VSI_PROP_SERIES_ID, &vId))
				{
					CComVariant vPath(m_strFile);
					hr = SetProperty(VSI_PROP_SERIES_ID, &vPath);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				if (m_pStudy != NULL)
				{
					UpdateNamespace();
				}

				WCHAR szMsg[300];
				err.FormatErrorMsg(szMsg, _countof(szMsg));
				VSI_LOG_MSG(VSI_LOG_ERROR | VSI_LOG_NO_BOX, VsiFormatMsg(L"Read file '%s' failed", pszFile));
				VSI_FAIL(VSI_LOG_ERROR | VSI_LOG_NO_BOX, VsiFormatMsg(L"SAXErrorHandler - %s", szMsg));
			}

			// Sets namespace
			{
				CComVariant vId;
				if (S_OK == GetProperty(VSI_PROP_SERIES_ID, &vId))
				{
					if (m_pStudy != NULL)
					{
						UpdateNamespace();
					}
				}
			}

			m_dwErrorCode &= ~VSI_SERIES_ERROR_LOAD_XML;

			// Checks version
			CComVariant vVersionRequired;
			HRESULT hrTest = GetProperty(VSI_PROP_SERIES_VERSION_REQUIRED, &vVersionRequired);
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
					m_dwErrorCode |= VSI_SERIES_ERROR_VERSION_INCOMPATIBLE;
				}
			}
		}

		if ((VSI_DATA_TYPE_IMAGE_LIST & dwFlags) && (VSI_SERIES_ERROR_NONE == m_dwErrorCode))
		{
			// Skip if already loaded
			if (0 == m_mapImage.size())
			{
				m_dwErrorCode |= VSI_SERIES_ERROR_LOAD_IMAGES;

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

				// Load images
				m_mapImage.clear();

				if (NULL == pszFile)
				{
					VSI_VERIFY(0 < m_strFile.Length(), VSI_LOG_ERROR, NULL);
					pszFile = m_strFile;
				}

				LPCWSTR pszSpt = wcsrchr(pszFile, L'\\');
				CString strPath(pszFile, (int)(pszSpt - pszFile));

				// Find images
				CString strFilter;
				strFilter.Format(L"%s\\*." VSI_FILE_EXT_IMAGE, (LPCWSTR)strPath);

				CString strImageFile;
				LPWSTR pszFileFilter;
				CVsiFileIterator fileIter(strFilter);
				while (NULL != (pszFileFilter = fileIter.Next()))
				{
					if (0 == _wcsicmp(CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME)), pszFileFilter))
						continue;
					if (0 == _wcsicmp(VSI_FILE_MEASUREMENT, pszFileFilter))
						continue;

					if (0 == (VSI_DATA_TYPE_GET_COUNT_ONLY & dwFlags))
					{
						strImageFile.Format(L"%s\\%s", (LPCWSTR)strPath, pszFileFilter);

						CComPtr<IVsiImage> pImage;
						hr = pImage.CoCreateInstance(__uuidof(VsiImage));
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						hr = pImage->SetSeries(static_cast<IVsiSeries*>(this));
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComVariant vId;
						CComQIPtr<IVsiPersistImage> pPersist(pImage);
						hr = pPersist->LoadImageData(strImageFile, NULL, NULL, dwFlags);
						if (FAILED(hr))
						{
							VSI_LOG_MSG(VSI_LOG_WARNING,
								VsiFormatMsg(L"Load image failed - '%s'",
								(LPCWSTR)strImageFile));
							hr = pImage->GetProperty(VSI_PROP_IMAGE_ID, &vId);
							if (S_OK != hr)
							{
								// Cannot even get the id, use file path as key
								vId = strImageFile;
							}
						}
						else
						{
							hr = pImage->GetProperty(VSI_PROP_IMAGE_ID, &vId);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if (S_OK != hr)
							{
								// Cannot even get the id, use file path as key
								vId = strImageFile;
							}
						}

						if (NULL != pSession)
						{
							CComVariant vNs;
							hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
							if (S_OK == hr)
							{
								int iSlot(-1);
								if (S_OK == pSession->GetIsImageLoaded(V_BSTR(&vNs), &iSlot))
								{
									pImage.Release();
									hr = pSession->GetImage(iSlot, &pImage);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								}
							}
						}

						m_mapImage.insert(
							CVsiMapImage::value_type(CString(V_BSTR(&vId)), pImage));
					}
					else
					{
						m_mapImage[pszFileFilter] = NULL;
					}

					if (pbCancel && (FALSE != *pbCancel))
					{
						m_mapImage.clear();
						break;
					}
				}

				m_dwErrorCode &= ~VSI_SERIES_ERROR_LOAD_IMAGES;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::SaveSeriesData(LPCOLESTR pszFile, DWORD dwFlags)
{
	UNREFERENCED_PARAMETER(dwFlags);

	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);
	HANDLE hTransaction = INVALID_HANDLE_VALUE;

	try
	{
		if (NULL != pszFile)
		{
			m_strFile = pszFile;
		}

		hTransaction = VsiTxF::VsiCreateTransaction(
			NULL,
			0,		
			TRANSACTION_DO_NOT_PROMOTE,	
			0,		
			0,		
			INFINITE,
			L"Series xml save transaction");
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

		// Create a temporary file using the SaveSeriesData, then use the transacted 
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

		hr = SaveSeriesDataInternal(szTempFileName);
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

HRESULT CVsiSeries::SaveSeriesDataInternal(LPCOLESTR pszFile)
{
	HRESULT hr = S_OK;

	try
	{
		// Add version modified
		CString strVersion;
		strVersion.Format(
			L"%u.%u.%u.%u",
			VSI_SOFTWARE_MAJOR, VSI_SOFTWARE_MIDDLE, VSI_SOFTWARE_MINOR, VSI_BUILD_NUMBER);		

		CComVariant vVersionModified(strVersion.GetString());
		hr = SetProperty(VSI_PROP_SERIES_VERSION_MODIFIED, &vVersionModified);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

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

		DWORD dwTagXmlIndex = 0;
		CVsiMapPropIter iter;
		for (iter = m_mapProp.begin(); iter != m_mapProp.end(); ++iter)
		{
			dwTagXmlIndex = iter->first - VSI_PROP_SERIES_MIN - 1;
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
			VSI_SERIES_XML_ELM_SERIES, VSI_STATIC_STRING_COUNT(VSI_SERIES_XML_ELM_SERIES), pa.p);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		pmxa->clear();

		// Saves property labels
		hr = SavePropertyLabel(pch);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pch->endElement(szNull, 0, szNull, 0,
			VSI_SERIES_XML_ELM_SERIES, VSI_STATIC_STRING_COUNT(VSI_SERIES_XML_ELM_SERIES));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pch->endDocument();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		out.Close();

	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::LoadPropertyLabel(LPCOLESTR pszFile)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszFile, VSI_LOG_ERROR, L"pszFile");

		m_mapPropLabel.clear();

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
			WCHAR szMsg[300];
			err.FormatErrorMsg(szMsg, _countof(szMsg));
			VSI_LOG_MSG(VSI_LOG_ERROR | VSI_LOG_NO_BOX, VsiFormatMsg(L"Read file '%s' failed", pszFile));
			VSI_FAIL(VSI_LOG_ERROR, szMsg);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::startElement(
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
		// series?
		bMatched = VsiXmlIsEqualElement(
			pwszLocalName,
			cchLocalName,
			VSI_SERIES_XML_ELM_SERIES,
			VSI_STATIC_STRING_COUNT(VSI_SERIES_XML_ELM_SERIES)
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
					g_pszXmlTag[i].pszName,
					lstrlen(g_pszXmlTag[i].pszName),
					&pszValue, &iValue);
				if (hr != S_OK)
				{
					// Checks tag is optional
					VSI_VERIFY(g_pszXmlTag[i].bOptional,
						VSI_LOG_ERROR | VSI_LOG_NO_BOX, L"Attribute missing");
					hr = S_OK;
					continue;
				}

				VsiReadProperty(
					g_pszXmlTag[i].dwId,
					g_pszXmlTag[i].type,
					pszValue, iValue,
					m_mapProp);
			}

			return hr;
		}

		// field
		bMatched = VsiXmlIsEqualElement(
			pwszLocalName,
			cchLocalName,
			VSI_SERIES_XML_ELM_FIELD,
			VSI_STATIC_STRING_COUNT(VSI_SERIES_XML_ELM_FIELD)
		);
		if (bMatched)
		{
			hr = pAttributes->getValueFromName(
				L"", 0,
				VSI_SERIES_XML_FIELD_ATT_NAME,
				VSI_STATIC_STRING_COUNT(VSI_SERIES_XML_FIELD_ATT_NAME),
				&pszValue, &iValue);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CString strName(pszValue, iValue);

			hr = pAttributes->getValueFromName(
				L"", 0,
				VSI_SERIES_XML_FIELD_ATT_LABEL,
				VSI_STATIC_STRING_COUNT(VSI_SERIES_XML_FIELD_ATT_LABEL),
				&pszValue, &iValue);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CString strLabel(pszValue, iValue);

			m_mapPropLabel.insert(CVsiMapPropLabel::value_type(strName, strLabel));

			return hr;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::InitNew()
{
	HRESULT hr = S_OK;

	try
	{
		m_strFile.Empty();
		m_dwErrorCode = VSI_SERIES_ERROR_NONE;
		m_mapProp.clear();
		m_mapImage.clear();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSeries::GetClassID(CLSID *pClassID)
{
	if (pClassID == NULL)
		return E_POINTER;

	*pClassID = GetObjectCLSID();

	return S_OK;
}

STDMETHODIMP CVsiSeries::IsDirty()
{
	return (m_dwModified > 0) ? S_OK : S_FALSE;
}

STDMETHODIMP CVsiSeries::Load(LPSTREAM pStm)
{
	HRESULT hr = S_OK;
	ULONG ulRead;

	try
	{
		// Images not loaded

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

STDMETHODIMP CVsiSeries::Save(LPSTREAM pStm, BOOL fClearDirty)
{
	UNREFERENCED_PARAMETER(fClearDirty);

	HRESULT hr = S_OK;
	ULONG ulWritten;

	try
	{
		// Images not saved

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

STDMETHODIMP CVsiSeries::GetSizeMax(ULARGE_INTEGER* pcbSize)
{
	UNREFERENCED_PARAMETER(pcbSize);

	return E_NOTIMPL;
}

void CVsiSeries::UpdateNamespace() throw(...)
{
	HRESULT hr;

	CComVariant vSeriesId;
	hr = GetProperty(VSI_PROP_SERIES_ID, &vSeriesId);
	if (S_OK == hr)
	{
		VSI_VERIFY(m_pStudy != NULL, VSI_LOG_ERROR, NULL);

		CComVariant vStudyId;
		hr = m_pStudy->GetProperty(VSI_PROP_STUDY_ID, &vStudyId);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		// Formats
		CString strNs;
		strNs.Format(L"%s/%s", V_BSTR(&vStudyId), V_BSTR(&vSeriesId));

		// Sets
		CComVariant vNs((LPCWSTR)strNs);
		hr = SetProperty(VSI_PROP_SERIES_NS, &vNs);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
}

HRESULT CVsiSeries::SavePropertyLabel(IUnknown *pUnkContent)
{
	HRESULT hr = S_OK;

	try
	{
		CComQIPtr<ISAXContentHandler> pch(pUnkContent);
		VSI_CHECK_INTERFACE(pch, VSI_LOG_ERROR, NULL);

		CComPtr<IMXAttributes> pmxa;
		hr = VsiCreateXmlSaxAttributes(&pmxa);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		WCHAR szNull[] = L"";
		CComBSTR bstrNull(L"");

		CComQIPtr<ISAXAttributes> pa(pmxa);

		// fields
		{
			hr = pch->startElement(szNull, 0, szNull, 0,
				VSI_SERIES_XML_ELM_FIELDS, VSI_STATIC_STRING_COUNT(VSI_SERIES_XML_ELM_FIELDS), pa.p);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant v;
			for (auto iter = m_mapPropLabel.begin();
				iter != m_mapPropLabel.end();
				++iter)
			{
				pmxa->clear();

				CComBSTR bstrNameValue(iter->first);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComBSTR bstrName(VSI_SERIES_XML_FIELD_ATT_NAME);
				hr = pmxa->addAttribute(bstrNull, bstrName, bstrName, bstrNull, bstrNameValue);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComBSTR bstrLabelValue(iter->second);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComBSTR bstrLabel(VSI_SERIES_XML_FIELD_ATT_LABEL);
				hr = pmxa->addAttribute(bstrNull, bstrLabel, bstrLabel, bstrNull, bstrLabelValue);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pch->startElement(szNull, 0, szNull, 0,
					VSI_SERIES_XML_ELM_FIELD, VSI_STATIC_STRING_COUNT(VSI_SERIES_XML_ELM_FIELD), pa.p);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pch->endElement(szNull, 0, szNull, 0,
					VSI_SERIES_XML_ELM_FIELD, VSI_STATIC_STRING_COUNT(VSI_SERIES_XML_ELM_FIELD));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			hr = pch->endElement(szNull, 0, szNull, 0,
				VSI_SERIES_XML_ELM_FIELDS, VSI_STATIC_STRING_COUNT(VSI_SERIES_XML_ELM_FIELDS));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return hr;
}
