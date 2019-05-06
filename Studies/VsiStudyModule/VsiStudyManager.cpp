/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudyManager.cpp
**
**	Description:
**		Implementation of CVsiStudyManager
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiSaxUtils.h>
#include <VsiCommUtl.h>
#include <VsiServiceProvider.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterUtils.h>
#include <VsiParameterSystem.h>
#include <VsiServiceKey.h>
#include <VsiOperatorModule.h>
#include <VsiStudyXml.h>
#include <VsiLicenseUtils.h>
#include "VsiStudyManager.h"

CVsiStudyManager::CVsiStudyManager()
{
	VSI_LOG_MSG(VSI_LOG_INFO, L"Study Manager Created");
}

CVsiStudyManager::~CVsiStudyManager()
{
	_ASSERT(0 == m_mapStudy.size());
	VSI_LOG_MSG(VSI_LOG_INFO, L"Study Manager Destroyed");
}

STDMETHODIMP CVsiStudyManager::Initialize(IVsiApp *pApp)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_LOG_MSG(VSI_LOG_INFO, L"Study Manager initialization started");

		m_pApp = pApp;
		VSI_CHECK_INTERFACE(m_pApp, VSI_LOG_ERROR,
			L"Failure while initializing Study Manager");

		// Caches app
		if (1 == InterlockedIncrement(&g_lRefApp))
			g_pApp = m_pApp;

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR,
			L"Failure getting service provider");

		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		VSI_LOG_MSG(VSI_LOG_INFO, L"Study Manager initialization completed");
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
			VSI_ERROR_LOG(err);

		Uninitialize();
	}

	return hr;
}

STDMETHODIMP CVsiStudyManager::Uninitialize()
{
	VSI_LOG_MSG(VSI_LOG_INFO, L"Study Manager un-initialization started");

	// Clear studies
	m_mapStudy.clear();

	m_pPdm.Release();

	if (0 == InterlockedDecrement(&g_lRefApp))
	{
		g_pApp = NULL;
	}

	// Release App
	m_pApp.Release();

	VSI_LOG_MSG(VSI_LOG_INFO, L"Study Manager un-initialization completed");

	return S_OK;
}

STDMETHODIMP CVsiStudyManager::LoadData(
	LPCOLESTR pszPath, DWORD dwTypes, const VSI_LOAD_STUDY_FILTER *pFilter)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		if (VSI_DATA_TYPE_STUDY_LIST & dwTypes)
		{
			CComPtr<IVsiStudy> pStudy;
			CString strStudyFile;
			CComPtr<IVsiOperatorManager> pOperatorMgr;
			std::set<CString> setOperatorInGroup;
			std::set<CString> setOperatorNotInGroup;

			// Get session
			CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
			VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

			CComPtr<IVsiSession> pSession;
			if (0 == (VSI_DATA_TYPE_NO_SESSION_LINK & dwTypes))
			{
				hr = pServiceProvider->QueryService(
					VSI_SERVICE_SESSION_MGR,
					__uuidof(IVsiSession),
					(IUnknown**)&pSession);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			// Clear old
			m_mapStudy.clear();

			// Set new
			if (NULL != pszPath)
			{
				m_strPath = pszPath;
			}

			VSI_VERIFY(!m_strPath.IsEmpty(), VSI_LOG_ERROR, NULL);

			if (NULL != pFilter)
			{
				m_pFilter.reset(new CVsiFilter(pFilter->pszOperator, pFilter->pszGroup, pFilter->bPublic));

				hr = pServiceProvider->QueryService(
					VSI_SERVICE_OPERATOR_MGR,
					__uuidof(IVsiOperatorManager),
					(IUnknown**)&pOperatorMgr);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else
			{
				m_pFilter.reset();
			}

			// Find studies
			CString strFilter;
			strFilter.Format(L"%s\\*", pszPath);

			LPWSTR pszDir;
			CVsiDirectoryIterator dirIter(strFilter);
			while (NULL != (pszDir = dirIter.Next()))
			{
				// Skip .tmp folder (they are leftover from bad file operations)
				if (lstrlen(pszDir) > lstrlen(VSI_FILE_EXT_TMP))
				{
					LPWSTR pszDot = wcsrchr(pszDir, L'.');
					if (NULL != pszDot)
					{
						if (0 == wcscmp(pszDot + 1, VSI_FILE_EXT_TMP))
						{
							continue;  // Skip
						}
					}
				}

				strStudyFile.Format(L"%s\\%s\\%s", pszPath, pszDir, CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));

				if (0 == _waccess(strStudyFile, 0))
				{
					hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vId;
					CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
					hr = pPersist->LoadStudyData(strStudyFile, dwTypes, NULL);
					if (FAILED(hr))
					{
						VSI_LOG_MSG(VSI_LOG_WARNING,
							VsiFormatMsg(L"Load study failed - '%s'",
							(LPCWSTR)strStudyFile));
						hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
						if (S_OK != hr)
						{
							// Cannot even get the id, use file path as key
							vId = strStudyFile;
						}
					}
					else
					{
						if (NULL != pSession)
						{
							CComVariant vStudyNs;
							hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vStudyNs);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if (S_OK == hr)
							{
								int iSlot(0);
								if (S_OK == pSession->GetIsStudyLoaded(V_BSTR(&vStudyNs), &iSlot))
								{
									pStudy.Release();

									hr = pSession->GetStudy(iSlot, &pStudy);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR,	VsiFormatMsg(L"Failure getting study from slot %d", iSlot));
								}
							}
						}

						hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (S_OK != hr)
						{
							vId = strStudyFile;
						}
					}

					// Filter
					bool bAdd(true);

					if (NULL != m_pFilter)
					{
						bAdd = IsStudyVisible(pStudy, pOperatorMgr, setOperatorInGroup, setOperatorNotInGroup);
					}

					if (bAdd)
					{
						// Check to see if this study is in the map already
						auto iter = m_mapStudy.find(V_BSTR(&vId));
						if (iter != m_mapStudy.end())
						{
							VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"Study ID is duplicated: %s", strStudyFile.GetString()));
						}

						m_mapStudy[V_BSTR(&vId)] = pStudy;
					}

					pStudy.Release();
				}
			}

			hr = S_OK;
		}
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
		{
			VSI_ERROR_LOG(err);
		}

		m_mapStudy.clear();
	}

	return hr;
}

STDMETHODIMP CVsiStudyManager::LoadStudy(
	LPCOLESTR pszPath,
	BOOL bFailAllow,
	IVsiStudy** ppStudy)
{
	HRESULT hr = S_OK;
	CComPtr<IVsiStudy> pStudy;
	CString strStudyFile;

	try
	{
		VSI_CHECK_ARG(ppStudy, VSI_LOG_ERROR, L"ppStudy");

		strStudyFile.Format(L"%s\\%s", pszPath, CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)));

		int iRet = _waccess(strStudyFile, 0);
		VSI_VERIFY(0 == iRet, VSI_LOG_ERROR, VsiFormatMsg(L"LoadStudy failed - %s", strStudyFile));

		hr = pStudy.CoCreateInstance(__uuidof(VsiStudy));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating study object");

		CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
		hr = pPersist->LoadStudyData(strStudyFile, 0, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, VsiFormatMsg(L"Failure loading study data for study %s", strStudyFile));

		*ppStudy = pStudy.Detach();
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr) && !bFailAllow)
		{
			VSI_ERROR_LOG(err);
		}
	}

	return hr;
}

STDMETHODIMP CVsiStudyManager::GetStudyCount(LONG *plCount)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(plCount, VSI_LOG_ERROR, L"plCount");

		*plCount = (LONG)m_mapStudy.size();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyManager::GetStudyEnumerator(
	BOOL bExcludeCorruptedAndNotCompatible,
	IVsiEnumStudy **ppEnum)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	typedef CComObject<
		CComEnum<
			IVsiEnumStudy,
			&__uuidof(IVsiEnumStudy),
			IVsiStudy*, _CopyInterface<IVsiStudy>
		>
	> CVsiEnumStudy;

	std::unique_ptr<IVsiStudy*[]> ppStudies;
	int i = 0;

	try
	{
		VSI_CHECK_ARG(ppEnum, VSI_LOG_ERROR, L"ppEnum");

		ppStudies.reset(new IVsiStudy*[m_mapStudy.size()]);

		if (bExcludeCorruptedAndNotCompatible)
		{
			for (auto iter = m_mapStudy.begin(); iter != m_mapStudy.end(); ++iter)
			{
				CComPtr<IVsiStudy> pStudy(iter->second);

				VSI_STUDY_ERROR error;
				hr = pStudy->GetErrorCode(&error);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (VSI_STUDY_ERROR_NONE == error)
				{
					ppStudies[i++] = pStudy.Detach();
				}
			}
		}
		else
		{
			for (auto iter = m_mapStudy.begin(); iter != m_mapStudy.end(); ++iter)
			{
				CComPtr<IVsiStudy> pStudy(iter->second);
				ppStudies[i++] = pStudy.Detach();
			}
		}

		std::unique_ptr<CVsiEnumStudy> pEnum(new CVsiEnumStudy);

		hr = pEnum->Init(ppStudies.get(), ppStudies.get() + i, NULL, AtlFlagTakeOwnership);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"GetStudyEnumerator failed");

		ppStudies.release();  // pEnum is the owner

		hr = pEnum->QueryInterface(__uuidof(IVsiEnumStudy), (void**)ppEnum);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"GetStudyEnumerator failed");

		pEnum.release();  // ppEnum is the owner
	}
	VSI_CATCH(hr);

	if (NULL != ppStudies)
	{
		while (--i >= 0)
		{
			if (NULL != ppStudies[i])
			{
				ppStudies[i]->Release();
			}
		}
	}

	return hr;
}

STDMETHODIMP CVsiStudyManager::GetStudy(LPCOLESTR pszNs, IVsiStudy **ppStudy)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");
		VSI_CHECK_ARG(ppStudy, VSI_LOG_ERROR, L"ppStudy");

		auto iter = m_mapStudy.find(pszNs);
		if (iter != m_mapStudy.end())
		{
			*ppStudy = iter->second;
			(*ppStudy)->AddRef();

			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyManager::GetItem(LPCOLESTR pszNs, IUnknown **ppUnk)
{
	HRESULT hr(S_FALSE);
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");
		VSI_CHECK_ARG(ppUnk, VSI_LOG_ERROR, L"ppUnk");

		CComPtr<IVsiStudy> pStudy;

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

			auto iter = m_mapStudy.find(strId);
			if (iter != m_mapStudy.end())
			{
				pStudy = iter->second;

				if (NULL != pszNext && 0 != *pszNext)
				{
					hr = pStudy->GetItem(pszNext, ppUnk);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
				else
				{
					*ppUnk = pStudy.Detach();
					hr = S_OK;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyManager::GetIsItemPresent(LPCOLESTR pszNs)
{
	HRESULT hr(S_FALSE);
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");

		CComPtr<IUnknown> pUnkItem;
		hr = GetItem(pszNs, &pUnkItem);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyManager::GetIsStudyPresent(
	VSI_PROP_STUDY dwIdType,
	LPCOLESTR pszId,
	LPCOLESTR pszOperatorName,
	IVsiStudy **ppStudy)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);
	CComVariant vId;
	CComVariant vOwner;

	try
	{
		VSI_CHECK_ARG(pszId, VSI_LOG_ERROR, L"pszId");
		VSI_CHECK_ARG(pszOperatorName, VSI_LOG_ERROR, L"pszOperatorName");

		BOOL bFound = FALSE;

		for (auto iter = m_mapStudy.begin(); iter != m_mapStudy.end(); ++iter)
		{
			IVsiStudy *pStudy = iter->second;

			vId.Clear();
			hr = pStudy->GetProperty(dwIdType, &vId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr)
			{
				vOwner.Clear();
				hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if ((S_OK == hr) &&
					(0 == _wcsicmp(V_BSTR(&vId), pszId)) &&
					(0 == _wcsicmp(V_BSTR(&vOwner), pszOperatorName)))
				{
					bFound = TRUE;

					if (NULL != ppStudy)
					{
						*ppStudy = pStudy;
						pStudy->AddRef();
					}

					break;
				}
			}
		}

		hr = bFound ? S_OK : S_FALSE;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyManager::GetStudyCopyNumber(LPCOLESTR pszCreatedId, LPCOLESTR pszOwner, int* piNum)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszCreatedId, VSI_LOG_ERROR, L"pszCreatedId");
		VSI_CHECK_ARG(pszOwner, VSI_LOG_ERROR, L"pszOwner");
		VSI_CHECK_ARG(piNum, VSI_LOG_ERROR, L"piNum");

		int iIdx = 0;
		int iNum = 0;

		const WCHAR szCopyOf[] = L"Copy of";

		for (auto iter = m_mapStudy.begin(); iter != m_mapStudy.end(); ++iter)
		{
			IVsiStudy *pStudy = iter->second;

			CComVariant vCreatedId;
			hr = pStudy->GetProperty(VSI_PROP_STUDY_ID_CREATED, &vCreatedId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vOwner;
			hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if ((S_OK == hr) &&
				(0 == _wcsicmp(V_BSTR(&vCreatedId), pszCreatedId) &&
				(0 == _wcsicmp(V_BSTR(&vOwner), pszOwner))))
			{
				iNum++;
				CComVariant vName;
				hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CString strName(V_BSTR(&vName));

				if (0 == wcsncmp((LPCWSTR)strName, szCopyOf, _countof(szCopyOf) - 1))
				{
					if (strName.GetAt(strName.GetLength() - 1) == L')')
					{
						int idx = strName.ReverseFind(L'(');
						if (idx != -1)
						{
							CString strNum = strName.Mid(idx + 1, strName.GetLength() - idx - 2);
							BOOL bAreDigits = TRUE;
							for (int iCnt = 0; iCnt < strNum.GetLength() - 1; iCnt++)
							{
								WCHAR chTmp = strNum.GetAt(iCnt);
								if (chTmp < L'0' || chTmp > L'9')
								{
									bAreDigits = FALSE;
									break;
								}
							}
							if (bAreDigits)//make sure this is a valid copy tag
							{
								int num = _wtoi(strNum);
								if (num > iIdx)
								{
									iIdx = num;
								}
							}
						}
					}
				}
			}
		}

		if (0 == iIdx && iNum > 0)
		{
			iIdx = 1;
		}
		else if (0 < iIdx)
		{
			++iIdx;  // Next id
		}

		*piNum = iIdx;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyManager::AddStudy(IVsiStudy *pStudy)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pStudy, VSI_LOG_ERROR, L"pStudy");

		CComVariant vNs;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vNs);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"AddOperator failed");

		m_mapStudy[V_BSTR(&vNs)] = pStudy;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyManager::RemoveStudy(LPCOLESTR pszNs)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");

		auto iter = m_mapStudy.find(pszNs);
		if (iter != m_mapStudy.end())
		{
			m_mapStudy.erase(iter);

			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyManager::MoveSeries(IVsiSeries *pSrcSeries, IVsiStudy *pDesStudy, LPOLESTR *ppszResult)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pSrcSeries, VSI_LOG_ERROR, L"pSrcSeries");
		VSI_CHECK_ARG(pDesStudy, VSI_LOG_ERROR, L"pDesStudy");
		VSI_CHECK_ARG(ppszResult, VSI_LOG_ERROR, L"ppszResult");

		CComVariant vSeriesNs;
		hr = pSrcSeries->GetProperty(VSI_PROP_SERIES_NS, &vSeriesNs);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		// Gets series Id
		CComVariant vSeriesId;
		hr = pSrcSeries->GetProperty(VSI_PROP_SERIES_ID, &vSeriesId);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		CComHeapPtr<OLECHAR> pszSrcDataPath;
		CComHeapPtr<OLECHAR> pszDesDataPath;

		// Source
		hr = pSrcSeries->GetDataPath(&pszSrcDataPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

		PathRemoveFileSpec(pszSrcDataPath);

		LPWSTR pszFolder = wcsrchr(pszSrcDataPath, L'\\');
		VSI_VERIFY(NULL != pszFolder, VSI_LOG_ERROR,	NULL);
		CString strSeriesFolder(pszFolder + 1);

		// Des
		hr = pDesStudy->GetDataPath(&pszDesDataPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

		PathRemoveFileSpec(pszDesDataPath);

		CString strDesDataPath;
		strDesDataPath.Format(L"%s\\%s", pszDesDataPath, LPCWSTR(strSeriesFolder));

		BOOL bMove = MoveFile(pszSrcDataPath, strDesDataPath);
		if (bMove)
		{
			CComPtr<IVsiStudy> pSrcStudy;
			hr = pSrcSeries->GetStudy(&pSrcStudy);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

			// Remove Series from source Study
			hr = pSrcStudy->RemoveSeries(V_BSTR(&vSeriesId));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

			// Update the Data path
			strDesDataPath += L"\\";
			strDesDataPath += CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME));
			hr = pSrcSeries->SetDataPath(strDesDataPath);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

			// Add Series to destination Study
			hr = pDesStudy->AddSeries(pSrcSeries);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

			// Construct update xml file
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

				// Remove series from study
				{
					CComPtr<IXMLDOMElement> pElmRemove;
					hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_REMOVE), &pElmRemove);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure creating '%s' element", VSI_UPDATE_XML_ELM_REMOVE));

					CComVariant vNum(1);
					hr = pElmRemove->setAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_SERIES), vNum);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failed to set attribute (%s)", VSI_REMOVE_XML_ATT_NUMBER_SERIES));

					hr = pElmUpdate->appendChild(pElmRemove, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IXMLDOMElement> pElmItem;
					hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_ITEM), &pElmItem);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure creating '%s' element", VSI_UPDATE_XML_ELM_ITEM));

					hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), vSeriesNs);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failed to set attribute (%s)", VSI_UPDATE_XML_ATT_NS));

					hr = pElmRemove->appendChild(pElmItem, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				// Add series to the new study
				{
					CComPtr<IXMLDOMElement> pElmAdd;
					hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_ADD), &pElmAdd);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure creating '%s' element", VSI_UPDATE_XML_ELM_ADD));

					hr = pElmUpdate->appendChild(pElmAdd, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IXMLDOMElement> pElmItem;
					hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_ITEM), &pElmItem);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure creating '%s' element", VSI_UPDATE_XML_ELM_ITEM));

					CComVariant vNs;

					hr = pSrcSeries->GetProperty(VSI_PROP_SERIES_NS, &vNs);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), vNs);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failed to set attribute (%s)", VSI_UPDATE_XML_ATT_NS));

					CComHeapPtr<OLECHAR> pszPath;
					hr = pSrcSeries->GetDataPath(&pszPath);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vPath(pszPath);
					hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_PATH), vPath);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failed to set attribute (%s)", VSI_UPDATE_XML_ATT_PATH));

					CComVariant vSelected(true);
					hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_SELECT), vSelected);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failed to set attribute (%s)", VSI_UPDATE_XML_SELECT));

					hr = pElmAdd->appendChild(pElmItem, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

#ifdef _DEBUG
{
				WCHAR szPath[MAX_PATH];
				GetModuleFileName(NULL, szPath, _countof(szPath));
				LPWSTR pszSpt = wcsrchr(szPath, L'\\');
				lstrcpy(pszSpt + 1, VSI_FOLDER_DEBUG);
				lstrcat(szPath, L"\\move.xml");
				pUpdateXmlDoc->save(CComVariant(szPath));
}
#endif

				// Fire update event
				{
					CComVariant vUpdateXmlDoc((IUnknown*)pUpdateXmlDoc);
					NotifyEvent(VSI_STUDY_MGR_EVENT_UPDATE, &vUpdateXmlDoc);
				}
			}
		}
		else
		{
			CString strMsg;
			DWORD dwError = GetLastError();
			if (ERROR_ALREADY_EXISTS == dwError)
			{
				strMsg = CString(MAKEINTRESOURCE(IDS_MOVESER_SERIES_EXISTS));
			}
			else
			{
				strMsg = VsiGetErrorDescription(HRESULT_FROM_WIN32(dwError));
			}
			VSI_LOG_MSG(VSI_LOG_WARNING, VsiFormatMsg(L"MoveFile failed - %s", strMsg));

			*ppszResult = AtlAllocTaskWideString(strMsg);
			VSI_CHECK_MEM(*ppszResult, VSI_LOG_ERROR, NULL);

			hr = S_FALSE;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyManager::Update(IXMLDOMDocument *pXmlDoc)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pXmlDoc, VSI_LOG_ERROR, NULL);

		// Get session
		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiSession> pSession;
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_SESSION_MGR,
			__uuidof(IVsiSession),
			(IUnknown**)&pSession);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Remove
		{
			CComPtr<IXMLDOMNodeList> pListItemRemove;
			hr = pXmlDoc->selectNodes(
				CComBSTR(L"//" VSI_UPDATE_XML_ELM_UPDATE L"/"
					VSI_UPDATE_XML_ELM_REMOVE L"/" VSI_UPDATE_XML_ELM_ITEM),
				&pListItemRemove);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			long lCount = 0;
			hr = pListItemRemove->get_length(&lCount);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

			CComVariant vNs;
			for (long l = 0; l < lCount; ++l)
			{
				CComPtr<IXMLDOMNode> pNode;
				hr = pListItemRemove->get_item(l, &pNode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

				CComQIPtr<IXMLDOMElement> pElemItem(pNode);

				vNs.Clear();
				hr = pElemItem->getAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), &vNs);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_ATT_NS));

				LPCWSTR pszNs = V_BSTR(&vNs);

				CComPtr<IUnknown> pUnkItem;
				hr = GetItem(pszNs, &pUnkItem);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

				if (S_OK == hr)
				{
					CComQIPtr<IVsiStudy> pStudy(pUnkItem);
					if (NULL != pStudy)
					{
						hr = RemoveStudy(pszNs);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
					}
					else
					{
						CComQIPtr<IVsiSeries> pSeries(pUnkItem);
						if (NULL != pSeries)
						{
							CComPtr<IVsiStudy> pStudyParent;

							hr = pSeries->GetStudy(&pStudyParent);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

							LPCWSTR pszNsSeries = wcschr(pszNs, L'/') + 1;

							hr = pStudyParent->RemoveSeries(pszNsSeries);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
						}
						else
						{
							CComQIPtr<IVsiImage> pImage(pUnkItem);
							if (NULL != pImage)
							{
								CComPtr<IVsiSeries> pSeriesParent;

								hr = pImage->GetSeries(&pSeriesParent);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

								LPCWSTR pszNsSeries = wcschr(pszNs, L'/') + 1;
								LPCWSTR pszNsImage = wcschr(pszNsSeries, L'/') + 1;

								hr = pSeriesParent->RemoveImage(pszNsImage);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
							}
							else
							{
								_ASSERT(0);
							}
						}
					}
				}
			}
		}

		// Add study
		{
			// Filtering support
			CComPtr<IVsiOperatorManager> pOperatorMgr;
			std::set<CString> setOperatorInGroup;
			std::set<CString> setOperatorNotInGroup;

			if (NULL != m_pFilter)
			{
				// Get operator manager
				CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
				VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

				hr = pServiceProvider->QueryService(
					VSI_SERVICE_OPERATOR_MGR,
					__uuidof(IVsiOperatorManager),
					(IUnknown**)&pOperatorMgr);
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
			for (long l = 0; l < lCount; ++l)
			{
				CComPtr<IXMLDOMNode> pNode;
				hr = pListItemAdd->get_item(l, &pNode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

				CComQIPtr<IXMLDOMElement> pElemItem(pNode);

				vNs.Clear();
				hr = pElemItem->getAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), &vNs);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_ATT_NS));

				vPath.Clear();
				hr = pElemItem->getAttribute(CComBSTR(VSI_UPDATE_XML_ATT_PATH), &vPath);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_ATT_PATH));

				LPCWSTR pszNs = V_BSTR(&vNs);
				LPCWSTR pszNsTest = pszNs;
				LPCWSTR pszNsNext = wcschr(pszNsTest, L'/');
				if (NULL == pszNsNext)
				{
					// Is a study
					if (S_OK != GetIsItemPresent(pszNs))
					{
						CComPtr<IVsiStudy> pStudy;

						int iSlot(0);
						if (S_OK == pSession->GetIsStudyLoaded(pszNs, &iSlot))
						{
							hr = pSession->GetStudy(iSlot, &pStudy);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,	VsiFormatMsg(L"Failure getting study from slot %d", iSlot));
						}
						else
						{
							hr = LoadStudy(V_BSTR(&vPath), FALSE, &pStudy);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,	VsiFormatMsg(L"Failure loading study %s", V_BSTR(&vPath)));
						}

						bool bAddStudy(true);

						if (NULL != m_pFilter)
						{
							bAddStudy = IsStudyVisible(pStudy, pOperatorMgr, setOperatorInGroup, setOperatorNotInGroup);
						}

						if (bAddStudy)
						{
							hr = AddStudy(pStudy);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure adding study");
						}
					}
				}
			}
		}

		// Refresh - only required when there is filtering (e.g. changing study owner or share setting)
		if (NULL != m_pFilter)
		{
			CComPtr<IVsiOperatorManager> pOperatorMgr;
			std::set<CString> setOperatorInGroup;
			std::set<CString> setOperatorNotInGroup;
			CComQIPtr<IXMLDOMNode> pNodeRemove;
			int iNumStudy(0);

			// Add Remove element if it is not there
			{
				hr = pXmlDoc->selectSingleNode(
					CComBSTR(L"//" VSI_UPDATE_XML_ELM_UPDATE L"/" VSI_UPDATE_XML_ELM_REMOVE),
					&pNodeRemove);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK != hr)
				{
					CComPtr<IXMLDOMNode> pNodeUpdate;
					hr = pXmlDoc->selectSingleNode(CComBSTR(L"//" VSI_UPDATE_XML_ELM_UPDATE), &pNodeUpdate);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IXMLDOMElement> pElmRemove;
					hr = pXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_REMOVE), &pElmRemove);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pNodeUpdate->appendChild(pElmRemove, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					pNodeRemove = pElmRemove;
					VSI_CHECK_INTERFACE(pNodeRemove, VSI_LOG_ERROR, NULL);
				}
				else
				{
					CComVariant vNum;
					CComQIPtr<IXMLDOMElement> pElmRemove(pNodeRemove);
					hr = pElmRemove->getAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_STUDIES), &vNum);
					if (S_OK == hr)
					{
						VsiVariantDecodeValue(vNum, VT_I4);
						iNumStudy = V_I4(&vNum);
					}
				}
			}

			// Get operator manager
			CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
			VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

			hr = pServiceProvider->QueryService(
				VSI_SERVICE_OPERATOR_MGR,
				__uuidof(IVsiOperatorManager),
				(IUnknown**)&pOperatorMgr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Refresh
			CComPtr<IXMLDOMNodeList> pListItemRefresh;
			hr = pXmlDoc->selectNodes(
				CComBSTR(L"//" VSI_UPDATE_XML_ELM_UPDATE L"/"
					VSI_UPDATE_XML_ELM_REFRESH L"/" VSI_UPDATE_XML_ELM_ITEM),
				&pListItemRefresh);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			long lCount = 0;
			hr = pListItemRefresh->get_length(&lCount);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

			CComVariant vNs;
			for (long l = 0; l < lCount; ++l)
			{
				CComPtr<IXMLDOMNode> pNode;
				hr = pListItemRefresh->get_item(l, &pNode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

				CComQIPtr<IXMLDOMElement> pElemItem(pNode);

				vNs.Clear();
				hr = pElemItem->getAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), &vNs);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_UPDATE_XML_ATT_NS));

				LPCWSTR pszNs = V_BSTR(&vNs);

				CComPtr<IUnknown> pUnkItem;
				hr = GetItem(pszNs, &pUnkItem);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

				if (S_OK == hr)
				{
					CComQIPtr<IVsiStudy> pStudy(pUnkItem);
					if (NULL != pStudy)
					{
						if (! IsStudyVisible(pStudy, pOperatorMgr, setOperatorInGroup, setOperatorNotInGroup))
						{
							hr = RemoveStudy(pszNs);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

							CComVariant vNs;
							hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vNs);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

							CComPtr<IXMLDOMElement> pElmItem;
							hr = pXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_ITEM), &pElmItem);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), vNs);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							hr = pNodeRemove->appendChild(pElmItem, NULL);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							++iNumStudy;
						}
					}
				}
			}

			if (0 < iNumStudy)
			{
				CComVariant vNum(iNumStudy);

				CComQIPtr<IXMLDOMElement> pElmRemove(pNodeRemove);
				VSI_CHECK_INTERFACE(pElmRemove, VSI_LOG_ERROR, NULL);

				hr = pElmRemove->setAttribute(CComBSTR(VSI_REMOVE_XML_ATT_NUMBER_STUDIES), vNum);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

		// Fire update event
		{
			CComVariant vUpdateXmlDoc((IUnknown*)pXmlDoc);
			NotifyEvent(VSI_STUDY_MGR_EVENT_UPDATE, &vUpdateXmlDoc);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

bool CVsiStudyManager::IsStudyVisible(
	IVsiStudy *pStudy,
	IVsiOperatorManager *pOperatorMgr,
	std::set<CString> &setOperatorInGroup,
	std::set<CString> &setOperatorNotInGroup)
{
	HRESULT hr(S_OK);
	bool bVisible(true);

	try
	{
		CComVariant vOwner;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK != hr)
		{
			bVisible = false;
		}
		else if (m_pFilter->m_strOperator != V_BSTR(&vOwner))
		{
			// Not the same owner
			if (m_pFilter->m_bPublic)
			{
				CComVariant vAccess;
				hr = pStudy->GetProperty(VSI_PROP_STUDY_ACCESS, &vAccess);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr && 0 == wcscmp(VSI_STUDY_ACCESS_PRIVATE, V_BSTR(&vAccess)))
				{
					bVisible = false;
				}
				else if (S_OK == hr && 0 == wcscmp(VSI_STUDY_ACCESS_GROUP, V_BSTR(&vAccess)))
				{
					// Same group?

					// Check cache
					auto iter = setOperatorInGroup.find(V_BSTR(&vOwner));
					if (iter == setOperatorInGroup.end())
					{
						auto iterNot = setOperatorNotInGroup.find(V_BSTR(&vOwner));
						if (iterNot == setOperatorNotInGroup.end())
						{
							// Not in cache, check owner group
							CComPtr<IVsiOperator> pOperator;
							hr = pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, V_BSTR(&vOwner), &pOperator);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if (S_OK == hr)
							{
								CComHeapPtr<OLECHAR> pszGroup;
								hr = pOperator->GetGroup(&pszGroup);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								if ((0 != *pszGroup) && (m_pFilter->m_strGroup == pszGroup))
								{
									// Same group
									setOperatorInGroup.insert(V_BSTR(&vOwner));
								}
								else
								{
									// Not same group
									setOperatorNotInGroup.insert(V_BSTR(&vOwner));
									bVisible = false;
								}
							}
							else
							{
								// Owner not found in local system
								setOperatorNotInGroup.insert(V_BSTR(&vOwner));
								bVisible = false;
							}
						}
						else
						{
							// Not same group
							bVisible = false;
						}
					}
				}
			}
			else if (! m_pFilter->m_strGroup.IsEmpty())
			{
				CComVariant vAccess;
				hr = pStudy->GetProperty(VSI_PROP_STUDY_ACCESS, &vAccess);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Means public if property is missing
				if (S_OK != hr || 0 != wcscmp(VSI_STUDY_ACCESS_PRIVATE, V_BSTR(&vAccess)))
				{
					// Same group?

					// Check cache
					auto iter = setOperatorInGroup.find(V_BSTR(&vOwner));
					if (iter == setOperatorInGroup.end())
					{
						auto iterNot = setOperatorNotInGroup.find(V_BSTR(&vOwner));
						if (iterNot == setOperatorNotInGroup.end())
						{
							// Not in cache, check owner group
							CComPtr<IVsiOperator> pOperator;
							hr = pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, V_BSTR(&vOwner), &pOperator);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							if (S_OK == hr)
							{
								CComHeapPtr<OLECHAR> pszGroup;
								hr = pOperator->GetGroup(&pszGroup);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								if ((0 != *pszGroup) && (m_pFilter->m_strGroup == pszGroup))
								{
									// Same group
									setOperatorInGroup.insert(V_BSTR(&vOwner));
								}
								else
								{
									// Not same group
									setOperatorNotInGroup.insert(V_BSTR(&vOwner));
									bVisible = false;
								}
							}
							else
							{
								// Owner not found in local system
								setOperatorNotInGroup.insert(V_BSTR(&vOwner));
								bVisible = false;
							}
						}
						else
						{
							// Not same group
							bVisible = false;
						}
					}
				}
				else
				{
					bVisible = false;
				}
			}
			else
			{
				bVisible = false;
			}
		}
	}
	VSI_CATCH(hr);

	return bVisible;
}
