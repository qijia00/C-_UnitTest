/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiSession.cpp
**
**	Description:
**		Implementation of CVsiSession
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiSaxUtils.h>
#include <VsiBuildNumber.h>
#include <VsiServiceProvider.h>
#include <VsiCommUtlLib.h>
#include <VsiCommUtl.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiServiceKey.h>
#include <VsiResProduct.h>
#include <VsiOperatorModule.h>
#include "VsiSession.h"


CVsiSession::CVsiSession() :
	m_iSlotActive(0)
{
	VSI_LOG_MSG(VSI_LOG_INFO, L"Session Manager Created");
}

CVsiSession::~CVsiSession()
{
	_ASSERT(m_pStudy == NULL);
	_ASSERT(m_pSeries == NULL);

	for (int i = 0; i < _countof(m_ppStudy); ++i)
	{
		_ASSERT(m_ppStudy[i] == NULL);
		_ASSERT(m_ppSeries[i] == NULL);
		_ASSERT(m_ppImage[i] == NULL);
		_ASSERT(m_ppMode[i] == NULL);
	}

	VSI_LOG_MSG(VSI_LOG_INFO, L"Session Manager Destroyed");
}

STDMETHODIMP CVsiSession::Initialize(IVsiApp *pApp)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		m_pApp = pApp;
		VSI_CHECK_INTERFACE(m_pApp, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

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

STDMETHODIMP CVsiSession::Uninitialize()
{
	CCritSecLock csl(m_cs);

	for (int i = 0; i < _countof(m_ppMode); ++i)
	{
		m_ppMode[i].Release();
		m_ppImage[i].Release();
		m_ppSeries[i].Release();
		m_ppStudy[i].Release();
	}

	m_pSeries.Release();
	m_pStudy.Release();

	m_pOperatorMgr.Release();
	m_pPdm.Release();
	m_pApp.Release();

	return S_OK;
}

STDMETHODIMP CVsiSession::GetSessionMode(VSI_SESSION_MODE *pMode)
{
	HRESULT hr(S_OK);

	try
	{
		VSI_CHECK_ARG(pMode, VSI_LOG_ERROR, L"pMode");
	
		*pMode = m_mode;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::SetSessionMode(VSI_SESSION_MODE mode)
{
	m_mode = mode;

	return S_OK;
}

STDMETHODIMP CVsiSession::GetPrimaryStudy(IVsiStudy **ppStudy)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_csStudy);

	try
	{
		VSI_CHECK_ARG(ppStudy, VSI_LOG_ERROR, L"ppStudy");

		if (m_pStudy != NULL)
		{
			*ppStudy = m_pStudy;
			(*ppStudy)->AddRef();
			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::SetPrimaryStudy(IVsiStudy *pStudy, BOOL bJournal)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_csStudy);

	try
	{
		m_pStudy = pStudy;

		if (bJournal && (pStudy != NULL))
		{
			CComVariant vId;
			hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszFile;
			hr = pStudy->GetDataPath(&pszFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vPath(pszFile);
			Record(VSI_SESSION_EVENT_STUDY_OPEN, &vId, &vPath);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::GetPrimarySeries(IVsiSeries **ppSeries)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_csSeries);

	try
	{
		VSI_CHECK_ARG(ppSeries, VSI_LOG_ERROR, L"ppSeries");

		if (m_pSeries != NULL)
		{
			*ppSeries = m_pSeries;
			(*ppSeries)->AddRef();
			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::SetPrimarySeries(IVsiSeries *pSeries, BOOL bJournal)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_csSeries);

	try
	{
		if (m_pSeries != NULL && pSeries != NULL)
		{
			// Checks whether the series are same
			if (IsSeriesIdentical(m_pSeries, pSeries))
				return S_FALSE;  // No need to switch
		}

		m_pSeries = pSeries;

		if (bJournal && (pSeries != NULL))
		{
			CComVariant vId;
			hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszFile;
			hr = pSeries->GetDataPath(&pszFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vPath(pszFile);
			Record(VSI_SESSION_EVENT_SERIES_OPEN, &vId, &vPath);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::GetActiveSlot(int *piSlot)
{
	*piSlot = m_iSlotActive;

	return S_OK;
}

STDMETHODIMP CVsiSession::SetActiveSlot(int iSlot)
{
	m_iSlotActive = iSlot;

	return S_OK;
}

STDMETHODIMP CVsiSession::GetStudy(int iSlot, IVsiStudy **ppStudy)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_csStudy);

	try
	{
		VSI_CHECK_ARG(ppStudy, VSI_LOG_ERROR, L"ppStudy");

		if (iSlot == VSI_SLOT_ACTIVE)
			iSlot = m_iSlotActive;

		if (m_ppStudy[iSlot] != NULL)
		{
			*ppStudy = m_ppStudy[iSlot];
			(*ppStudy)->AddRef();
			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::SetStudy(int iSlot, IVsiStudy *pStudy, BOOL bJournal)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_csStudy);

	try
	{
		if (iSlot == VSI_SLOT_ACTIVE)
			iSlot = m_iSlotActive;

		if (m_ppStudy[iSlot] != NULL && pStudy != NULL)
		{
			// Checks whether the studies are the same
			if (IsStudyIdentical(m_ppStudy[iSlot], pStudy))
				return S_FALSE;
		}

		m_ppStudy[iSlot] = pStudy;

		if (bJournal && (pStudy != NULL))
		{
			CComVariant vId;
			hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszFile;
			hr = pStudy->GetDataPath(&pszFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vPath(pszFile);
			Record(VSI_SESSION_EVENT_STUDY_OPEN, &vId, &vPath);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::GetSeries(int iSlot, IVsiSeries **ppSeries)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_csSeries);

	try
	{
		VSI_CHECK_ARG(ppSeries, VSI_LOG_ERROR, L"ppSeries");

		if (iSlot == VSI_SLOT_ACTIVE)
			iSlot = m_iSlotActive;

		if (m_ppSeries[iSlot] != NULL)
		{
			*ppSeries = m_ppSeries[iSlot];
			(*ppSeries)->AddRef();
			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::SetSeries(int iSlot, IVsiSeries *pSeries, BOOL bJournal)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_csSeries);

	try
	{
		if (iSlot == VSI_SLOT_ACTIVE)
			iSlot = m_iSlotActive;

		if (m_ppSeries[iSlot] != NULL && pSeries != NULL)
		{
			// Checks whether the series are same
			if (IsSeriesIdentical(m_ppSeries[iSlot], pSeries))
				return S_FALSE;  // No need to switch
		}

		m_ppSeries[iSlot] = pSeries;

		if (bJournal && (pSeries != NULL))
		{
			CComVariant vId;
			hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszFile;
			hr = pSeries->GetDataPath(&pszFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vPath(pszFile);
			Record(VSI_SESSION_EVENT_SERIES_OPEN, &vId, &vPath);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::GetImage(int iSlot, IVsiImage **ppImage)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_csImage);

	try
	{
		VSI_CHECK_ARG(ppImage, VSI_LOG_ERROR, L"ppImage");

		if (iSlot == VSI_SLOT_ACTIVE)
			iSlot = m_iSlotActive;

		if (m_ppImage[iSlot] != NULL)
		{
			*ppImage = m_ppImage[iSlot];
			(*ppImage)->AddRef();
			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::SetImage(
	int iSlot, IVsiImage *pImage, BOOL bSetReviewed, BOOL bJournal)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_csImage);

	try
	{
		if (iSlot == VSI_SLOT_ACTIVE)
			iSlot = m_iSlotActive;

		if (m_ppImage[iSlot] != NULL && pImage != NULL)
		{
			// Checks whether the images are same
			if (IsImageIdentical(m_ppImage[iSlot], pImage))
				return S_FALSE;  // No need to switch
		}

		m_ppImage[iSlot] = pImage;
		
		if (NULL != pImage)
		{
			if (bSetReviewed)
			{
				CComVariant vNsImage;
				hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNsImage);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);
			
				CString strNsImage(V_BSTR(&vNsImage));
				m_setReviewed.insert(strNsImage);
			}

			if (bJournal)
			{
				CComVariant vId;
				hr = pImage->GetProperty(VSI_PROP_IMAGE_ID, &vId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComHeapPtr<OLECHAR> pszFile;
				hr = pImage->GetDataPath(&pszFile);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vPath(pszFile);
				Record(VSI_SESSION_EVENT_IMAGE_OPEN, &vId, &vPath);
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::GetMode(int iSlot, IUnknown **ppUnkMode)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_csMode);

	try
	{
		VSI_CHECK_ARG(ppUnkMode, VSI_LOG_ERROR, L"ppUnkMode");

		if (iSlot == VSI_SLOT_ACTIVE)
			iSlot = m_iSlotActive;

		if (m_ppMode[iSlot] != NULL)
		{
			*ppUnkMode = m_ppMode[iSlot];
			(*ppUnkMode)->AddRef();
			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::SetMode(int iSlot, IUnknown *pUnkMode)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_csMode);

	try
	{
		if (iSlot == VSI_SLOT_ACTIVE)
			iSlot = m_iSlotActive;

		m_ppMode[iSlot].Release();

		if (pUnkMode != NULL)
		{
			hr = pUnkMode->QueryInterface(
				__uuidof(IVsiMode),
				(void**)&m_ppMode[iSlot]);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComQIPtr<IVsiPropertyBag> pProperties(m_ppMode[iSlot]);
			VSI_CHECK_INTERFACE(pProperties, VSI_LOG_ERROR, NULL);

			CComVariant vSlot(iSlot);
			hr = pProperties->WriteId(VSI_PROP_MODE_SLOT_ID, &vSlot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::GetModeView(int iSlot, IUnknown **ppUnkModeView)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_csModeView);

	try
	{
		VSI_CHECK_ARG(ppUnkModeView, VSI_LOG_ERROR, L"ppUnkModeView");

		if (iSlot == VSI_SLOT_ACTIVE)
			iSlot = m_iSlotActive;

		if (m_ppModeView[iSlot] != NULL)
		{
			*ppUnkModeView = m_ppModeView[iSlot];
			(*ppUnkModeView)->AddRef();
			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::SetModeView(int iSlot, IUnknown *pUnkModeView)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_csModeView);

	try
	{
		if (iSlot == VSI_SLOT_ACTIVE)
			iSlot = m_iSlotActive;

		m_ppModeView[iSlot].Release();

		if (pUnkModeView != NULL)
		{
			hr = pUnkModeView->QueryInterface(
				__uuidof(IVsiModeView),
				(void**)&m_ppModeView[iSlot]);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::GetIsStudyLoaded(LPCOLESTR pszNs, int *piSlot)
{
	HRESULT hrFound(S_FALSE);
	HRESULT hr(S_OK);
	CCritSecLock csl(m_csStudy);

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");
		if (NULL != piSlot)
		{
			*piSlot = -1;
		}

		for (int i = 0; i < _countof(m_ppStudy); ++i)
		{
			if (m_ppStudy[i] != NULL)
			{
				CComVariant vNs;
				hr = m_ppStudy[i]->GetProperty(VSI_PROP_STUDY_NS, &vNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (0 == lstrcmp(pszNs, V_BSTR(&vNs)))
				{
					if (NULL != piSlot)
					{
						*piSlot = i;
					}
					hrFound = S_OK;
					break;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return SUCCEEDED(hr) ? hrFound : hr;
}

STDMETHODIMP CVsiSession::GetIsSeriesLoaded(LPCOLESTR pszNs, int *piSlot)
{
	HRESULT hrFound(S_FALSE);
	HRESULT hr(S_OK);
	CCritSecLock csl(m_csSeries);

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");

		if (NULL != piSlot)
		{
			*piSlot = -1;
		}

		for (int i = 0; i < _countof(m_ppSeries); ++i)
		{
			if (m_ppSeries[i] != NULL)
			{
				CComVariant vNsSeries;
				hr = m_ppSeries[i]->GetProperty(VSI_PROP_SERIES_NS, &vNsSeries);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (0 == lstrcmp(pszNs, V_BSTR(&vNsSeries)))
				{
					if (NULL != piSlot)
					{
						*piSlot = i;
					}
					hrFound = S_OK;
					break;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return SUCCEEDED(hr) ? hrFound : hr;
}

STDMETHODIMP CVsiSession::GetIsImageLoaded(LPCOLESTR pszNs, int *piSlot)
{
	HRESULT hrFound(S_FALSE);
	HRESULT hr(S_FALSE);
	CCritSecLock csl(m_csImage);

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");

		if (NULL != piSlot)
		{
			*piSlot = -1;
		}

		for (int i = 0; i < _countof(m_ppImage); ++i)
		{
			if (m_ppImage[i] != NULL)
			{
				CComVariant vNsImage;
				hr = m_ppImage[i]->GetProperty(VSI_PROP_IMAGE_NS, &vNsImage);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (0 == lstrcmp(pszNs, V_BSTR(&vNsImage)))
				{
					if (NULL != piSlot)
					{
						*piSlot = i;
					}
					hrFound = S_OK;
					break;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return SUCCEEDED(hr) ? hrFound : hr;
}

STDMETHODIMP CVsiSession::GetIsItemLoaded(LPCOLESTR pszNs, int *piSlot)
{
	HRESULT hr = S_FALSE;

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");

		int iSpt(0);
		LPCOLESTR pszTest(pszNs);
		while (0 != *pszTest)
		{
			if (L'/' == *pszTest)
				++iSpt;
			++pszTest;
		}

		switch (iSpt)
		{
		case 0:
			hr = GetIsStudyLoaded(pszNs, piSlot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			break;
		case 1:
			hr = GetIsSeriesLoaded(pszNs, piSlot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			break;
		case 2:
			hr = GetIsImageLoaded(pszNs, piSlot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			break;
		default:
			VSI_FAIL(VSI_LOG_ERROR, VsiFormatMsg(L"Invalid namespace - %s", pszNs));
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::GetIsItemReviewed(LPCOLESTR pszNs)
{
	HRESULT hr = S_FALSE;

	try
	{
		VSI_CHECK_ARG(pszNs, VSI_LOG_ERROR, L"pszNs");
		
		CString strNs(pszNs);
		
		CVsiSetItemIter iter = m_setReviewed.find(strNs);
		
		if (iter != m_setReviewed.end())
			hr = S_OK;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::ReleaseSlot(int iSlot)
{
	HRESULT hr = S_OK;

	try
	{
		{
			CCritSecLock csl(m_csModeView);
			m_ppModeView[iSlot].Release();
		}

		{
			CCritSecLock csl(m_csMode);
			if (m_ppMode[iSlot] != NULL)
			{
				hr = m_ppMode[iSlot]->Uninitialize();
				if (FAILED(hr))
					VSI_LOG_MSG(VSI_LOG_WARNING, L"ReleaseSlot failed");
				m_ppMode[iSlot].Release();
			}
		}

		{
			CCritSecLock csl(m_csImage);
			m_ppImage[iSlot].Release();
		}

		{
			CCritSecLock csl(m_csSeries);
			m_ppSeries[iSlot].Release();
		}

		{
			CCritSecLock csl(m_csStudy);
			m_ppStudy[iSlot].Release();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiSession::CheckActiveSeries(BOOL bReset, LPCOLESTR pszOperatorId, LPOLESTR *ppszNs)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);
	BOOL bActiveSeries(FALSE);

	try
	{
		// Get data path
		CString strSeriesPath;

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

		bool bSecured = VsiGetBooleanValue<bool>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
			VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
		if (bSecured)
		{
			CComPtr<IVsiOperator> pOperator;
			if (NULL == pszOperatorId)
			{
				hr = m_pOperatorMgr->GetCurrentOperator(&pOperator);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else
			{
				hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_ID, pszOperatorId, &pOperator);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			if (S_OK == hr)
			{
				CComHeapPtr<OLECHAR> pszName;
				hr = pOperator->GetName(&pszName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				VSI_OPERATOR_TYPE dwType;
				hr = pOperator->GetType(&dwType);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
				CString strName(pszName);
				if (VSI_OPERATOR_TYPE_SERVICE_MODE & dwType)
				{
					strName.Trim(L"*");
				}

				strSeriesPath.Format(L"%s\\%s-" VSI_FILE_ACTIVE_SERIES, V_BSTR(&vPath), strName);
			}
		}
		else
		{
			strSeriesPath.Format(L"%s\\" VSI_FILE_ACTIVE_SERIES, V_BSTR(&vPath));
		}

		if (bReset && (NULL == ppszNs))
		{
			if (0 == _waccess(strSeriesPath, 0))
			{
				BOOL bRet = DeleteFile(strSeriesPath);
				VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, NULL);
			}
		}
		else
		{
			CComPtr<IXMLDOMDocument> pSeriesDoc;
			hr = VsiCreateDOMDocument(&pSeriesDoc);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
			CComVariant vSeriesPath((LPCWSTR)strSeriesPath);
			VARIANT_BOOL bLoaded = VARIANT_FALSE;
			hr = pSeriesDoc->load(vSeriesPath, &bLoaded);
			if (S_OK == hr && VARIANT_FALSE != bLoaded)
			{
				CComPtr<IXMLDOMElement> pXmlElmRoot;
				hr = pSeriesDoc->get_documentElement(&pXmlElmRoot);
				VSI_CHECK_HR(hr, VSI_LOG_WARNING, L"Invalid active series file");
				VSI_VERIFY(NULL != pXmlElmRoot, VSI_LOG_WARNING, L"Invalid active series file");

				CComBSTR bstrName;
				hr = pXmlElmRoot->get_nodeName(&bstrName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Invalid active series file");

				if (0 != wcscmp(bstrName, L"activeSeries"))
				{
					// Script tag missing
					VSI_FAIL(VSI_LOG_WARNING, L"Invalid active series file");
				}

				// Get file
				CComVariant vNs;
				hr = pXmlElmRoot->getAttribute(CComBSTR(L"ns"), &vNs);
				if (S_OK == hr)
				{
					if (ppszNs != NULL)
					{
						*ppszNs = AtlAllocTaskWideString(V_BSTR(&vNs));
						VSI_CHECK_MEM(*ppszNs, VSI_LOG_ERROR, NULL);
					}
						
					if (bReset)
					{
						pSeriesDoc.Release();

						if (0 == _waccess(strSeriesPath, 0))
						{
							BOOL bRet = DeleteFile((LPCWSTR)strSeriesPath);
							VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, NULL);
						}
					}

					bActiveSeries = TRUE;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return SUCCEEDED(hr) ? (bActiveSeries ? S_OK : S_FALSE) : hr;
}

STDMETHODIMP CVsiSession::Record(
	DWORD dwEvent,
	const VARIANT *pvParam1,
	const VARIANT *pvParam2)
{
	HRESULT hr = S_OK;

	try
	{
		switch (dwEvent)
		{
		case VSI_SESSION_EVENT_STUDY_NEW:
		case VSI_SESSION_EVENT_STUDY_OPEN:
			break;
		case VSI_SESSION_EVENT_STUDY_COMMIT:
			break;
		case VSI_SESSION_EVENT_STUDY_DISCARD:
			break;

		case VSI_SESSION_EVENT_SERIES_NEW:
			{
				// Get data path
				CString strNewSeriesPath;

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

				bool bSecured = VsiGetBooleanValue<bool>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
					VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
				if (bSecured)
				{
					CComPtr<IVsiOperator> pOperator;
					hr = m_pOperatorMgr->GetCurrentOperator(&pOperator);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						CComHeapPtr<OLECHAR> pszName;
						hr = pOperator->GetName(&pszName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						VSI_OPERATOR_TYPE dwType;
						hr = pOperator->GetType(&dwType);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
						CString strName(pszName);
						if (VSI_OPERATOR_TYPE_SERVICE_MODE & dwType)
						{
							strName.Trim(L"*");
						}

						strNewSeriesPath.Format(L"%s\\%s-" VSI_FILE_ACTIVE_SERIES, V_BSTR(&vPath), strName);
					}
				}
				else
				{
					strNewSeriesPath.Format(L"%s\\" VSI_FILE_ACTIVE_SERIES, V_BSTR(&vPath));
				}

				CComPtr<IXMLDOMDocument> pNewSeriesDoc;
				hr = VsiCreateDOMDocument(&pNewSeriesDoc);			
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComPtr<IXMLDOMProcessingInstruction> pPI;
				hr = pNewSeriesDoc->createProcessingInstruction(
					CComBSTR(L"xml"), CComBSTR(L"version='1.0' encoding='utf-8'"), &pPI);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pNewSeriesDoc->appendChild(pPI, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Series info
				CComPtr<IXMLDOMElement> pRoot;
				hr = pNewSeriesDoc->createElement(CComBSTR(L"activeSeries"), &pRoot);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pRoot->setAttribute(CComBSTR(L"version"), CComVariant(1));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pRoot->setAttribute(CComBSTR(L"ns"), *pvParam1);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pRoot->setAttribute(CComBSTR(L"file"), *pvParam2);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pNewSeriesDoc->appendChild(pRoot, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pNewSeriesDoc->save(CComVariant((LPCWSTR)strNewSeriesPath));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			break;

		case VSI_SESSION_EVENT_SERIES_CLOSE:
			{
				// Delete new series file
				DeleteActiveSeriesFile();
			}
			break;

		case VSI_SESSION_EVENT_SERIES_OPEN:
		case VSI_SESSION_EVENT_SERIES_DEL:
		case VSI_SESSION_EVENT_IMAGE_NEW:
		case VSI_SESSION_EVENT_IMAGE_OPEN:
		case VSI_SESSION_EVENT_IMAGE_RESAVE:
		case VSI_SESSION_EVENT_IMAGE_CLOSE:
		case VSI_SESSION_EVENT_IMAGE_DEL:
			break;

		default:
			_ASSERT(0);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

BOOL CVsiSession::IsStudyIdentical(IVsiStudy *pStudy1, IVsiStudy *pStudy2) throw(...)
{
	HRESULT hr;
	CComVariant vNs1, vNs2;

	hr = pStudy1->GetProperty(VSI_PROP_STUDY_NS, &vNs1);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	hr = pStudy2->GetProperty(VSI_PROP_STUDY_NS, &vNs2);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	return (0 == wcscmp(V_BSTR(&vNs1), V_BSTR(&vNs2)));
}

BOOL CVsiSession::IsSeriesIdentical(IVsiSeries *pSeries1, IVsiSeries *pSeries2) throw(...)
{
	HRESULT hr;
	CComVariant vNs1, vNs2;

	hr = pSeries1->GetProperty(VSI_PROP_SERIES_NS, &vNs1);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	hr = pSeries2->GetProperty(VSI_PROP_SERIES_NS, &vNs2);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	return (0 == wcscmp(V_BSTR(&vNs1), V_BSTR(&vNs2)));
}

BOOL CVsiSession::IsImageIdentical(IVsiImage *pImage1, IVsiImage *pImage2) throw(...)
{
	HRESULT hr;
	CComVariant vNs1, vNs2;

	hr = pImage1->GetProperty(VSI_PROP_IMAGE_NS, &vNs1);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	hr = pImage2->GetProperty(VSI_PROP_IMAGE_NS, &vNs2);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	return (0 == wcscmp(V_BSTR(&vNs1), V_BSTR(&vNs2)));
}

void CVsiSession::DeleteActiveSeriesFile()
{
	HRESULT hr(S_OK);

	try
	{
		CString strNewSeriesPath;

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

		bool bSecured = VsiGetBooleanValue<bool>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
			VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
		if (bSecured)
		{
			CComPtr<IVsiOperator> pOperator;
			hr = m_pOperatorMgr->GetCurrentOperator(&pOperator);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr)
			{
				CComHeapPtr<OLECHAR> pszName;
				hr = pOperator->GetName(&pszName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				VSI_OPERATOR_TYPE dwType;
				hr = pOperator->GetType(&dwType);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
				CString strName(pszName);
				if (VSI_OPERATOR_TYPE_SERVICE_MODE & dwType)
				{
					strName.Trim(L"*");
				}

				strNewSeriesPath.Format(L"%s\\%s-" VSI_FILE_ACTIVE_SERIES, V_BSTR(&vPath), strName);
			}
		}
		else
		{
			strNewSeriesPath.Format(L"%s\\" VSI_FILE_ACTIVE_SERIES, V_BSTR(&vPath));
		}

		if (0 == _waccess((LPCWSTR)strNewSeriesPath, 0))
		{
			BOOL bRet = DeleteFile((LPCWSTR)strNewSeriesPath);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);
}