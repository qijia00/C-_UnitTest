/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiImage.cpp
**
**	Description:
**		Implementation of CVsiImage
**
*******************************************************************************/

#include "stdafx.h"
#include <ATLComTime.h>
#include <VsiBuildNumber.h>
#include <VsiGdiPlus.h>
#include <VsiSaxUtils.h>
#include <VsiCommUtl.h>
#include <VsiPdmModule.h>
#include <VsiConfigXML.h>
#include <VsiParameterXml.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterUtils.h>
#include <VsiParameterSystem.h>
#include <VsiParameterMode.h>
#include <VsiServiceKey.h>
#include <VsiModeModule.h>
#include "VsiStudyXml.h"
#include "VsiImage.h"


// Tag info
typedef struct tagVSI_IMAGE_TAG
{
	DWORD dwId;
	VARTYPE type;
	BOOL bOptional;
	LPCWSTR pszName;
} VSI_IMAGE_TAG;

// Must follow the order defined in IDL
static const VSI_IMAGE_TAG g_pszXmlTag[] =
{
	{ VSI_PROP_IMAGE_ID, VT_BSTR, FALSE, VSI_IMAGE_XML_ATT_ID },
	{ VSI_PROP_IMAGE_ID_SERIES, VT_BSTR, FALSE, VSI_IMAGE_XML_ATT_ID_SERIES },
	{ VSI_PROP_IMAGE_ID_STUDY, VT_BSTR, FALSE, VSI_IMAGE_XML_ATT_ID_STUDY },
	{ VSI_PROP_IMAGE_NS, VT_BSTR, TRUE, NULL },  // No persist
	{ VSI_PROP_IMAGE_NAME, VT_BSTR, FALSE, VSI_IMAGE_XML_ATT_NAME },
	{ VSI_PROP_IMAGE_ACQUIRED_DATE, VT_DATE, FALSE, VSI_IMAGE_XML_ATT_ACQUIRED_DATE },
	{ VSI_PROP_IMAGE_CREATED_DATE, VT_DATE, FALSE, VSI_IMAGE_XML_ATT_CREATED_DATE },
	{ VSI_PROP_IMAGE_MODIFIED_DATE, VT_DATE, FALSE, VSI_IMAGE_XML_ATT_MODIFIED_DATE },
	{ VSI_PROP_IMAGE_MODE, VT_BSTR, FALSE, VSI_IMAGE_XML_ATT_MODE },
	{ VSI_PROP_IMAGE_MODE_DISPLAY, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_MODE_DISPLAY },
	{ VSI_PROP_IMAGE_LENGTH, VT_R8, FALSE, VSI_IMAGE_XML_ATT_LENGTH },
	{ VSI_PROP_IMAGE_ACQID, VT_BSTR, FALSE, VSI_IMAGE_XML_ATT_ACQID },
	{ VSI_PROP_IMAGE_KEY, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_KEY },
	{ VSI_PROP_IMAGE_VERSION_REQUIRED, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_VERSION_REQUIRED },
	{ VSI_PROP_IMAGE_VERSION_CREATED, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_VERSION_CREATED },
	{ VSI_PROP_IMAGE_VERSION_MODIFIED, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_VERSION_MODIFIED },
	{ VSI_PROP_IMAGE_FRAME_TYPE, VT_BOOL, TRUE, VSI_IMAGE_XML_ATT_FRAME_TYPE },
	{ VSI_PROP_IMAGE_VERSION, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_VERSION },
	{ VSI_PROP_IMAGE_THUMBNAIL_FRAME, VT_R8, TRUE, VSI_IMAGE_XML_ATT_THUMBNAIL_FRAME },
	{ VSI_PROP_IMAGE_RF_DATA, VT_BOOL, TRUE, VSI_IMAGE_XML_ATT_RF_DATA },
	{ VSI_PROP_IMAGE_VEVOSTRAIN, VT_BOOL, TRUE, NULL },
	{ VSI_PROP_IMAGE_VEVOSTRAIN_UPDATED, VT_BOOL, TRUE, VSI_IMAGE_XML_ATT_VEVOSTRAIN_UPDATED},
	{ VSI_PROP_IMAGE_VEVOSTRAIN_SERIES_NAME, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_VEVOSTRAIN_SERIES_NAME},
	{ VSI_PROP_IMAGE_VEVOSTRAIN_STUDY_NAME, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_VEVOSTRAIN_STUDY_NAME},
	{ VSI_PROP_IMAGE_VEVOSTRAIN_NAME, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_VEVOSTRAIN_NAME},
	{ VSI_PROP_IMAGE_VEVOSTRAIN_ANIMAL_ID, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_VEVOSTRAIN_ANIMAL_ID},
	{ VSI_PROP_IMAGE_VEVOCQ, VT_BOOL, TRUE, NULL },
	{ VSI_PROP_IMAGE_VEVOCQ_UPDATED, VT_BOOL, TRUE, VSI_IMAGE_XML_ATT_SONOVEVO_UPDATED},
	{ VSI_PROP_IMAGE_VEVOCQ_SERIES_NAME, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_SONOVEVO_SERIES_NAME},
	{ VSI_PROP_IMAGE_VEVOCQ_STUDY_NAME, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_SONOVEVO_STUDY_NAME},
	{ VSI_PROP_IMAGE_VEVOCQ_NAME, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_SONOVEVO_NAME},
	{ VSI_PROP_IMAGE_VEVOCQ_ANIMAL_ID, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_SONOVEVO_ANIMAL_ID},
	{ VSI_PROP_IMAGE_ADVCONTMIP, VT_BOOL, TRUE, NULL },
	{ VSI_PROP_IMAGE_PA_ACQ_TYPE, VT_UI4, TRUE, VSI_IMAGE_XML_ATT_PA_ACQ_TYPE_ID },
	{ VSI_PROP_IMAGE_EKV, VT_BOOL, TRUE, VSI_IMAGE_XML_ATT_EKV_ACQ_TYPE_ID },
	{ VSI_PROP_IMAGE_ANALYSIS, VT_BSTR, TRUE, NULL },  // No persist
	{ VSI_PROP_IMAGE_LAST_MODIFIED_BY, VT_BSTR, TRUE, VSI_IMAGE_XML_ATT_LAST_MODIFIED_BY },
	{ VSI_PROP_IMAGE_LAST_MODIFIED_DATE, VT_DATE, TRUE, VSI_IMAGE_XML_ATT_CREATED_LAST_MODIFIED_DATE },

};


/// <summary>
/// Image reader
/// </summary>
class CVsiImageHandler :
	public CVsiIUnknownImpl<CVsiImageHandler>,
	public IVsiSAXContentHandlerImpl
{
private:
	CVsiMapProp *m_pmapProp;
	CComPtr<IVsiImage> m_pImage;
	CComPtr<ISAXXMLReader> m_pReader;
	CComPtr<ISAXContentHandler> m_pParent;
	CComPtr<IVsiMode> m_pMode;
	CComPtr<IVsiModeData> m_pModeData;
	CComPtr<IVsiModeView> m_pModeView;
	CVsiSaxSkipHandler m_skipHandler;

public:
	// IUnknown
	VSI_IUNKNOWN_IMPL(CVsiImageHandler, ISAXContentHandler)

	CVsiImageHandler(
		CVsiMapProp *pmapProp,
		IVsiImage *pImage,
		IVsiMode *pMode,
		ISAXXMLReader *pReader,
		ISAXContentHandler *pParent = NULL)
	:
		m_pmapProp(pmapProp),
		m_pImage(pImage),
		m_pReader(pReader),
		m_pParent(pParent),
		m_pMode(pMode)
	{
	}

	STDMETHOD(startElement)(
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
			// image?
			bMatched = VsiXmlIsEqualElement(
				pwszLocalName,
				cchLocalName,
				VSI_IMAGE_XML_ELM_IMAGE,
				VSI_STATIC_STRING_COUNT(VSI_IMAGE_XML_ELM_IMAGE)
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
						VSI_VERIFY(g_pszXmlTag[i].bOptional, 
							VSI_LOG_ERROR | VSI_LOG_NO_BOX, 
							VsiFormatMsg(L"Attribute [%s] is missing", g_pszXmlTag[i].pszName));
						hr = S_OK;
						continue;
					}

					VsiReadProperty(
						g_pszXmlTag[i].dwId,
						g_pszXmlTag[i].type,
						pszValue, iValue,
						*m_pmapProp);
				}

				if (m_pMode == NULL)
				{
					// Detach
					m_pReader->putContentHandler(m_pParent);  // Done
				}

				return S_OK;
			}

			// modeData?
			bMatched = VsiXmlIsEqualElement(
				pwszLocalName,
				cchLocalName,
				VSI_IMAGE_XML_ELM_MODE_DATA,
				VSI_STATIC_STRING_COUNT(VSI_IMAGE_XML_ELM_MODE_DATA)
			);
			if (bMatched && m_pMode != NULL)
			{
				hr = m_pMode->GetModeData(&m_pModeData);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = m_pModeData->LoadProperties(
					m_pReader, this, pAttributes);
				if (E_NOTIMPL != hr)  // optional
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				return S_OK;
			}

			// modeView?
			bMatched = VsiXmlIsEqualElement(
				pwszLocalName,
				cchLocalName,
				VSI_IMAGE_XML_ELM_MODE_VIEW,
				VSI_STATIC_STRING_COUNT(VSI_IMAGE_XML_ELM_MODE_VIEW)
			);
			if (bMatched && m_pMode != NULL)
			{
				hr = m_pMode->GetModeView(&m_pModeView);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComQIPtr<IVsiModeViewPersist> pPersist(m_pModeView);
				hr = pPersist->LoadProperties(
					m_pReader, this, pAttributes);
				if (E_NOTIMPL != hr)  // optional
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				return S_OK;
			}

			// settings?
			bMatched = VsiXmlIsEqualElement(
				pwszLocalName,
				cchLocalName,
				VSI_IMAGE_XML_ELM_SETTINGS,
				VSI_STATIC_STRING_COUNT(VSI_IMAGE_XML_ELM_SETTINGS)
			);
			if (bMatched)
			{
				// Skip all settings for now
				m_skipHandler.SetReader(m_pReader);
				m_skipHandler.SetParent(this);

				m_pReader->putContentHandler(&m_skipHandler);
				return S_OK;
			}
		}
		VSI_CATCH(hr);

		return hr;
	}
};


CVsiImage::CVsiImage() :
	m_pSeries(NULL),
	m_dwModified(0),
	m_dwErrorCode(VSI_IMAGE_ERROR_NONE)
{
}

CVsiImage::~CVsiImage()
{
}

STDMETHODIMP CVsiImage::Clone(IVsiImage **ppImage)
{
	HRESULT hr = S_OK;

	try
	{
		CComPtr<IStream> pStm;
		CComPtr<IVsiImage> pImage;
		CComVariant v;

		hr = pImage.CoCreateInstance(__uuidof(VsiImage));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IPersistStreamInit> pPersistDst(pImage);
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

		*ppImage = pImage.Detach();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::GetErrorCode(VSI_IMAGE_ERROR *pdwErrorCode)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pdwErrorCode, VSI_LOG_ERROR, NULL);

		*pdwErrorCode = m_dwErrorCode;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::GetDataPath(LPOLESTR *ppszPath)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(ppszPath, VSI_LOG_ERROR, NULL);

		*ppszPath = AtlAllocTaskWideString((LPCWSTR)m_strFile);
		VSI_CHECK_MEM(*ppszPath, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::SetDataPath(LPCOLESTR pszPath)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pszPath, VSI_LOG_ERROR, NULL);
		m_strFile = pszPath;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::GetSeries(IVsiSeries **ppSeries)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(ppSeries, VSI_LOG_ERROR, NULL);

		if (m_pSeries != NULL)
		{
			*ppSeries = m_pSeries;
			(*ppSeries)->AddRef();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::SetSeries(IVsiSeries *pSeries)
{
	HRESULT hr = S_OK;

	try
	{
		m_pSeries = pSeries;

		// Updates namespace
		if (m_pSeries != NULL)
		{
			UpdateNamespace();
		}
		else
		{
			// Removes namespace
			hr = SetProperty(VSI_PROP_IMAGE_NS, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::GetProperty(VSI_PROP_IMAGE dwProp, VARIANT* pvtValue)
{
	CVsiMapPropIter iterProp = m_mapProp.find(dwProp);
	if (iterProp == m_mapProp.end())
	{
		if (VSI_PROP_IMAGE_NS == dwProp)
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

	return VariantCopy(pvtValue, &iterProp->second);
}

STDMETHODIMP CVsiImage::SetProperty(VSI_PROP_IMAGE dwProp, const VARIANT *pvValue)
{
	m_mapProp.erase(dwProp);

	if (NULL != pvValue && VT_EMPTY != V_VT(pvValue))
		m_mapProp.insert(CVsiMapProp::value_type(dwProp, *pvValue));

	return S_OK;
}

STDMETHODIMP CVsiImage::GetThumbnailFile(LPOLESTR *ppszFile)
{
	HRESULT hr = S_FALSE;

	try
	{
		if (m_strFile.Length() > 0)
		{
			WCHAR szFile[MAX_PATH];
			lstrcpy(szFile, m_strFile);
			
			BOOL bRet = PathRenameExtension(szFile, L"." VSI_FILE_EXT_IMAGE_THUMB);
			VSI_VERIFY(FALSE != bRet, VSI_LOG_ERROR, NULL);

			*ppszFile = AtlAllocTaskWideString(szFile);
			VSI_CHECK_MEM(*ppszFile, VSI_LOG_ERROR, NULL);

			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::GetStudy(IVsiStudy **ppStudy)
{
	HRESULT hr(S_FALSE);

	try
	{
		VSI_CHECK_ARG(ppStudy, VSI_LOG_ERROR, NULL);

		if (m_pSeries != NULL)
		{
			CComPtr<IVsiStudy> pStudy;
			hr = m_pSeries->GetStudy(&pStudy);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			*ppStudy = pStudy;
			(*ppStudy)->AddRef();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::Compare(IVsiImage *pImage, VSI_PROP_IMAGE dwProp, int *piCmp)
{
	HRESULT hr(S_OK);

	try
	{
		VSI_CHECK_ARG(piCmp, VSI_LOG_ERROR, NULL);
		*piCmp = 0;

		HRESULT hr1, hr2;
		CComVariant v1, v2;

		if (VSI_PROP_STUDY_MIN <= dwProp && dwProp <= VSI_PROP_STUDY_MAX)
		{
			CComPtr<IVsiStudy> pStudy1, pStudy2;

			GetStudy(&pStudy1);
			pImage->GetStudy(&pStudy2);
			if (NULL != pStudy1 && NULL != pStudy2)
			{	
				pStudy1->Compare(pStudy2, static_cast<VSI_PROP_STUDY>(dwProp), piCmp);
				if (*piCmp == 0)
				{
					// If same then compare by date
					*piCmp = CompareDates(pImage);
				}
			}
			else if (NULL != pStudy1)
			{
				*piCmp = 1;
			}
			else if (NULL != pStudy2)
			{
				*piCmp = -1;
			}
		}
		else if (VSI_PROP_SERIES_MIN <= dwProp && dwProp <= VSI_PROP_SERIES_MAX)
		{
			CComPtr<IVsiSeries> pSeries1, pSeries2;

			GetSeries(&pSeries1);
			pImage->GetSeries(&pSeries2);
			if (NULL != pSeries1 && NULL != pSeries2)
			{	
				pSeries1->Compare(pSeries2, static_cast<VSI_PROP_SERIES>(dwProp), piCmp);
				if (*piCmp == 0)
				{
					// If same then compare by date
					*piCmp = CompareDates(pImage);
				}
			}
			else if (NULL != pSeries1)
			{
				*piCmp = 1;
			}
			else if (NULL != pSeries2)
			{
				*piCmp = -1;
			}
		}
		else
		{
			switch (dwProp)
			{
			case VSI_PROP_IMAGE_NAME:
			case VSI_PROP_IMAGE_MODE_DISPLAY:
				hr1 = GetProperty(dwProp, &v1);
				hr2 = pImage->GetProperty(dwProp, &v2);
				if (S_OK == hr1 && S_OK == hr2)
				{
					*piCmp = lstrcmp(V_BSTR(&v1), V_BSTR(&v2));
					if (*piCmp == 0)
					{
						// If names are the same then compare by date
						*piCmp = CompareDates(pImage);
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
				break;

			case VSI_PROP_IMAGE_ACQUIRED_DATE:
				*piCmp = CompareDates(pImage);
				break;

			case VSI_PROP_IMAGE_LENGTH:
				{
					CComVariant vMode1, vMode2;
					hr1 = GetProperty(VSI_PROP_IMAGE_MODE, &vMode1);
					hr2 = pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode2);
					if (S_OK == hr1 && S_OK == hr2)
					{
						CComQIPtr<IVsiServiceProvider> pServiceProvider(g_pApp);
						VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

						// Query Mode Manager
						CComPtr<IVsiModeManager> pModeMgr;
						pServiceProvider->QueryService(
							VSI_SERVICE_MODE_MANAGER,
							__uuidof(IVsiModeManager),
							(IUnknown**)&pModeMgr);
						VSI_CHECK_INTERFACE(pModeMgr, VSI_LOG_ERROR, NULL);

						CComVariant vType1, vType2;
						hr1 = pModeMgr->GetModePropertyFromInternalName(
							V_BSTR(&vMode1), L"lengthType", &vType1);
						hr2 = pModeMgr->GetModePropertyFromInternalName(
							V_BSTR(&vMode2), L"lengthType", &vType2);

						if (S_OK == hr1 && S_OK == hr2)
						{
							if (V_I4(&vType1) == V_I4(&vType2))
							{
								// Same type - sort by length now
								hr1 = GetProperty(VSI_PROP_IMAGE_LENGTH, &v1);
								hr2 = pImage->GetProperty(VSI_PROP_IMAGE_LENGTH, &v2);
								if (S_OK == hr1 && S_OK == hr2)
								{
									if (V_R8(&v1) < V_R8(&v2))
									{
										*piCmp = -1;
									}
									else if (V_R8(&v2) < V_R8(&v1))
									{
										*piCmp = 1;
									}
									else
									{
										// If lengths are the same then compare by date
										*piCmp = CompareDates(pImage);
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
							else
							{
								// not same type
								*piCmp = V_I4(&vType1) < V_I4(&vType2) ? -1 : 1;
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

			default:
				_ASSERT(0);
			}  // switch
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::LoadImageData(
	LPCOLESTR pszFile,
	IUnknown *pUnkMode,
	IUnknown *pUnkPdm,
	DWORD dwFlags)
{
	HRESULT hr = S_OK;

	try
	{
		if (NULL != pszFile)
			m_strFile = pszFile;

		VSI_VERIFY(m_strFile.Length() > 0, VSI_LOG_ERROR, NULL);

		m_mapProp.clear();

		m_dwErrorCode |= VSI_IMAGE_ERROR_LOAD_XML;

		CComQIPtr<IVsiMode> pMode;
		if (pUnkMode != NULL)
			pMode = pUnkMode;

		CComQIPtr<IVsiPdm> pPdm;
		if (pUnkPdm != NULL)
			pPdm = pUnkPdm;

		CComPtr<ISAXXMLReader> pReader;
		hr = VsiCreateXmlSaxReader(&pReader);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CVsiImageHandler ih(&m_mapProp, this, pMode, pReader);

		hr = pReader->putContentHandler(static_cast<ISAXContentHandler*>(&ih));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CVsiSAXErrorHandler err;
		hr = pReader->putErrorHandler(&err);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
		hr = pReader->parseURL(m_strFile);
		if (FAILED(hr))
		{
			CComVariant vName;
			if ((S_OK != GetProperty(VSI_PROP_IMAGE_NAME, &vName)) || (0 == *V_BSTR(&vName)))
			{
				CComVariant vError(L"Corrupted Image");
				SetProperty(VSI_PROP_IMAGE_NAME, &vError);
			}

			// set Id and NS in the case of a failure
			CComVariant vId;
			if (S_OK != GetProperty(VSI_PROP_IMAGE_ID, &vId))
			{
				CComVariant vPath(m_strFile);
				hr = SetProperty(VSI_PROP_IMAGE_ID, &vPath);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			if (m_pSeries != NULL)
			{
				UpdateNamespace();
			}

			WCHAR szMsg[300];
			err.FormatErrorMsg(szMsg, _countof(szMsg));
			VSI_LOG_MSG(VSI_LOG_ERROR | VSI_LOG_NO_BOX, VsiFormatMsg(L"Read file '%s' failed", (LPCWSTR)m_strFile));
			VSI_FAIL(VSI_LOG_ERROR | VSI_LOG_NO_BOX, VsiFormatMsg(L"SAXErrorHandler - %s", szMsg));
		}

		// Patch old versions
		hr = DoConversion();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Sets namespace
		{
			CComVariant vId;
			if (S_OK == GetProperty(VSI_PROP_IMAGE_ID, &vId))
			{
				UpdateNamespace();
			}
		}

		CComVariant vMode;
		if (S_OK == GetProperty(VSI_PROP_IMAGE_MODE, &vMode))
		{
			// For performance reasons, don't try accessing the file if this is not the correct mode

			if (0 == lstrcmp(V_BSTR(&vMode), VSI_MODE_KEY_CONTRAST_MODE_ADVANCED))
			{
				// Check if we have sonovevo files for this image
				CString strSono(m_strFile);
				strSono = strSono.Left(strSono.ReverseFind(L'.'));  //remove extension
				strSono += L".sono.dcm";

				CComVariant vSonoVevo(VsiGetFileExists(strSono) ? true : false);
				hr = SetProperty(VSI_PROP_IMAGE_VEVOCQ, &vSonoVevo);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else if (0 == lstrcmp(V_BSTR(&vMode), VSI_MODE_KEY_B_MODE))
			{
				// Check if we have vevostrain files for this image
				CString strStrain(m_strFile);
				strStrain = strStrain.Left(strStrain.ReverseFind(L'.'));  //remove extension
				strStrain += L".dcm";

				CComVariant vVevoStrain(VsiGetFileExists(strStrain) ? true : false);
				hr = SetProperty(VSI_PROP_IMAGE_VEVOSTRAIN, &vVevoStrain);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

		CComVariant vVersionRequired;
		LPCWSTR pszVersionRequired(NULL);
		HRESULT hrTest = GetProperty(VSI_PROP_IMAGE_VERSION_REQUIRED, &vVersionRequired);
		if (S_OK == hrTest)
		{
			pszVersionRequired = V_BSTR(&vVersionRequired);
		}

		// Checks version
		if (NULL != pszVersionRequired)
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
				m_dwErrorCode |= VSI_IMAGE_ERROR_VERSION_INCOMPATIBLE;
			}
		}

		if (0 == (m_dwErrorCode & VSI_IMAGE_ERROR_VERSION_INCOMPATIBLE))
		{
			// Load PDM parameters 1st
			if (pPdm != NULL)
			{
				_ASSERT(pMode != NULL);

				CComQIPtr<IVsiPropertyBag> pProperties(pMode);
				VSI_CHECK_INTERFACE(pProperties, VSI_LOG_ERROR, NULL);

				CComVariant vRoot;
				hr = pProperties->ReadId(VSI_PROP_MODE_ROOT, &vRoot);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Open the image parameter file and get the parameter node.
				CComPtr<IXMLDOMDocument> pImageDoc;
				hr = VsiCreateDOMDocument(&pImageDoc);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
				pImageDoc->put_validateOnParse(VARIANT_TRUE);
				pImageDoc->put_resolveExternals(VARIANT_FALSE);

				VARIANT_BOOL bLoaded = VARIANT_FALSE;
				hr = pImageDoc->load(CComVariant((LPCWSTR)m_strFile), &bLoaded);
				// handle error
				if (hr != S_OK || bLoaded == VARIANT_FALSE)
				{
					CComPtr<IXMLDOMParseError> pParseError;
					HRESULT hrParseError = pImageDoc->get_parseError(&pParseError);
					if (SUCCEEDED(hrParseError))
					{
						long lLine = 0;
						CComBSTR bstrDesc;
						HRESULT hrLine = pParseError->get_line(&lLine);
						HRESULT hrDesc = pParseError->get_reason(&bstrDesc);
						if (SUCCEEDED(hrLine) && SUCCEEDED(hrDesc))
						{
							// Display error
							VSI_FAIL(VSI_LOG_ERROR,
								VsiFormatMsg(L"Read parameter types file '%s' failed - line %d: %s",
									pszFile, lLine, bstrDesc));
						}
						else if (SUCCEEDED(hrDesc))
						{
							// Display error
							VSI_FAIL(VSI_LOG_ERROR,
								VsiFormatMsg(L"Read parameter types file '%s' failed - %s",
									pszFile, bstrDesc));
						}
					}

					VSI_FAIL(VSI_LOG_ERROR, NULL);
				}

				CComQIPtr<IVsiPdmPersist> pPdmPersist(pPdm);
				VSI_CHECK_INTERFACE(pPdmPersist, VSI_LOG_ERROR, NULL);

				CComPtr<IXMLDOMNodeList> pXmlNodeList;
				CComPtr<IXMLDOMNode> pXmlNode;
				CComPtr<IXMLDOMNamedNodeMap> pXmlNameNodeMap;

				// For all roots
				hr = pImageDoc->selectNodes(
					CComBSTR(L"//" VSI_IMAGE_XML_ELM_SETTINGS L"/"
						VSI_CFG_ELM_ROOTS L"/" VSI_CFG_ELM_ROOT),
					&pXmlNodeList);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				LONG lCount = 0;
				hr = pXmlNodeList->get_length(&lCount);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Retrieve each element in the list.
				for (LONG i = 0; i < lCount; ++i)
				{
					pXmlNode.Release();
					hr = pXmlNodeList->get_item(i, &pXmlNode);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Get name attribute
					pXmlNameNodeMap.Release();
					hr = pXmlNode->get_attributes(&pXmlNameNodeMap);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IXMLDOMNode> pXmlNodeName;
					hr = pXmlNameNodeMap->getNamedItem(
						CComBSTR(VSI_ATT_NAME), &pXmlNodeName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					VSI_VERIFY(S_OK == hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"'%s' attribute missing",	VSI_ATT_NAME));

					CComVariant vName;
					hr = pXmlNodeName->get_nodeValue(&vName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Add this root if it does not already exist 
					// (For example, the Export root is not created at startup)
					hr = pPdm->AddRoot(V_BSTR(&vRoot), VSI_PDM_ROOT_ACQ);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pPdmPersist->LoadRoot(V_BSTR(&vRoot), pXmlNode, pszVersionRequired);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				// force auto playback to be off
				CVsiBitfield<ULONG> pState;
				hr = pPdm->GetParameter(
					V_BSTR(&vRoot), VSI_PDM_GROUP_MODE,
					VSI_PARAMETER_SETTINGS_STATE, &pState);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				pState.SetBits(0, VSI_SETTINGS_STATE_AUTOPLAY);

			}

			m_dwErrorCode &= ~VSI_IMAGE_ERROR_LOAD_XML;

			// Load mode data
			if ((pMode != NULL) && (VSI_IMAGE_ERROR_NONE == m_dwErrorCode))
			{
				CComPtr<IVsiModeData> pModeData;
				hr = pMode->GetModeData(&pModeData);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Update mode data properties
				{
					CComQIPtr<IVsiPropertyBag> pProps(pModeData);

					// Data id
					CComVariant vId;
					hr = GetProperty(VSI_PROP_IMAGE_ACQID, &vId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pProps->WriteId(VSI_PROP_MODE_DATA_ID, &vId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Acq date
					CComVariant vDate;
					hr = GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vDate);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					hr = pProps->WriteId(VSI_PROP_MODE_DATA_ACQ_START_TIME, &vDate);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Label
					CComVariant vLabel;
					hr = GetProperty(VSI_PROP_IMAGE_NAME, &vLabel);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						hr = pProps->WriteId(VSI_PROP_MODE_DATA_NAME, &vLabel);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}

				if (0 == (VSI_DATA_TYPE_IMAGE_PARAMS_ONLY & dwFlags))
				{
					// Load data
					hr = pModeData->LoadModeData(static_cast<IVsiImage*>(this));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::SaveImageData(
	LPCOLESTR pszFile,
	IUnknown *pUnkMode,
	IUnknown *pUnkPdm,
	DWORD dwFlags)
{
	HRESULT hr = S_OK;
	HANDLE hTransaction = INVALID_HANDLE_VALUE;

	try
	{
		if (NULL != pszFile)
		{
			m_strFile = pszFile;
		}

		WCHAR szTempFileName[VSI_MAX_PATH];
		errno_t err = wcscpy_s(szTempFileName, m_strFile);
		VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");

		if (0 == (VSI_DATA_TYPE_MSMNT_PARAMS_ONLY & dwFlags))
		{
			hTransaction = VsiTxF::VsiCreateTransaction(
				NULL,
				0,		
				TRANSACTION_DO_NOT_PROMOTE,	
				0,		
				0,		
				INFINITE,
				L"Image xml save transaction");
			VSI_WIN32_VERIFY(INVALID_HANDLE_VALUE != hTransaction, VSI_LOG_ERROR, L"Failure creating transaction");

			// Create a backup of the original file
			WCHAR szBackupFileName[VSI_MAX_PATH];
			err = wcscpy_s(szBackupFileName, m_strFile);
			VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");

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

			// Create a temporary file using SaveImageDataInternal, then use the transacted 
			// file methods to move the temp file back to the original.
			// Use commutl to create unique temp file.

			VsiPathRemoveExtension(szTempFileName);

			WCHAR szUniqueName[VSI_MAX_PATH];
			hr = VsiCreateUniqueName(L"", L"", szUniqueName, VSI_MAX_PATH);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting unique name");

			err = wcscat_s(szTempFileName, szUniqueName);
			VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");
			err = wcscat_s(szTempFileName, L".vxml");
			VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");
		}

		// Saves a .vxml file and a .png file that need to be moved 
		hr = SaveImageDataInternal(
			szTempFileName, 
			pUnkMode,
			pUnkPdm,
			dwFlags);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// S_FALSE when image is not saved
		if (S_OK == hr && 0 == (VSI_DATA_TYPE_MSMNT_PARAMS_ONLY & dwFlags))
		{
			BOOL bRet = VsiTxF::VsiMoveFileTransacted(
				szTempFileName,
				m_strFile,	
				NULL,	
				NULL,	
				MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING,
				hTransaction);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, L"Move file failed");

			// Now move the thumbnail
			WCHAR szThumb[MAX_PATH];
			err = wcscpy_s(szThumb, (LPCWSTR)m_strFile);
			VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");

			bRet = PathRenameExtension(szThumb, L"." VSI_FILE_EXT_IMAGE_THUMB);
			VSI_VERIFY(FALSE != bRet, VSI_LOG_ERROR, NULL);

			bRet = PathRenameExtension(szTempFileName, L"." VSI_FILE_EXT_IMAGE_THUMB);
			VSI_VERIFY(FALSE != bRet, VSI_LOG_ERROR, NULL);

			bRet = VsiTxF::VsiMoveFileTransacted(
				szTempFileName,
				szThumb,	
				NULL,	
				NULL,	
				MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING,
				hTransaction);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, L"Move file failed");

			bRet = VsiTxF::VsiCommitTransaction(hTransaction);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, L"Commit transaction failed");
		}
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

HRESULT CVsiImage::SaveImageDataInternal(
	LPCOLESTR pszFile,
	IUnknown *pUnkMode,
	IUnknown *pUnkPdm,
	DWORD dwFlags)
{
	HRESULT hr = S_OK;
	HBITMAP hBmp = NULL;
	std::unique_ptr<Gdiplus::Bitmap> pBmp;

	try
	{
		VSI_CHECK_ARG(pUnkMode, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiMode> pMode(pUnkMode);
		VSI_CHECK_INTERFACE(pMode, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiModeData> pModeData;
		hr = pMode->GetModeData(&pModeData);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Update image properties
		{
			// Add version modified
			CString strVersion;
			strVersion.Format(
				L"%u.%u.%u.%u",
				VSI_SOFTWARE_MAJOR, VSI_SOFTWARE_MIDDLE, VSI_SOFTWARE_MINOR, VSI_BUILD_NUMBER);		

			CComVariant vVersionModified(strVersion.GetString());
			hr = SetProperty(VSI_PROP_IMAGE_VERSION_MODIFIED, &vVersionModified);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);


			CComQIPtr<IVsiPropertyBag> pProps(pModeData);

			// Id
			CComVariant vId;
			hr = pProps->ReadId(VSI_PROP_MODE_DATA_ID, &vId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = SetProperty(VSI_PROP_IMAGE_ACQID, &vId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Label
			CComVariant vLabel;
			hr = pProps->ReadId(VSI_PROP_MODE_DATA_NAME, &vLabel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr)
			{
				hr = SetProperty(VSI_PROP_IMAGE_NAME, &vLabel);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

		CComPtr<IVsiModeView> pModeView;
		hr = pMode->GetModeView(&pModeView);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiModeViewPersist> pModeViewPersist(pModeView);
		VSI_CHECK_INTERFACE(pModeViewPersist, VSI_LOG_ERROR, NULL);

		// Init mode save flags
		DWORD dwModeSaveFlags = 0;
		if (VSI_DATA_TYPE_RESAVE & dwFlags)
		{
			dwModeSaveFlags |= VSI_MODE_DATA_SAVE_RESAVE;
		}
		if (VSI_DATA_TYPE_FRAME & dwFlags)
		{
			dwModeSaveFlags |= VSI_MODE_DATA_SAVE_FRAME;
		}
		if (VSI_DATA_TYPE_LOOP & dwFlags)
		{
			dwModeSaveFlags |= VSI_MODE_DATA_SAVE_LOOP;
		}
		if (VSI_DATA_TYPE_IMAGE_PARAMS_ONLY & dwFlags)
		{
			dwModeSaveFlags |= VSI_MODE_DATA_SAVE_IMAGE_ONLY;
		}

		// Save mode data
		hr = pModeViewPersist->SaveModeData(
			static_cast<IVsiImage*>(this), dwModeSaveFlags);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// If we return S_FALSE then the user chose not to save this image
		if (hr == S_OK)
		{
			// Saves thumbnail
			CComQIPtr<IVsiModeExport> pExport(pModeView);
			if (pExport != NULL)
			{
				WCHAR szThumb[MAX_PATH];
				lstrcpyn(szThumb, (LPCWSTR)pszFile, _countof(szThumb));

				BOOL bRet = PathRenameExtension(szThumb, L"." VSI_FILE_EXT_IMAGE_THUMB);
				VSI_VERIFY(FALSE != bRet, VSI_LOG_ERROR, NULL);

				DWORD dwFalgs = VSI_MODE_CAPTURE_SECTION | VSI_MODE_CAPTURE_THUMBNAIL;
				if (VSI_DATA_TYPE_FRAME & dwFlags)
					dwFalgs |= VSI_MODE_CAPTURE_FRAME;
				if (VSI_DATA_TYPE_LOOP & dwFlags)
					dwFalgs |= VSI_MODE_CAPTURE_LOOP;

				hr = pExport->CaptureImage(
					&hBmp,
					NULL,
					NULL,
					NULL,
					32,
					dwFalgs,
					NULL, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				pBmp.reset(Gdiplus::Bitmap::FromHBITMAP(hBmp, NULL));
				VSI_CHECK_PTR(pBmp.get(), VSI_LOG_ERROR, NULL);

				CLSID clsidEncoder;
				VsiGetEncoderClsid(L"image/png", &clsidEncoder);

				Gdiplus::Status stat = pBmp->Save(szThumb, &clsidEncoder, NULL);
				VSI_VERIFY(Gdiplus::Ok == stat, VSI_LOG_ERROR, NULL);
			}

			if (0 == (VSI_DATA_TYPE_MSMNT_PARAMS_ONLY & dwFlags))
			{
				//
				// Saves XML
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
					dwTagXmlIndex = iter->first - VSI_PROP_IMAGE_MIN - 1;
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
					VSI_IMAGE_XML_ELM_IMAGE, VSI_STATIC_STRING_COUNT(VSI_IMAGE_XML_ELM_IMAGE), pa.p);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				pmxa->clear();

				// Saves mode data properties
				hr = pModeData->SaveProperties(pch);
				if (E_NOTIMPL != hr)
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Saves mode view properties
				hr = pModeViewPersist->SaveProperties(pch);
				if (E_NOTIMPL != hr)
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Creates settings
				hr = pch->startElement(szNull, 0, szNull, 0,
					VSI_IMAGE_XML_ELM_SETTINGS, VSI_STATIC_STRING_COUNT(VSI_IMAGE_XML_ELM_SETTINGS), pa.p);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pch->startElement(szNull, 0, szNull, 0,
					VSI_CFG_ELM_ROOTS, VSI_STATIC_STRING_COUNT(VSI_CFG_ELM_ROOTS), pa.p);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pch->endElement(szNull, 0, szNull, 0,
					VSI_CFG_ELM_ROOTS, VSI_STATIC_STRING_COUNT(VSI_CFG_ELM_ROOTS));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pch->endElement(szNull, 0, szNull, 0,
					VSI_IMAGE_XML_ELM_SETTINGS, VSI_STATIC_STRING_COUNT(VSI_IMAGE_XML_ELM_SETTINGS));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pch->endElement(szNull, 0, szNull, 0,
					VSI_IMAGE_XML_ELM_IMAGE, VSI_STATIC_STRING_COUNT(VSI_IMAGE_XML_ELM_IMAGE));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pch->endDocument();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				out.Close();

				// Save parameters
				{
					CComQIPtr<IVsiPdm> pPdm(pUnkPdm);
					VSI_CHECK_INTERFACE(pPdm, VSI_LOG_ERROR, NULL);

					CComQIPtr<IVsiPdmPersist> pPdmPersist(pPdm);
					VSI_CHECK_INTERFACE(pPdmPersist, VSI_LOG_ERROR, NULL);

					// Open the image parameter file and get the parameter node.
					CComPtr<IXMLDOMDocument> pImageDoc;
					hr = VsiCreateDOMDocument(&pImageDoc);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
					pImageDoc->put_validateOnParse(VARIANT_TRUE);
					pImageDoc->put_resolveExternals(VARIANT_FALSE);

					VARIANT_BOOL bLoaded = VARIANT_FALSE;
					hr = pImageDoc->load(CComVariant((LPCWSTR)pszFile), &bLoaded);
					// handle error
					if (hr != S_OK || bLoaded == VARIANT_FALSE)
					{
						CComPtr<IXMLDOMParseError> pParseError;
						HRESULT hrParseError = pImageDoc->get_parseError(&pParseError);
						if (SUCCEEDED(hrParseError))
						{
							long lLine = 0;
							CComBSTR bstrDesc;
							HRESULT hrLine = pParseError->get_line(&lLine);
							HRESULT hrDesc = pParseError->get_reason(&bstrDesc);
							if (SUCCEEDED(hrLine) && SUCCEEDED(hrDesc))
							{
								// Display error
								VSI_FAIL(VSI_LOG_ERROR,
									VsiFormatMsg(L"Read parameter types file '%s' failed - line %d: %s",
										pszFile, lLine, bstrDesc));
							}
							else if (SUCCEEDED(hrDesc))
							{
								// Display error
								VSI_FAIL(VSI_LOG_ERROR,
									VsiFormatMsg(L"Read parameter types file '%s' failed - %s",
										pszFile, bstrDesc));
							}
						}

						VSI_FAIL(VSI_LOG_ERROR, NULL);
					}

					CComPtr<IXMLDOMNode> pXmlNodeRoots;

					// Gets "roots" node
					hr = pImageDoc->selectSingleNode(
						CComBSTR(L"//" VSI_IMAGE_XML_ELM_SETTINGS L"/" VSI_CFG_ELM_ROOTS),
						&pXmlNodeRoots);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IXMLDOMElement> pReviewParameters;
					hr = pImageDoc->createElement(CComBSTR(VSI_CFG_ELM_ROOT), &pReviewParameters);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pReviewParameters->setAttribute(
						CComBSTR(VSI_PDM_ATT_NAME), CComVariant(VSI_PDM_ROOT_REVIEW));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pXmlNodeRoots->appendChild(pReviewParameters, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComQIPtr<IVsiPropertyBag> pProperties(pMode);
					VSI_CHECK_INTERFACE(pProperties, VSI_LOG_ERROR, NULL);

					CComVariant vRoot;
					hr = pProperties->ReadId(VSI_PROP_MODE_ROOT, &vRoot);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					// force auto playback to be off
					CVsiBitfield<ULONG> pState;
					hr = pPdm->GetParameter(
						V_BSTR(&vRoot), VSI_PDM_GROUP_MODE,
						VSI_PARAMETER_SETTINGS_STATE, &pState);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					pState.SetBits(0, VSI_SETTINGS_STATE_AUTOPLAY);

					// Temp. modify VSI_PARAMETER_SETTINGS_SAVED_FRAME before saving the 
					// parameter if this is not a resave
					// e.g. Frame Store on a cine, VSI_PARAMETER_SETTINGS_SAVED_FRAME should set to true
					BOOL bSavedFrame(FALSE);

					CComQIPtr<IVsiModeFrameData> pFrameData(pModeData);

					if ((NULL != pFrameData) && (0 == (VSI_DATA_TYPE_RESAVE & dwFlags)))
					{
						bSavedFrame = VsiGetBooleanValue<BOOL>(
							V_BSTR(&vRoot), VSI_PDM_GROUP_GLOBAL,
							VSI_PARAMETER_SETTINGS_SAVED_FRAME, pPdm);

						// Set SavedFrame param
						VsiSetBooleanValue<BOOL>(
							(VSI_DATA_TYPE_FRAME & dwFlags) > 0,
							V_BSTR(&vRoot), VSI_PDM_GROUP_GLOBAL,
							VSI_PARAMETER_SETTINGS_SAVED_FRAME, pPdm);
					}

					// Get persist namespaces
					CComQIPtr<IVsiServiceProvider> pServiceProvider(g_pApp);
					VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

					// Query Mode Manager
					CComPtr<IVsiModeManager> pModeMgr;
					pServiceProvider->QueryService(
						VSI_SERVICE_MODE_MANAGER,
						__uuidof(IVsiModeManager),
						(IUnknown**)&pModeMgr);
					VSI_CHECK_INTERFACE(pModeMgr, VSI_LOG_ERROR, NULL);

					CComQIPtr<IVsiPropertyBag> pProps(pMode);
					VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

					CComVariant vModeId;
					hr = pProps->ReadId(VSI_PROP_MODE_MODE_ID, &vModeId);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					CComVariant vPersistNs;
					hr = pModeMgr->GetModePropertyFromId(V_I4(&vModeId), L"persistNs", &vPersistNs);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					// Save PDM root
					hr = pPdmPersist->SaveRoot(V_BSTR(&vRoot), V_BSTR(&vPersistNs), pReviewParameters);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if ((NULL != pFrameData) && (0 == (VSI_DATA_TYPE_RESAVE & dwFlags)))
					{
						// Restore SavedFrame param
						VsiSetBooleanValue<BOOL>(
							bSavedFrame,
							V_BSTR(&vRoot), VSI_PDM_GROUP_GLOBAL,
							VSI_PARAMETER_SETTINGS_SAVED_FRAME, pPdm);
					}

					// Parse into sax reader
					hr = SaxParseDom(pImageDoc, pszFile);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}
		}
	}
	VSI_CATCH(hr);

	pBmp.reset();

	if (NULL != hBmp)
		DeleteObject(hBmp);

	return hr;
}

STDMETHODIMP CVsiImage::ResaveProperties(LPCOLESTR pszFile)
{
	HRESULT hr = S_OK;

	try
	{
		if (NULL == pszFile)
			pszFile = m_strFile;

		VSI_VERIFY(0 != *pszFile, VSI_LOG_ERROR, NULL);

		CComPtr<IXMLDOMDocument> pXmlDoc;
		hr = VsiCreateDOMDocument(&pXmlDoc);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Load the XML document
		VARIANT_BOOL bLoaded = VARIANT_FALSE;
		hr = pXmlDoc->load(CComVariant(pszFile), &bLoaded);
		if (hr != S_OK || bLoaded == VARIANT_FALSE)
		{
			CComPtr<IXMLDOMParseError> pParseError;
			hr = pXmlDoc->get_parseError(&pParseError);
			if (SUCCEEDED(hr))
			{
				CComBSTR bstrDesc;
				hr = pParseError->get_reason(&bstrDesc);
				if (SUCCEEDED(hr))
				{
					// Display error
					VSI_LOG_MSG(VSI_LOG_ERROR,
						VsiFormatMsg(L"Load image file failed - %s", bstrDesc));
				}
			}
			else
			{
				VSI_LOG_MSG(VSI_LOG_INFO, L"Failure getting xml error type");
			}

			VSI_FAIL(VSI_LOG_ERROR, VsiFormatMsg(L"Load image file '%s' failed", pszFile));
		}

		CComPtr<IXMLDOMNode> pNode;
		hr = pXmlDoc->selectSingleNode(CComBSTR(L"/" VSI_IMAGE_XML_ELM_IMAGE), &pNode);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IXMLDOMElement> pElm(pNode);

		DWORD dwTagXmlIndex = 0;
		CVsiMapPropIter iter;
		for (iter = m_mapProp.begin(); iter != m_mapProp.end(); ++iter)
		{
			dwTagXmlIndex = iter->first - VSI_PROP_IMAGE_MIN - 1;
			_ASSERT(g_pszXmlTag[dwTagXmlIndex].dwId == iter->first);

			if (VT_EMPTY == g_pszXmlTag[dwTagXmlIndex].type)
				continue;  // Not used

			if (NULL == g_pszXmlTag[dwTagXmlIndex].pszName)
				continue;  // No persist

			CComBSTR bstrName(g_pszXmlTag[dwTagXmlIndex].pszName);

			CComVariant v;

			hr = VsiXmlVariantToVariantXml(&(iter->second), &v);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pElm->setAttribute(bstrName, v);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		hr = pXmlDoc->save(CComVariant(pszFile));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::InitNew()
{
	HRESULT hr = S_OK;

	try
	{
		m_strFile.Empty();
		m_dwErrorCode = VSI_IMAGE_ERROR_NONE;
		m_mapProp.clear();
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiImage::GetClassID(CLSID *pClassID)
{
	if (pClassID == NULL)
		return E_POINTER;

	*pClassID = GetObjectCLSID();

	return S_OK;
}

STDMETHODIMP CVsiImage::IsDirty()
{
	return (m_dwModified > 0) ? S_OK : S_FALSE;
}

STDMETHODIMP CVsiImage::Load(LPSTREAM pStm)
{
	HRESULT hr = S_OK;
	ULONG ulRead;

	try
	{
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

STDMETHODIMP CVsiImage::Save(LPSTREAM pStm, BOOL fClearDirty)
{
	UNREFERENCED_PARAMETER(fClearDirty);

	HRESULT hr = S_OK;
	ULONG ulWritten;

	try
	{
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

STDMETHODIMP CVsiImage::GetSizeMax(ULARGE_INTEGER* pcbSize)
{
	UNREFERENCED_PARAMETER(pcbSize);

	return E_NOTIMPL;
}

HRESULT CVsiImage::SaxParseDom(IXMLDOMDocument *pXmlDoc, LPCWSTR pszFile)
{
	HRESULT hr = S_OK;

	try
	{
		CComPtr<ISAXXMLReader> pReader;
		hr = VsiCreateXmlSaxReader(&pReader);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComPtr<IMXWriter> pWriter;
		hr = VsiCreateXmlSaxWriter(&pWriter);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pWriter->put_indent(VARIANT_TRUE);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pWriter->put_encoding(CComBSTR(L"utf-8"));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<ISAXContentHandler> pch(pWriter);
		VSI_CHECK_INTERFACE(pch, VSI_LOG_ERROR, NULL);

		hr = pReader->putContentHandler(pch);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CVsiMXFileWriterStream out;
		hr = out.Open(pszFile);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pWriter->put_output(CComVariant(static_cast<IStream*>(&out)));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CVsiSAXErrorHandler err;
		hr = pReader->putErrorHandler(&err);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pReader->parse(CComVariant(pXmlDoc));
		if (FAILED(hr))
		{
			WCHAR szMsg[300];
			err.FormatErrorMsg(szMsg, _countof(szMsg));
			VSI_LOG_MSG(VSI_LOG_ERROR, VsiFormatMsg(L"Write file '%s' failed", pszFile));
			VSI_LOG_MSG(VSI_LOG_ERROR, szMsg);
		}
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		out.Close();
	}
	VSI_CATCH(hr);

	return hr;
}

void CVsiImage::UpdateNamespace() throw(...)
{
	HRESULT hr;

	CComVariant vImageId;
	hr = GetProperty(VSI_PROP_IMAGE_ID, &vImageId);
	if (S_OK == hr)
	{
		VSI_VERIFY(m_pSeries != NULL, VSI_LOG_ERROR, NULL);

		// Gets study Id
		CComPtr<IVsiStudy> pStudy;
		m_pSeries->GetStudy(&pStudy);
		if (pStudy != NULL)
		{
			CComVariant vStudyId;
			hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vStudyId);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CComVariant vSeriesId;
			hr = m_pSeries->GetProperty(VSI_PROP_SERIES_ID, &vSeriesId);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			// Formats
			CString strNs;
			strNs.Format(L"%s/%s/%s",
				V_BSTR(&vStudyId), V_BSTR(&vSeriesId), V_BSTR(&vImageId));

			// Sets
			CComVariant vNs((LPCWSTR)strNs);
			hr = SetProperty(VSI_PROP_IMAGE_NS, &vNs);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
}

/// <summary>
/// Perform image data upgrade
/// </summary>
HRESULT CVsiImage::DoConversion()
{
	HRESULT hr = S_OK;

	try
	{
		CComVariant vMode;
		if (S_OK == GetProperty(VSI_PROP_IMAGE_MODE, &vMode))
		{
			// 1201-(2009-4-8) - Change "Contrast Mode - Advanced" to "Contrast Mode - Nonlinear"
			if (0 == lstrcmp(V_BSTR(&vMode), L"Contrast Mode - Advanced"))
			{
				vMode = VSI_MODE_KEY_CONTRAST_MODE_ADVANCED;

				hr = SetProperty(VSI_PROP_IMAGE_MODE, &vMode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

		// Older version doesn't have VSI_PROP_IMAGE_MODE_DISPLAY
		CComVariant vModeDisplay;
		if (S_OK == GetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vModeDisplay))
		{
			// 1201-(2009-4-8  6-0-36) - Change Contrast Mode to Linear Contrast Mode
			// 1201-(2009-4-8  6-0-36) - Change Contrast 3D Mode to Linear Contrast 3D Mode
			// 1201-(2009-4-8  6-0-36) - Change RF Contrast Mode to Linear RF Contrast Mode
			// 1201-(2009-4-8  6-0-36) - Change RF Contrast 3D Mode to RF Linear Contrast 3D Mode
			if (0 == lstrcmp(V_BSTR(&vModeDisplay), L"Contrast Mode"))
			{
				vModeDisplay = L"Linear Contrast Mode";
				
				hr = SetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vModeDisplay);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else if (0 == lstrcmp(V_BSTR(&vModeDisplay), L"Contrast 3D Mode"))
			{
				vModeDisplay = L"Linear Contrast 3D Mode";
				
				hr = SetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vModeDisplay);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else if (0 == lstrcmp(V_BSTR(&vModeDisplay), L"RF Contrast Mode"))
			{
				vModeDisplay = L"RF Linear Contrast Mode";
				
				hr = SetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vModeDisplay);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else if (0 == lstrcmp(V_BSTR(&vModeDisplay), L"RF Contrast 3D Mode"))
			{
				vModeDisplay = L"RF Linear Contrast 3D Mode";
				
				hr = SetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vModeDisplay);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else if (0 == lstrcmp(V_BSTR(&vModeDisplay), L"PA-Mode") ||
				0 == lstrcmp(V_BSTR(&vModeDisplay), L"PA-Mode 3D") ||
				0 == lstrcmp(V_BSTR(&vModeDisplay), L"RF PA-Mode") ||
				0 == lstrcmp(V_BSTR(&vModeDisplay), L"RF PA-Mode 3D"))
			{
				CString strName(vModeDisplay);

				// how do we get sub mode from this
				CComVariant vPaAcqType;
				if (S_OK == GetProperty(VSI_PROP_IMAGE_PA_ACQ_TYPE, &vPaAcqType))
				{
					vPaAcqType.ChangeType(VT_I4);
					DWORD dwType = V_I4(&vPaAcqType);
					switch (dwType)
					{
					case VSI_PA_MODE_ACQUISITION_SINGLE:
						// No change
						break;
					case VSI_PA_MODE_ACQUISITION_OXYHEMO:
						strName += L" (Oxy-Hemo)";
						break;
					case VSI_PA_MODE_ACQUISITION_NANOSTEPPER:
						strName += L" (NanoStepper)";
						break;
					case VSI_PA_MODE_ACQUISITION_SPECTRO:
						strName += L" (Spectro)";
						break;
					}
				}

				vModeDisplay = strName;
				hr = SetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vModeDisplay);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}			
		}
		else
		{
			// Older version - copy from VSI_PROP_IMAGE_MODE
			CComVariant vMode;
			if (S_OK == GetProperty(VSI_PROP_IMAGE_MODE, &vMode))
			{
				CComQIPtr<IVsiServiceProvider> pServiceProvider(g_pApp);
				VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

				// Query Mode Manager
				CComPtr<IVsiModeManager> pModeMgr;
				pServiceProvider->QueryService(
					VSI_SERVICE_MODE_MANAGER,
					__uuidof(IVsiModeManager),
					(IUnknown**)&pModeMgr);
				VSI_CHECK_INTERFACE(pModeMgr, VSI_LOG_ERROR, NULL);

				hr = pModeMgr->GetModeNameFromInternalName(
					V_BSTR(&vMode), VSI_MODE_FLAG_NONE, &vModeDisplay);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	
				hr = SetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vModeDisplay);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}

	}
	VSI_CATCH(hr);

	return hr;
}

int CVsiImage::CompareDates(IVsiImage* pImage)
{
	int iCmp = 0;
	HRESULT hr1, hr2;

	CComVariant v1, v2;
	hr1 = GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &v1);
	hr2 = pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &v2);
	if (S_OK == hr1 && S_OK == hr2)
	{
		COleDateTime cDate1(v1), cDate2(v2);

		// Same if the differences are less than a second
		if (abs(cDate1.m_dt - cDate2.m_dt) < (1.0 / 24.0 / 60.0 / 60.0))
		{
			// Same - checks creation date
			CComVariant vCreate1, vCreate2;
			hr1 = GetProperty(VSI_PROP_IMAGE_CREATED_DATE, &vCreate1);
			hr2 = pImage->GetProperty(VSI_PROP_IMAGE_CREATED_DATE, &vCreate2);
			if (S_OK == hr1 && S_OK == hr2)
			{
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
			else if (S_OK == hr1)
			{
				iCmp = 1;
			}
			else if (S_OK == hr2)
			{
				iCmp = -1;
			}
		}
		else if (cDate1 < cDate2)
		{
			iCmp = -1;
		}
		else if (cDate2 < cDate1)
		{
			iCmp = 1;
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
