/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperator.cpp
**
**	Description:
**		Implementation of CVsiOperator
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiResProduct.h>
#include <VsiSaxUtils.h>
#include <VsiCommUtlLib.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterUtils.h>
#include "VsiOperatorXml.h"
#include "VsiOperator.h"



CVsiOperator::CVsiOperator() :
	m_dwType(VSI_OPERATOR_TYPE_STANDARD),
	m_dwState(VSI_OPERATOR_STATE_ENABLE),
	m_dwDefaultStudyAccess(VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE)
{
}

CVsiOperator::~CVsiOperator()
{
}

STDMETHODIMP CVsiOperator::GetId(LPOLESTR *ppszId)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(ppszId, VSI_LOG_ERROR, L"GetId failed");

		*ppszId = AtlAllocTaskWideString((LPCWSTR)m_strId);
		VSI_CHECK_MEM(*ppszId, VSI_LOG_ERROR, L"GetId failed");
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::SetId(LPCOLESTR pszId)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pszId, VSI_LOG_ERROR, L"SetId failed");

		m_strId = pszId;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::GetName(LPOLESTR *ppszName)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(ppszName, VSI_LOG_ERROR, L"GetName failed");

		*ppszName = AtlAllocTaskWideString(m_strName.GetString());
		VSI_CHECK_MEM(*ppszName, VSI_LOG_ERROR, L"GetName failed");
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::SetName(LPCOLESTR pszName)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pszName, VSI_LOG_ERROR, L"SetName failed");

		m_strName = pszName;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::GetPassword(LPOLESTR *ppszPassword)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(ppszPassword, VSI_LOG_ERROR, L"GetPassword failed");

		*ppszPassword = AtlAllocTaskWideString((LPCWSTR)m_strPassword);
		VSI_CHECK_MEM(*ppszPassword, VSI_LOG_ERROR, L"GetPassword failed");
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::SetPassword(LPCOLESTR pszPassword, BOOL bHashed)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pszPassword, VSI_LOG_ERROR, L"SetPassword failed");

		// Hash the password
		if (bHashed)
		{
			m_strPassword = pszPassword;
		}
		else
		{
			VSI_VERIFY(!m_strName.IsEmpty(), VSI_LOG_ERROR, L"Name missing");

			CString strMessage(m_strName);
			strMessage += L"^";
			strMessage += pszPassword;

			hr = VsiCreateSha1(
				(const BYTE*)strMessage.GetString(),
				strMessage.GetLength() * sizeof(WCHAR),
				m_strPassword.GetBufferSetLength(41),
				41);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_strPassword.ReleaseBuffer();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::TestPassword(LPCOLESTR pszPassword)
{
	HRESULT hr = S_FALSE;
	BOOL bPassed = FALSE;

	try
	{
		VSI_CHECK_ARG(pszPassword, VSI_LOG_ERROR, L"TestPassword failed");
		VSI_VERIFY(!m_strName.IsEmpty(), VSI_LOG_ERROR, L"Name missing");

		CString strHashed;

		if (VSI_OPERATOR_TYPE_SERVICE_MODE & m_dwType)
		{
			SYSTEMTIME stLocal;
			GetLocalTime(&stLocal);

			CString strDate;
			strDate.Format(L"%4d%2d%2d", stLocal.wYear, stLocal.wMonth, stLocal.wDay);

			hr = VsiCreateSha1(
				(const BYTE*)strDate.GetString(),
				strDate.GetLength() * sizeof(WCHAR),
				strHashed.GetBufferSetLength(41),
				41);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			strHashed.ReleaseBuffer();

			bPassed = pszPassword == strHashed.Left(10);
		}
		else
		{
			// Hash the password
			CString strMessage(m_strName);
			strMessage += L"^";
			strMessage += pszPassword;

			hr = VsiCreateSha1(
				(const BYTE*)strMessage.GetString(),
				strMessage.GetLength() * sizeof(WCHAR),
				strHashed.GetBufferSetLength(41),
				41);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			strHashed.ReleaseBuffer();

			bPassed = m_strPassword == strHashed;
		}
	}
	VSI_CATCH(hr);

	return bPassed ? S_OK : SUCCEEDED(hr) ? S_FALSE : hr;
}

STDMETHODIMP CVsiOperator::HasPassword()
{
	BOOL bHasPassword = FALSE;

	HRESULT hr = S_OK;

	try
	{
		if (VSI_OPERATOR_TYPE_SERVICE_MODE & m_dwType)
		{
			bHasPassword = TRUE;
		}
		else
		{
			VSI_VERIFY(!m_strName.IsEmpty(), VSI_LOG_ERROR, L"Name missing");

			if (!m_strPassword.IsEmpty())
			{
				CString strMessage(m_strName);
				strMessage += L"^";

				CString strEmptyPassword;
				hr = VsiCreateSha1(
					(const BYTE*)strMessage.GetString(),
					strMessage.GetLength() * sizeof(WCHAR),
					strEmptyPassword.GetBufferSetLength(41),
					41);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				strEmptyPassword.ReleaseBuffer();

				bHasPassword = strEmptyPassword != m_strPassword;
			}
		}
	}
	VSI_CATCH(hr);

	return SUCCEEDED(hr) ? (bHasPassword ? S_OK : S_FALSE) : hr;
}


STDMETHODIMP CVsiOperator::GetGroup(LPOLESTR *ppszGroup)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(ppszGroup, VSI_LOG_ERROR, L"GetGroup failed");

		*ppszGroup = AtlAllocTaskWideString((LPCWSTR)m_strGroup);
		VSI_CHECK_MEM(*ppszGroup, VSI_LOG_ERROR, L"GetGroup failed");
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::SetGroup(LPCOLESTR pszGroup)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pszGroup, VSI_LOG_ERROR, L"SetGroup failed");

		m_strGroup = pszGroup;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::GetType(VSI_OPERATOR_TYPE *pdwType)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pdwType, VSI_LOG_ERROR, NULL);

		*pdwType = m_dwType;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::SetType(VSI_OPERATOR_TYPE dwType)
{
	m_dwType = dwType;

	return S_OK;
}

STDMETHODIMP CVsiOperator::GetState(VSI_OPERATOR_STATE *pdwState)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pdwState, VSI_LOG_ERROR, NULL);

		*pdwState = m_dwState;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::SetState(VSI_OPERATOR_STATE dwState)
{
	m_dwState = dwState;

	return S_OK;
}

STDMETHODIMP CVsiOperator::GetDefaultStudyAccess(VSI_OPERATOR_DEFAULT_STUDY_ACCESS *pdwStudyAccess)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pdwStudyAccess, VSI_LOG_ERROR, NULL);

		*pdwStudyAccess = m_dwDefaultStudyAccess;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::SetDefaultStudyAccess(VSI_OPERATOR_DEFAULT_STUDY_ACCESS dwStudyAccess)
{
	m_dwDefaultStudyAccess = dwStudyAccess;

	return S_OK;
}

STDMETHODIMP CVsiOperator::GetSessionInfo(VSI_OPERATOR_SESSION_INFO *pInfo)
{
	HRESULT hr = S_OK;
	CCritSecLock cs(m_cs);

	try
	{
		VSI_CHECK_ARG(pInfo, VSI_LOG_ERROR, NULL);

		CopyMemory(pInfo, &m_sessionInfo, sizeof(VSI_OPERATOR_SESSION_INFO));
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperator::SetSessionInfo(const VSI_OPERATOR_SESSION_INFO *pInfo)
{
	HRESULT hr = S_OK;
	CCritSecLock cs(m_cs);

	try
	{
		VSI_CHECK_ARG(pInfo, VSI_LOG_ERROR, NULL);

		CopyMemory(&m_sessionInfo, pInfo, sizeof(VSI_OPERATOR_SESSION_INFO));
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// CVsiOperator::LoadXml
/// </summary>
STDMETHODIMP CVsiOperator::LoadXml(IXMLDOMNode* pXmlNode)
{
	HRESULT hr = S_OK;
	CComPtr<IXMLDOMNamedNodeMap> pXmlNameNodeMap;
	CComPtr<IXMLDOMNode> pOperatorNode;

	CCritSecLock cs(m_cs);

	try
	{
		CComQIPtr<IXMLDOMElement> pOperatorElm(pXmlNode);
		VSI_CHECK_INTERFACE(pOperatorElm, VSI_LOG_ERROR, NULL);

		// Name
		CComVariant vName;
		hr = pOperatorElm->getAttribute(CComBSTR(VSI_OPERATOR_ATT_NAME), &vName);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		SetName(V_BSTR(&vName));

		// ID
		CComVariant vId;
		hr = pOperatorElm->getAttribute(CComBSTR(VSI_OPERATOR_ATT_UNIQUEID), &vId);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		SetId(V_BSTR(&vId));

		// Password
		CComVariant vPassword;
		hr = pOperatorElm->getAttribute(CComBSTR(VSI_OPERATOR_ATT_PASSWORD), &vPassword);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		if (S_OK == hr)
		{
			SetPassword(V_BSTR(&vPassword), TRUE);
		}

		// Type
		CComVariant vType;
		hr = pOperatorElm->getAttribute(CComBSTR(VSI_OPERATOR_ATT_TYPE), &vType);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		VsiVariantDecodeValue(vType, VT_UI4);
		SetType(static_cast<VSI_OPERATOR_TYPE>(V_UI4(&vType)));

		// State
		CComVariant vState;
		hr = pOperatorElm->getAttribute(CComBSTR(VSI_OPERATOR_ATT_STATE), &vState);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		VsiVariantDecodeValue(vState, VT_UI4);
		SetState(static_cast<VSI_OPERATOR_STATE>(V_UI4(&vState)));

		// Get Study Access
		CComVariant vStudyAccess;
		hr = pOperatorElm->getAttribute(CComBSTR(VSI_OPERATOR_ATT_STUDY_ACCESS), &vStudyAccess);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK == hr)
		{
			VsiVariantDecodeValue(vStudyAccess, VT_UI4);

			SetDefaultStudyAccess(static_cast<VSI_OPERATOR_DEFAULT_STUDY_ACCESS>(V_UI4(&vStudyAccess)));
		}
		else
		{
			// Default to public for old operators
			SetDefaultStudyAccess(VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PUBLIC);
		}

		// Get Group info
		CComVariant vGroup;
		hr = pOperatorElm->getAttribute(CComBSTR(VSI_OPERATOR_ATT_GROUP), &vGroup);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK == hr)
		{
			SetGroup(V_BSTR(&vGroup));
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// CVsiOperator::SaveXml
/// </summary>
STDMETHODIMP CVsiOperator::SaveXml(IXMLDOMNode* pXMLNodeRoot)
{
	HRESULT hr = S_OK;
	try
	{
		CComPtr<IXMLDOMDocument> pXmlDoc;

		hr = pXMLNodeRoot->get_ownerDocument(&pXmlDoc);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"SaveXml failed to setAttribute for Operator %s", m_strName.GetString()));

		CComPtr<IXMLDOMElement> pXmlElemOperator;
		hr = pXmlDoc->createElement(CComBSTR(VSI_OPERATOR_ELM_OPERATOR), &pXmlElemOperator);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"SaveXml failed to setAttribute for Operator %s", m_strName.GetString()));

		hr = pXMLNodeRoot->appendChild(pXmlElemOperator, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"SaveXml failed to appendChild for global Operator");

		// Name
		CComVariant vName(m_strName);
		hr = pXmlElemOperator->setAttribute(CComBSTR(VSI_OPERATOR_ATT_NAME), vName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"SaveXml failed to setAttribute for Operator %s", m_strName.GetString()));

		// Unique Id
		CComVariant vId(m_strId);
		hr = pXmlElemOperator->setAttribute(CComBSTR(VSI_OPERATOR_ATT_UNIQUEID), vId);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"SaveXml failed to setAttribute for Operator %s", m_strName.GetString()));

		// Password  
		CComVariant vPassword(m_strPassword);
		hr = pXmlElemOperator->setAttribute(CComBSTR(VSI_OPERATOR_ATT_PASSWORD), vPassword);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"SaveXml failed to setAttribute for Operator %s", m_strName.GetString()));

		// Rights
		CComVariant vType(m_dwType);
		hr = pXmlElemOperator->setAttribute(CComBSTR(VSI_OPERATOR_ATT_TYPE), vType);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"SaveXml failed to setAttribute for Operator %s", m_strName.GetString()));

		// State
		CComVariant vState(VSI_OPERATOR_STATE_DISABLE_SESSION == m_dwState ? VSI_OPERATOR_STATE_ENABLE : m_dwState);
		hr = pXmlElemOperator->setAttribute(CComBSTR(VSI_OPERATOR_ATT_STATE), vState);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"SaveXml failed to setAttribute for Operator %s", m_strName.GetString()));

		// Study access
		CComVariant vStudyAccess(m_dwDefaultStudyAccess);
		hr = pXmlElemOperator->setAttribute(CComBSTR(VSI_OPERATOR_ATT_STUDY_ACCESS), vStudyAccess);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"SaveXml failed to setAttribute for Operator %s", m_strName.GetString()));

		// Group
		CComVariant vGroup(m_strGroup);
		hr = pXmlElemOperator->setAttribute(CComBSTR(VSI_OPERATOR_ATT_GROUP), vGroup);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"SaveXml failed to setAttribute for Operator %s", m_strGroup.GetString()));
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// CVsiOperator::SaveSax
/// </summary>
STDMETHODIMP CVsiOperator::SaveSax(IUnknown* pUnkISAXContentHandler)
{
	HRESULT hr = S_OK;
	CCritSecLock cs(m_cs);

	try
	{
		CComQIPtr<ISAXContentHandler> pch(pUnkISAXContentHandler);
		VSI_CHECK_INTERFACE(pch, VSI_LOG_ERROR, NULL);

		CComPtr<IMXAttributes> pmxa;
		hr = VsiCreateXmlSaxAttributes(&pmxa);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<ISAXAttributes> pa(pmxa);

		WCHAR szNull[] = L"";
		CComBSTR bstrNull(L"");

		// Name
		CComBSTR bstrAtt(VSI_OPERATOR_ATT_NAME);
		CComBSTR bstrValue(m_strName);

		hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, bstrValue);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Unique Id
		bstrAtt = VSI_OPERATOR_ATT_UNIQUEID;
		bstrValue = m_strId;

		hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, bstrValue);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Password  
		bstrAtt = VSI_OPERATOR_ATT_PASSWORD;
		bstrValue = m_strPassword;

		hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, bstrValue);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Rights
		CComVariant vType(m_dwType);
		hr = vType.ChangeType(VT_BSTR);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		bstrAtt = VSI_OPERATOR_ATT_TYPE;

		hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, V_BSTR(&vType));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// State
		CComVariant vState(VSI_OPERATOR_STATE_DISABLE_SESSION == m_dwState ? VSI_OPERATOR_STATE_ENABLE : m_dwState);
		hr = vState.ChangeType(VT_BSTR);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		bstrAtt = VSI_OPERATOR_ATT_STATE;

		hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, V_BSTR(&vState));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Study access
		CComVariant vStudyAccess(m_dwDefaultStudyAccess);
		hr = vStudyAccess.ChangeType(VT_BSTR);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		bstrAtt = VSI_OPERATOR_ATT_STUDY_ACCESS;

		hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, V_BSTR(&vStudyAccess));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Group
		bstrAtt = VSI_OPERATOR_ATT_GROUP;
		bstrValue = m_strGroup;

		hr = pmxa->addAttribute(bstrNull, bstrAtt, bstrAtt, bstrNull, bstrValue);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pch->startElement(szNull, 0, szNull, 0,
			VSI_OPERATOR_ELM_OPERATOR, VSI_STATIC_STRING_COUNT(VSI_OPERATOR_ELM_OPERATOR), pa.p);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pch->endElement(szNull, 0, szNull, 0,
			VSI_OPERATOR_ELM_OPERATOR, VSI_STATIC_STRING_COUNT(VSI_OPERATOR_ELM_OPERATOR));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}
