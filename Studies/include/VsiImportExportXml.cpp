/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiImportExportXml.cpp
**
**	Description:
**		Implementation of Import/Export XML helpers
**
*******************************************************************************/

#include "stdafx.h"
#include <ATLComTime.h>
#include <VsiGlobalDef.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiStudyModule.h>
#include "VsiImportExportXml.h"


static HRESULT BuildDefaultDirectoryName(
	IVsiStudy *pStudy, LPWSTR pszBuffer, int iBufferLen)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pStudy, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(pszBuffer, VSI_LOG_ERROR, NULL);

		pszBuffer[0] = 0;

		// Study Name
		CComVariant vStudyName;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vStudyName);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);
		wcsncpy_s(pszBuffer, iBufferLen, V_BSTR(&vStudyName), _TRUNCATE);
		
		// Replace invalid chars
		LPWSTR pszTest = pszBuffer;
		while (0 != *pszTest)
		{
			if (L'\\' == *pszTest ||
				L'/' == *pszTest ||
				L':' == *pszTest ||
				L'*' == *pszTest ||
				L'?' == *pszTest ||
				L'\"' == *pszTest ||
				L'<' == *pszTest ||
				L'>' == *pszTest ||
				L'|' == *pszTest)
			{
				*pszTest = L'-';
			}
			++pszTest;
		}

		// Date
		SYSTEMTIME stDate;
		CComVariant vDate;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_CREATED_DATE, &vDate);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		COleDateTime date(vDate);
		date.GetAsSystemTime(stDate);

		SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

		CString strDate;
		strDate.Format(L"_%04d-%02d-%02d-%02d-%02d-%02d",
			stDate.wYear, stDate.wMonth, stDate.wDay,
			stDate.wHour, stDate.wMinute, stDate.wSecond);

		wcsncat_s(pszBuffer, iBufferLen, strDate, _TRUNCATE);

		// Owner
		CComVariant vOwner;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
		if (S_OK == hr && (0 != *V_BSTR(&vOwner)))
		{
			wcsncat_s(pszBuffer, iBufferLen, L"_", _TRUNCATE);
			wcsncat_s(pszBuffer, iBufferLen, V_BSTR(&vOwner), _TRUNCATE);
		}
	}
	VSI_CATCH(hr);

	return hr;
}


/// <summary>
///		Creates study XML element as a child of the specified parent.
/// </summary>
///
/// <param name= "pStudyDoc"></param>
/// <param name= "pParentElement"></param>
/// <param name= "pStudy"></param>
/// <param name= "ppStudyElement"></param>
///
/// <returns></returns>
HRESULT VsiCreateStudyElement(
	IXMLDOMDocument *pDoc,
	IXMLDOMElement *pParentElement,
	IVsiStudy *pStudy,
	IXMLDOMElement **ppStudyElement,
	LPCWSTR pszNewOwner,
	BOOL bTempTarget,
	BOOL bCalculateSize)
{
	HRESULT hr;
	CComVariant var;
	WCHAR szTempBuffer[256];
	INT64 iTotalSize = 0;
	LPWSTR pszSpt = NULL;
	CComPtr<IXMLDOMElement> pStudyElement;

	try
	{
		VSI_CHECK_ARG(pDoc, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(pParentElement, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(pStudy, VSI_LOG_ERROR, NULL);

		if (NULL != ppStudyElement)
			*ppStudyElement = NULL;

		hr = pDoc->createElement(CComBSTR(VSI_STUDY_XML_ELM_STUDY), &pStudyElement);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Id
		var.Clear();
		hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &var);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_ID), var);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Study name
		var.Clear();
		hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &var);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_NAME), var);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Created date
		SYSTEMTIME stDate;
		CComVariant vDate;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_CREATED_DATE, &vDate);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		COleDateTime date(vDate);
		date.GetAsSystemTime(stDate);

		SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

		GetDateFormat(
			LOCALE_USER_DEFAULT,
			DATE_SHORTDATE,
			&stDate,
			NULL,
			szTempBuffer,
			_countof(szTempBuffer));

		var = szTempBuffer;
		hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_CREATED_DATE), var);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Path
		CComHeapPtr<OLECHAR> pszPath;
		hr = pStudy->GetDataPath(&pszPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		pszSpt = wcsrchr(pszPath, L'\\');
		CString strPath(pszPath, (int)(pszSpt - pszPath));

		var = (LPCWSTR)strPath;
		hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_PATH), var);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Default constructed target folder name
		WCHAR szFolderName[VSI_MAX_PATH];
		if (bTempTarget)
		{
			WCHAR szTmp[VSI_MAX_PATH];
			hr = VsiGetUniqueFileName(L"", L"", L"tmp", szTmp, _countof(szTmp));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			lstrcpy(szFolderName, szTmp + 1);  // Skip '\\'
		}
		else
		{
			hr = BuildDefaultDirectoryName(pStudy, szFolderName, _countof(szFolderName));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		var = szFolderName;
		hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_TARGET), var);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Owner
		var.Clear();
		if (NULL != pszNewOwner)
		{
			var = pszNewOwner;
		}
		else
		{
			hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &var);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);
		}
		hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_OWNER), var);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Size
		if (bCalculateSize)
		{
			iTotalSize = VsiGetDirectorySize(strPath);
			var = iTotalSize;
			hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), var);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Locked
		var.Clear();
		hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &var);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		
		if (S_OK != hr)
			var = false;

		hr = pStudyElement->setAttribute(CComBSTR(VSI_STUDY_XML_ATT_LOCKED),
			(VARIANT_FALSE == V_BOOL(&var)) ? CComVariant(L"false") : CComVariant(L"true"));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParentElement->appendChild(pStudyElement, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (NULL != ppStudyElement)
		{
			*ppStudyElement = pStudyElement.Detach();
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
///		Creates series XML element as a child of the specified parent.
/// </summary>
///
/// <param name= "pStudyDoc"></param>
/// <param name= "pParentElement"></param>
/// <param name= "pSeries"></param>
///
/// <returns></returns>
HRESULT VsiCreateSeriesElement(
	IXMLDOMDocument* pDoc,
	IXMLDOMElement* pParentElement,
	IVsiSeries *pSeries,
	IXMLDOMElement **ppSeriesElement,
	INT64 *piTotalSize)
{
	HRESULT hr;
	CComVariant var;

	try
	{
		if (NULL != ppSeriesElement)
			*ppSeriesElement = NULL;

		if (NULL != piTotalSize)
			*piTotalSize = 0;

		CComPtr<IXMLDOMElement> pSeriesElement;
		hr = pDoc->createElement(CComBSTR(VSI_SERIES_XML_ELM_SERIES), &pSeriesElement);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Set id attribute
		CComVariant vId;
		hr = pSeries->GetProperty(VSI_PROP_SERIES_ID, &vId);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		hr = pSeriesElement->setAttribute(CComBSTR(VSI_SERIES_XML_ATT_ID), vId);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Set folder name attribute
		CComHeapPtr<OLECHAR> pszDataPath;
		hr = pSeries->GetDataPath(&pszDataPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		LPWSTR pszSpt1 = NULL, pszSpt2 = NULL;
		// Get folder name from the path
		pszSpt2 = wcsrchr(pszDataPath, L'\\');
		CString strPath(pszDataPath, (int)(pszSpt2 - pszDataPath));
		pszSpt1 = wcsrchr(strPath.GetBuffer(), L'\\');
		CString strFolderName(pszSpt1 + 1);

		var = (LPCWSTR)strFolderName;
		hr = pSeriesElement->setAttribute(CComBSTR(VSI_SERIES_XML_ATT_FOLDER_NAME), var);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParentElement->appendChild(pSeriesElement, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (NULL != ppSeriesElement)
		{
			*ppSeriesElement = pSeriesElement.Detach();
		}

		if (NULL != piTotalSize)
		{
			*piTotalSize = VsiGetDirectorySize(strPath);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
///	Creates series XML element as a child of the specified parent.
/// </summary>
HRESULT VsiCreateImageElement(
	IXMLDOMDocument* pDoc,
	IXMLDOMElement* pParentElement,
	IVsiImage *pImage,
	IXMLDOMElement **ppImageElement,
	INT64 *piTotalSize)
{
	HRESULT hr;
	CComVariant var;

	try
	{
		if (NULL != ppImageElement)
			*ppImageElement = NULL;

		if (NULL != piTotalSize)
			*piTotalSize = 0;

		CComPtr<IXMLDOMElement> pImageElement;
		hr = pDoc->createElement(CComBSTR(VSI_IMAGE_XML_ELM_IMAGE), &pImageElement);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComHeapPtr<OLECHAR> pszDataPath;
		hr = pImage->GetDataPath(&pszDataPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		WCHAR szImageFileName[MAX_PATH];
		wcscpy_s(szImageFileName, _countof(szImageFileName), PathFindFileName(pszDataPath));
		PathRemoveExtension(szImageFileName);

		var = (LPCWSTR)szImageFileName;
		hr = pImageElement->setAttribute(CComBSTR(VSI_IMAGE_XML_ATT_NAME), var);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		var = (LPCWSTR)pszDataPath;
		hr = pImageElement->setAttribute(CComBSTR(VSI_IMAGE_XML_ATT_PATH), var);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pParentElement->appendChild(pImageElement, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (NULL != ppImageElement)
		{
			*ppImageElement = pImageElement.Detach();
		}

		if (NULL != piTotalSize)
		{
			LARGE_INTEGER liSize = {0};
			INT64 iTotal(0);

			WCHAR szFilter[MAX_PATH];
			WCHAR szFilePath[MAX_PATH];
	
			wcscpy_s(szFilter, _countof(szFilter), pszDataPath);
			PathRenameExtension(szFilter, L".*");

			wcscpy_s(szFilePath, _countof(szFilePath), pszDataPath);
			LPWSTR pszFileName = PathFindFileName(szFilePath);
			int iFileNameMax = _countof(szFilePath) - (int)(pszFileName - szFilePath);

			CVsiFileIterator iter(szFilter);
			
			LPWSTR pszFile(NULL);
			while (NULL != (pszFile = iter.Next()))
			{
				wcscpy_s(pszFileName, iFileNameMax, pszFile);
			
				VsiGetFileSize(szFilePath, &liSize);
				
				iTotal += liSize.QuadPart;
			}

			*piTotalSize = iTotal;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

