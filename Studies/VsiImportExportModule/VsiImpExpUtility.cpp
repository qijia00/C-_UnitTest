/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiImpExpUtility.cpp
**
**	Description:
**		Implementation of CVsiImpExpUtility
**
********************************************************************************/

#include "stdafx.h"
#include <ATLComTime.h>
#include "VsiImpExpUtility.h"
#include "resource.h"


/// <summary>
///		Returns false if this is a invalid characters.
/// </summary>
///
/// <param name="c">The character to be tested.</param>
///
/// <returns>True if the charcater is valid in a Windows file name</returns>
BOOL VsiImpExpUtility::IsValidChar(WCHAR c)
{
	if (c == VK_BACK || IsValidClipboardChar(c))
		return true;

	return false;
}


//
/// <summary>
///		Returns false if this is a invalid characters.
/// </summary>
///
/// <param name="c">The character to be tested.</param>
///
/// <returns>True if the charcater is valid in a Windows file name</returns>
BOOL VsiImpExpUtility::IsValidClipboardChar(WCHAR c)
{
	return (
		c != L'\\' &&
		c != L'/' &&
		c != L':' &&
		c != L'*' &&
		c != L'?' &&
		c != L'\"' &&
		c != L'<' &&
		c != L'>' &&
		c != L'|');
}

/// <summary>
///		Filter a file name to remove illegal characters. The file name is
///		changed in place. Illegal characters are replaced with an underscore
/// </summary>
///
/// <param name="pszFileName">Pointer to the file name to be filtered.</param>
/// <param name="length">The maximum length of the target buffer.</param>
///
/// <returns>Pointer to the filtered file name</returns>
LPWSTR VsiImpExpUtility::FilterFileName(LPWSTR pszFileName, int length)
{
	LPTSTR p = pszFileName;

	while ((*p != NULL) && (--length > 0))
	{
		if (!IsValidChar(*p))
			*p = L'_';

		++p;
	}

	return pszFileName;
}

BOOL VsiImpExpUtility::BuildDefaultDirectoryName(IVsiStudy *pStudy, WCHAR *lpszBuffer, int iBufferLen)
{
	HRESULT hr = S_OK;

	if (NULL == lpszBuffer)
		return FALSE;

	lpszBuffer[0] = NULL;

	try
	{
		// Study Name
		CComVariant vStudyName;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vStudyName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		wcscpy_s(lpszBuffer, iBufferLen, V_BSTR(&vStudyName));

		// Date
		SYSTEMTIME stDate;
		CComVariant vDate;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_CREATED_DATE, &vDate);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		COleDateTime date(vDate);
		date.GetAsSystemTime(stDate);

		SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

		CString strDate;
		strDate.Format(L"%04d-%02d-%02d-%02d-%02d-%02d",
			stDate.wYear, stDate.wMonth, stDate.wDay,
			stDate.wHour, stDate.wMinute, stDate.wSecond);

		wcscat_s(lpszBuffer, iBufferLen, L"_");
		wcscat_s(lpszBuffer, iBufferLen, strDate);

		// Owner
		CComVariant vOwner;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		wcscat_s(lpszBuffer, iBufferLen, L"_");
		wcscat_s(lpszBuffer, iBufferLen, V_BSTR(&vOwner));
	}
	VSI_CATCH(hr);

	return SUCCEEDED(hr);
}
