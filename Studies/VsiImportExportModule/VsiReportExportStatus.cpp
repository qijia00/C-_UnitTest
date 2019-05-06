/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiReportExportStatus.cpp
**
**	Description:
**		Implementation of the CVsiReportExportStatus class.
**
********************************************************************************/

#include "stdafx.h"
#include <VsiImportExportModule.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiImportExportXml.h>
#include <VsiResProduct.h>
#include "resource.h"
#include "VsiImpExpUtility.h"
#include "VsiReportExportStatus.h"


CVsiReportExportStatus::CVsiReportExportStatus(
	IUnknown *pXmlDoc, REPORT_TYPE type)
:
	m_pXmlDoc(pXmlDoc),
	m_type(type)
{
}

CVsiReportExportStatus::~CVsiReportExportStatus()
{
}

/// <summary>
///	Generate the status report from the study export collection passed in
///	during construction. This code assumes that the list is order such that
///	all successful copies are first followed by the one that failed, followed
///	all the ones that weren't copied.
/// </summary>
LRESULT CVsiReportExportStatus::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	// Set font
	HFONT hFont;
	VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
	VsiThemeRecurSetFont(m_hWnd, hFont);

	HIMAGELIST hilHeader(NULL);
	VsiThemeGetImageList(VSI_THEME_IL_HEADER_INFO, &hilHeader);
	VsiThemeDialogBox(*this, TRUE, TRUE, hilHeader);

	CWindow wndListCtrl(GetDlgItem(IDC_REPORT_LIST));
	ListView_SetExtendedListViewStyleEx(
		wndListCtrl,
		LVS_EX_LABELTIP | LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER,
		LVS_EX_LABELTIP | LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);
	UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);
	wndListCtrl.SendMessage(uiThemeCmd, VSI_LV_CMD_NEW_THEME);

	CRect rc;
	wndListCtrl.GetClientRect(&rc);

	LVCOLUMN col;
	memset(&col, 0x00, sizeof(col));
	col.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
	col.fmt = LVCFMT_LEFT;
	col.pszText = L"";
	col.cx = rc.Width() - GetSystemMetrics(SM_CXVSCROLL);
	ListView_InsertColumn(wndListCtrl, 0, &col);

	if (m_pXmlDoc)
	{
		switch (m_type)
		{
		case REPORT_TYPE_STUDY_EXPORT:
			{
				SetWindowText(CString(MAKEINTRESOURCE(IDS_STYREP_COPY_STUDY_REPORT)));

				ReportStudySection(VSI_SEI_COPY_FAILED,
					CString(MAKEINTRESOURCE(IDS_STYREP_STUDIES_FAILED_TO_COPY)));
				ReportStudySection(VSI_SEI_COPY_OK_AUTO_DEL_FAILED,
					CString(MAKEINTRESOURCE(IDS_STYREP_STUDIES_FAILED_TO_DELETE)));
				ReportStudySection(
					VSI_SEI_COPY_OK | VSI_SEI_COPY_OK_AUTO_DEL | VSI_SEI_COPY_OK_AUTO_DEL_FAILED,
					CString(MAKEINTRESOURCE(IDS_STYREP_STUDIES_SUCCESS_COPY)));
				ReportStudySection(VSI_SEI_COPY_OK_AUTO_DEL,
					CString(MAKEINTRESOURCE(IDS_STYREP_STUDIES_WERE_DELETED)));
				ReportStudySection(VSI_SEI_COPY_CANCELED,
					CString(MAKEINTRESOURCE(IDS_STYREP_EXPORT_CANCELLED)));
				ReportStudySection(VSI_SEI_COPY_SKIP,
					CString(MAKEINTRESOURCE(IDS_STYREP_STUDIES_WERE_NOT_COPIED)));
			}
			break;

		case REPORT_TYPE_STUDY_IMPORT:
			{
				SetWindowText(CString(MAKEINTRESOURCE(IDS_STYREP_COPY_STUDY_REPORT)));

				ReportStudySection(VSI_SEI_COPY_FAILED,
					CString(MAKEINTRESOURCE(IDS_STYREP_STUDIES_FAILED_TO_COPY)));					
				ReportStudySection(VSI_SEI_COPY_OK,
					CString(MAKEINTRESOURCE(IDS_STYREP_STUDIES_SUCCESS_COPY)));				
				ReportStudySection(VSI_SEI_COPY_CANCELED,
					CString(MAKEINTRESOURCE(IDS_STYREP_IMPORT_CANCELLED)));					
				ReportStudySection(VSI_SEI_COPY_SKIP,
					CString(MAKEINTRESOURCE(IDS_STYREP_STUDIES_WERE_NOT_COPIED)));
			}
			break;

		case REPORT_TYPE_IMAGE_EXPORT:
			{
				SetWindowText(CString(MAKEINTRESOURCE(IDS_STYREP_IMAGE_EXPORT_REPORT)));

				ReportImageSection(VSI_SEI_EXPORT_FAILED,
					CString(MAKEINTRESOURCE(IDS_STYREP_IMAGES_FAILED_TO_EXPORT)));					
				ReportImageSection(VSI_SEI_EXPORT_CANCELED,
					CString(MAKEINTRESOURCE(IDS_STYREP_IMAGES_WERE_CANCELLED)));					
				ReportImageSection(VSI_SEI_EXPORT_NOTIMPL,
					CString(MAKEINTRESOURCE(IDS_STYREP_IMAGES_WERE_NOT_EXPORTED)));					
				ReportImageSection(VSI_SEI_EXPORT_OK,
					CString(MAKEINTRESOURCE(IDS_STYREP_IMAGES_WERE_EXPORTED)));					
			}
			break;

		case REPORT_TYPE_DICOM_EXPORT:
			{
				SetWindowText(CString(MAKEINTRESOURCE(IDS_STYREP_DICOM_EXPORT_REPORT)));

				ReportStudySection(VSI_SEI_COPY_FAILED,
					CString(MAKEINTRESOURCE(IDS_STYREP_DICOM_FAILED_TO_EXPORT)));
				ReportStudySection(VSI_SEI_COPY_OK,
					CString(MAKEINTRESOURCE(IDS_STYREP_DICOM_FOLLOWING_EXPORTED)));
				ReportStudySection(VSI_SEI_COPY_CANCELED,
					CString(MAKEINTRESOURCE(IDS_STYREP_DICOM_EXPORT_CANCELLED)));
				ReportStudySection(VSI_SEI_COPY_SKIP,
					CString(MAKEINTRESOURCE(IDS_STYREP_DICOM_WERE_NOT_EXPORTED)));
			}
			break;
			
		default:
			_ASSERT(0);
		}
		
		if (0 == ListView_GetItemCount(wndListCtrl))
		{
			CString strMsg;
		
			switch (m_type)
			{
			case REPORT_TYPE_STUDY_EXPORT:
				strMsg.LoadString(IDS_STYREP_NONE_STUDIES_EXPORTED);
				break;
			case REPORT_TYPE_STUDY_IMPORT:
				strMsg.LoadString(IDS_STYREP_NONE_STUDIES_IMPORTED);
				break;
			case REPORT_TYPE_DICOM_EXPORT:
				strMsg.LoadString(IDS_STYREP_NONE_IMAGES_EXPORTED);
				break;
			}

			LVITEM item = { 0 };
			item.mask = LVIF_TEXT;
			item.pszText = (LPWSTR)(LPCWSTR)strMsg;
			ListView_InsertItem(wndListCtrl, &item);
		}
	}

	return 0;
}

LRESULT CVsiReportExportStatus::OnClose(UINT, WPARAM, LPARAM, BOOL&)
{
	EndDialog(IDCANCEL);

	return 0;
}

LRESULT CVsiReportExportStatus::OnCmdOK(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Export Report - OK");

	EndDialog(IDOK);

	return 0;
}

/// <summary>
///	Add an export study object to the list box. The new entry is added at the
///	end of the list and no sorting is assumed
/// </summary>
///
/// <param name="pExportInfo">Pointer to the structure containing the information to be added.</param>
///
/// <returns>void</returns>
void CVsiReportExportStatus::AddStudyItem(IXMLDOMElement *pStudyElement)
{
	CWindow wndListCtrl(GetDlgItem(IDC_REPORT_LIST));

	CString strBuffer(L"    ");

	HRESULT hr = S_OK;
	CComVariant vStudyName;
	CComVariant vOwner;
	CComVariant vCreatedDate;
	CComVariant vFailureCode;
	CComVariant vRollbackCode;

	try
	{
		hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_NAME), &vStudyName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_NAME));

		// We do not have to check for error codes as these might not be there by design
		hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_OWNER), &vOwner);
		hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_CREATED_DATE), &vCreatedDate);
		

		hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_FAILURE_CODE), &vFailureCode);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_FAILURE_CODE));

		if (S_OK == hr)
		{
			hr = vFailureCode.ChangeType(VT_UI4);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_FAILURE_CODE));
		}

		// If we don't get rollback attrib, it succeeded
		pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_ROLLBACK_STATUS), &vRollbackCode);
	}
	VSI_CATCH(hr);

	if (FAILED(hr))
		return;

	// Made text and order match the columns in the study browser dialog
	strBuffer += V_BSTR(&vStudyName);

	if (*V_BSTR(&vOwner) != NULL)
	{
		strBuffer += L", Owner: ";
		strBuffer += V_BSTR(&vOwner);
	}

	if (*V_BSTR(&vCreatedDate) != NULL)
	{
		strBuffer += L", Date: ";
		strBuffer += V_BSTR(&vCreatedDate);
	}

	// Report the specific cause of the failure to copy
	if (VT_NULL != V_VT(&vFailureCode))
	{
		UINT uiFailureCode = V_UI4(&vFailureCode);

		if (VSI_SEI_FAILED_NO_FAILURE != uiFailureCode)
		{
			switch (uiFailureCode)
			{
			case VSI_SEI_FAILED_TARGET_LOCKED:
				strBuffer += CString(MAKEINTRESOURCE(IDS_STYREP_LOCKED));
				break;
			case VSI_SEI_FAILED_SOURCE_LOCKED:
				strBuffer += CString(MAKEINTRESOURCE(IDS_STYREP_SOURCE_STUDY_LOCKED));
				break;
			case VSI_SEI_FAILED_DISK_ERROR:
				strBuffer += CString(MAKEINTRESOURCE(IDS_STYREP_DISK_ERROR));
				break;
			case VSI_SEI_FAILED_ACCESS_DENIED:
				strBuffer += CString(MAKEINTRESOURCE(IDS_STYREP_ACCESS_DENIED));
				break;
			case VSI_SEI_FAILED_OVERWRITE_DISABLED:
				strBuffer += CString(MAKEINTRESOURCE(IDS_STYREP_STUDY_OWERWRITE_DISABLED));
				break;
			case VSI_SEI_FAILED_NO_ELIGIBLE_IMAGES_FOR_DICOM:
				strBuffer += CString(MAKEINTRESOURCE(IDS_STYREP_STUDY_NO_DICOM_EXPORTED));
				break;
			}
		}
	}

	LVITEM item = { 0 };
	item.mask = LVIF_TEXT;
	item.iItem = ListView_GetItemCount(wndListCtrl);
	item.pszText = (LPWSTR)(LPCWSTR)strBuffer;
	ListView_InsertItem(wndListCtrl, &item);

	if (VT_NULL != V_VT(&vRollbackCode) && *V_BSTR(&vRollbackCode) != NULL)
	{
		CString strRollback(L"    ");
		strRollback += V_BSTR(&vRollbackCode);
		item.mask = LVIF_TEXT;
		item.iItem = ListView_GetItemCount(wndListCtrl);
		item.pszText = (LPWSTR)(LPCWSTR)strRollback;
		ListView_InsertItem(wndListCtrl, &item);
	}
}

/// <summary>
///	Fill in one section of the status report.
/// </summary>
///
/// <param name="nItem">Reference to the item number in the list control</param>
/// <param name="copyCode">One of the defined status codes.</param>
///
/// <returns>The number of study items added to the list control.</returns>
int CVsiReportExportStatus::ReportStudySection(DWORD dwCode, LPCWSTR pszTitle)
{
	HRESULT hr = S_OK;
	bool bTitleAdded(false);
	HWND hListCtrl = GetDlgItem(IDC_REPORT_LIST);
	int iAdded(0);
	CComVariant vStatus;

	try
	{
		CComPtr<IXMLDOMNodeList> pListStudy;
		hr = m_pXmlDoc->selectNodes(
			CComBSTR(L"//" VSI_STUDY_XML_ELM_STUDIES L"/" VSI_STUDY_XML_ELM_STUDY),
			&pListStudy);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting studies");

		long lLength = 0;
		hr = pListStudy->get_length(&lLength);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
		for (long l = 0; l < lLength; l++)
		{
			CComPtr<IXMLDOMNode> pNode;
			hr = pListStudy->get_item(l, &pNode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComQIPtr<IXMLDOMElement> pStudyElement(pNode);

			vStatus.Clear();
			hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS), &vStatus);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));

			if (S_OK == hr)
			{
				hr = vStatus.ChangeType(VT_UI4);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (dwCode & V_UI4(&vStatus))
				{
					if (!bTitleAdded)
					{
						LVITEM item = { 0 };
						item.mask = LVIF_TEXT;
						item.iItem = ListView_GetItemCount(hListCtrl);
						item.pszText = (LPWSTR)pszTitle;
						ListView_InsertItem(hListCtrl, &item);

						bTitleAdded = true;
					}

					AddStudyItem(pStudyElement);

					++iAdded;
				}
			}
		}

		if (iAdded > 0)
		{
			LVITEM item = { 0 };
			item.mask = LVIF_TEXT;
			item.iItem = ListView_GetItemCount(hListCtrl);
			ListView_InsertItem(hListCtrl, &item);
		}
	}
	VSI_CATCH(hr);

	return iAdded;
}

void CVsiReportExportStatus::AddImageItem(IXMLDOMElement *pImageElement)
{
	CWindow wndListCtrl(GetDlgItem(IDC_REPORT_LIST));

	CString strBuffer(L"    ");

	HRESULT hr = S_OK;
	CComVariant vName, vFailureCode, vDate, vModeName;

	try
	{
		hr = pImageElement->getAttribute(CComBSTR(VSI_IMAGE_XML_ATT_NAME), &vName);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"Failure getting attribute (%s)", VSI_IMAGE_XML_ATT_NAME));

		//we do not have to check for error codes as these might not be there by design
		hr = pImageElement->getAttribute(CComBSTR(VSI_IMAGE_XML_ATT_CREATED_DATE), &vDate);
		hr = pImageElement->getAttribute(CComBSTR(VSI_IMAGE_XML_ATT_MODE_NAME), &vModeName);

		hr = pImageElement->getAttribute(CComBSTR(VSI_IMAGE_XML_ATT_FAILURE_CODE), &vFailureCode);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			VsiFormatMsg(L"Failure getting attribute (%s)", VSI_IMAGE_XML_ATT_FAILURE_CODE));

		if (S_OK == hr)
		{
			hr = vFailureCode.ChangeType(VT_UI4);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, VsiFormatMsg(L"Failure getting attribute (%s)", VSI_IMAGE_XML_ATT_FAILURE_CODE));
		}
	}
	VSI_CATCH(hr);

	if (FAILED(hr))
		return;

	strBuffer += V_BSTR(&vName);
	if (V_VT(&vDate) != VT_NULL)
	{
		strBuffer += L", created on ";
		strBuffer += V_BSTR(&vDate);
	}
	if (V_VT(&vModeName) != VT_NULL)
	{
		strBuffer += L", ";
		strBuffer += V_BSTR(&vModeName);
	}

	// Report the specific cause of the failure to copy
	if (VT_NULL != V_VT(&vFailureCode))
	{
		UINT uiFailureCode = V_UI4(&vFailureCode);

		if (VSI_SEI_FAILED_NO_FAILURE != uiFailureCode)
		{
			switch (uiFailureCode)
			{
			case VSI_SEI_FAILED_ERROR:
				break;
			case VSI_SEI_FAILED_FILE_TOO_LARGE:
				strBuffer += L" - File too large";
				break;
			case VSI_SEI_FAILED_CREATE_BACKUP:
				break;
			case VSI_SEI_NOTIMPL_TODICOM:
				strBuffer += L", export to DICOM";
				break;
			case VSI_SEI_NOTIMPL_TOCINELOOP:
				strBuffer += L", export to Cine Loop";
				break;
			case VSI_SEI_FAILED_NOT_ENOUGH_SPACE:
				strBuffer += L", There is not enough space on the disk";
				break;
			}
		}
	}

	LVITEM item = { 0 };
	item.mask = LVIF_TEXT;
	item.iItem = ListView_GetItemCount(wndListCtrl);
	item.pszText = (LPWSTR)(LPCWSTR)strBuffer;
	ListView_InsertItem(wndListCtrl, &item);
}

int CVsiReportExportStatus::ReportImageSection(DWORD dwCode, LPCWSTR pszTitle)
{
	HRESULT hr = S_OK;
	bool bTitleAdded(false);
	HWND hListCtrl = GetDlgItem(IDC_REPORT_LIST);
	int iAdded(0);
	CComVariant vStatus;

	try
	{
		CComPtr<IXMLDOMNodeList> pListImage;
		hr = m_pXmlDoc->selectNodes(
			CComBSTR(L"//" VSI_IMAGE_XML_ELM_IMAGES L"/" VSI_IMAGE_XML_ELM_IMAGE),
			&pListImage);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting images");

		long lLength = 0;
		hr = pListImage->get_length(&lLength);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
		for (long l = 0; l < lLength; l++)
		{
			CComPtr<IXMLDOMNode> pNode;
			hr = pListImage->get_item(l, &pNode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComQIPtr<IXMLDOMElement> pImageElement(pNode);

			vStatus.Clear();
			hr = pImageElement->getAttribute(CComBSTR(VSI_IMAGE_XML_ATT_EXPORT_STATUS), &vStatus);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure getting attribute (%s)", VSI_IMAGE_XML_ATT_EXPORT_STATUS));

			hr = vStatus.ChangeType(VT_UI4);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (dwCode & V_UI4(&vStatus))
			{
				if (!bTitleAdded)
				{
					LVITEM item = { 0 };
					item.mask = LVIF_TEXT;
					item.iItem = ListView_GetItemCount(hListCtrl);
					item.pszText = (LPWSTR)pszTitle;
					ListView_InsertItem(hListCtrl, &item);

					bTitleAdded = true;
				}

				AddImageItem(pImageElement);

				++iAdded;
			}
		}

		if (iAdded > 0)
		{
			LVITEM item = { 0 };
			item.mask = LVIF_TEXT;
			item.iItem = ListView_GetItemCount(hListCtrl);
			ListView_InsertItem(hListCtrl, &item);
		}
	}
	VSI_CATCH(hr);

	return iAdded;
}
