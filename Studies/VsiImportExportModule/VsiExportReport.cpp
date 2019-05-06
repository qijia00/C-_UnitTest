/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportReport.cpp
**
**	Description:
**		Implementation of CVsiExportReport
**
********************************************************************************/

#include "stdafx.h"
#include <VsiCommUtl.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiStudyModule.h>
#include "VsiCSVWriter.h"
#include "VsiExportReport.h"


CVsiExportReport::CVsiExportReport() :
	m_bCancelCopy(FALSE),
	m_hwndProgress(NULL)
{
}

CVsiExportReport::~CVsiExportReport()
{
}

HRESULT CVsiExportReport::DoExport(LPCWSTR pszPath, const VARIANT *pvParam, IVsiApp *pApp)
{
	HRESULT hr = S_OK;

	try
	{
		_ASSERT(NULL != pvParam);

		CComQIPtr<IXMLDOMDocument> pDoc(V_UNKNOWN(pvParam));
		VSI_CHECK_INTERFACE(pDoc, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiServiceProvider> pServiceProvider(pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Get PDM
		CComPtr<IVsiPdm> pPdm;
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CVsiCSVWriter::VSI_CSV_WRITER_FILTER filter(CVsiCSVWriter::VSI_CSV_WRITER_FILTER_NONE);

		CVsiBitfield<ULONG> pReportOutput;
		hr = pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_SYS_ANALYSIS_REPORT_OUTPUT, &pReportOutput);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (0 != pReportOutput.GetBits(VSI_SYS_ANALYSIS_REPORT_OUTPUT_CALCULATIONS))
		{
			filter |= CVsiCSVWriter::VSI_CSV_WRITER_FILTER_CALCULATIONS;
		}
		if (0 != pReportOutput.GetBits(VSI_SYS_ANALYSIS_REPORT_OUTPUT_MEASUREMENTS))
		{
			filter |= CVsiCSVWriter::VSI_CSV_WRITER_FILTER_MEASUREMENTS;
		}

		CVsiCSVWriter writer;
		hr = writer.ConvertFile(pDoc, pApp, filter, pszPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR | VSI_LOG_NO_BOX, L"Export of the measurement report failed");		

		VsiFlushVolume(pszPath);
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiExportReport::ExportReports(HWND hDlg, const VARIANT *pvParam)
{
	HRESULT hr = S_OK;

	try
	{
		m_pExportProps = (V_UNKNOWN(pvParam));
		VSI_CHECK_INTERFACE(m_pExportProps, VSI_LOG_ERROR, NULL);

		VsiProgressDialogBox(
			NULL, (LPCWSTR)MAKEINTRESOURCE(VSI_PROGRESS_TEMPLATE_MPC),
			hDlg, ExportReportCallback, this);
	}
	VSI_CATCH(hr);

	m_pExportProps.Release();

	return hr;
}

// TODO:
DWORD CVsiExportReport::ThreadProc()
{
	HRESULT hr = S_OK;
	CComVariant vPath;
	BOOL bBackup = FALSE;
	WCHAR szBackupPath[MAX_PATH];
	HANDLE hFile = NULL;

	try
	{
		hr = CoInitializeEx(0, COINIT_MULTITHREADED);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"CoInitializeEx failed");

		CComVariant vStudy;
		hr = m_pExportProps->Read(L"study", &vStudy);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiStudy> pStudy(V_UNKNOWN(&vStudy));
		VSI_CHECK_INTERFACE(pStudy, VSI_LOG_ERROR, NULL);

		CComVariant vStudyId;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vStudyId);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vStudyName;
		hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vStudyName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		if (lstrlen(V_BSTR(&vStudyName)) != 0)
			SetWindowText(m_hwndProgress, V_BSTR(&vStudyName));
		else
			SetWindowText(m_hwndProgress, L"Exporting Report");

		CComQIPtr<IVsiPersistStudy> pPersist(pStudy);
		VSI_CHECK_INTERFACE(pPersist, VSI_LOG_ERROR, NULL);

		hr = pPersist->LoadStudyData(NULL, VSI_DATA_TYPE_SERIES_LIST, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vType;
		hr = m_pExportProps->Read(L"type", &vType);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		CComVariant vMeasurement;
		hr = m_pExportProps->Read(L"measurement", &vMeasurement);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		hr = m_pExportProps->Read(L"path", &vPath);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		if (0 == _waccess(V_BSTR(&vPath), 0))
		{
			swprintf_s(szBackupPath, _countof(szBackupPath), L"%s.%s",
				V_BSTR(&vPath), L"bak");
			if (0 == _waccess(szBackupPath, 0))
			{
				// Delete old backup
				BOOL bDel = DeleteFile(szBackupPath);
				if (!bDel)
				{
					// TODO:
					return 0;
				}
			}
			// Backup
			BOOL bMove = MoveFile(V_BSTR(&vPath), szBackupPath);
			if (!bMove)
			{
				// TODO:
				return 0;
			}

			bBackup = TRUE;
		}

		m_strPath = V_BSTR(&vPath);
		hFile = CreateFile(
			m_strPath,
			GENERIC_WRITE,
			FILE_SHARE_READ,
			NULL,
			OPEN_ALWAYS,
			FILE_ATTRIBUTE_NORMAL,
			NULL);
		if (INVALID_HANDLE_VALUE == hFile)
		{
			// not good, allows continue without the log file
			VSI_LOG_MSG(
				VSI_LOG_WARNING,
				L"Severe system error has occurred. Unable to create CSV file!");
		}

		for (int i = 0; i < 100; ++i)
		{
			SendMessage(m_hwndProgress, WM_VSI_PROGRESS_SETPOS, i, 0);
// TODO:
			Sleep(50);

			if (m_bCancelCopy)
			{
// TODO: rollback

				break;
			}
		}
	}
	VSI_CATCH(hr);

	if (bBackup)
	{
		// Restore from backup
		BOOL bMove = MoveFile(szBackupPath, V_BSTR(&vPath));
		if (!bMove)
		{
			// TODO:
		}
	}

	if (hFile != NULL)
		CloseHandle(hFile);

	// Make sure files are flushed
	VsiFlushVolume(m_strPath);

	PostMessage(m_hwndProgress, WM_CLOSE, 0, 0);

	CoUninitialize();

	return 0;
}

DWORD CALLBACK CVsiExportReport::ExportReportCallback(HWND hWnd, UINT uiState, LPVOID pData)
{
	CVsiExportReport *pHelper = (CVsiExportReport *)pData;

	if (VSI_PROGRESS_STATE_INIT == uiState)
	{
		// cache progress dialog handle
		pHelper->m_hwndProgress = hWnd;
/*
		// set progress bar range to the number of reports we're exporting
		// based on the number of studies times (the number of modes sections
		// with measurements, calculated as log base-2 of k_ModeLastDefined,
		// plus a header section and a multi-step calculations section and
		// a generic measurements section)
		pHelper->m_pStudyInfo->m_iProgressSteps =
			(int)(pHelper->m_pStudyInfo->m_TotalSize) *
				((int)(log10((double)k_Mode_LastDefined)/log10(2.0) + 1.0)
					+ 1 + CALC_PROG_BAR_STEPS + 1);
		SendMessage(hWnd, WM_VSI_PROGRESS_SETMAX, pHelper->m_pStudyInfo->m_iProgressSteps, 0);
*/
		// temp
		SendMessage(hWnd, WM_VSI_PROGRESS_SETMAX, 100, 0);
		SendMessage(hWnd, WM_VSI_PROGRESS_SETMSG, 0, (LPARAM)L"Exporting Report Data...");

		pHelper->RunThread(L"Export Reports");
	}
	else if (VSI_PROGRESS_STATE_CANCEL == uiState)
	{
		pHelper->m_bCancelCopy = TRUE;
	}
	else if (VSI_PROGRESS_STATE_END == uiState)
	{
		// Nothing
	}

	return 0;
}

