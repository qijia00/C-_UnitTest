/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportStudy.h
**
**	Description:
**		Declaration of the CVsiExportReport
**
*******************************************************************************/

#pragma once


#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>

class CVsiExportReport : public CVsiThread<CVsiExportReport>
{
	friend CVsiThread<CVsiExportReport>;

private:
	CComQIPtr<IVsiPropertyBag> m_pExportProps;

	BOOL m_bCancelCopy;  // Set this TRUE to abort the current file copy.
	HWND m_hwndProgress;
	CString m_strPath;

public:
	CVsiExportReport();
	~CVsiExportReport();

	HRESULT DoExport(LPCWSTR pszPath, const VARIANT *pvParam, IVsiApp *pApp);
	HRESULT ExportReports(HWND hDlg, const VARIANT *pvParam);

private:
	DWORD ThreadProc();

	// callback function to handle msmt report exporting
	static DWORD CALLBACK ExportReportCallback(HWND hWnd, UINT uiState, LPVOID pData);
};
