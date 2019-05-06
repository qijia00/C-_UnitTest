/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportImageReportDlg.h
**
**	Description:
**		Declaration of the CVsiExportImageReportDlg
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiPdmModule.h>
#include "VsiExportImage.h"


class CVsiExportImageReportDlg :
	public CDialogImpl<CVsiExportImageReportDlg>
{
private:
	CVsiExportImage *m_pExportImage;
	CComPtr<IVsiPdm> m_pPdm;

public:
	enum { IDD = IDD_EXPORT_IMAGE_VIEW_REPORT };

public:
	CVsiExportImageReportDlg(CVsiExportImage *pExportImage) :
		m_pExportImage(pExportImage)
	{
	}

	void Initialize(IVsiPdm *pPdm)
	{
		m_pPdm = pPdm;
	}

	HRESULT SaveSettings()
	{
		HRESULT hr = S_OK;

		try
		{
			CVsiBitfield<ULONG> pReportOutput;
			hr = m_pPdm->GetParameter(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_ANALYSIS_REPORT_OUTPUT, &pReportOutput);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			ULONG ulSettings(VSI_SYS_ANALYSIS_REPORT_OUTPUT_NONE);

			if (BST_CHECKED == IsDlgButtonChecked(IDC_IE_REPORT_MEASUREMENTS))
			{
				ulSettings |= VSI_SYS_ANALYSIS_REPORT_OUTPUT_MEASUREMENTS;
			}
			if (BST_CHECKED == IsDlgButtonChecked(IDC_IE_REPORT_CALCULATIONS))
			{
				ulSettings |= VSI_SYS_ANALYSIS_REPORT_OUTPUT_CALCULATIONS;
			}

			pReportOutput.SetValue(ulSettings);
		}
		VSI_CATCH(hr);

		return hr;
	}

	BOOL UpdateUI()
	{
		return TRUE;
	}

protected:
	BEGIN_MSG_MAP(CVsiExportImageReportDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		COMMAND_HANDLER(IDC_IE_SAVEAS, EN_CHANGE, OnFileNameChanged)
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
	{
		HRESULT hr = S_OK;

		try
		{
			// Set font
			HFONT hFont;
			VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
			VsiThemeRecurSetFont(m_hWnd, hFont);

			// Fill types
			{
				DWORD dwExportTypeReport = VsiGetRangeValue<DWORD>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
					VSI_PARAMETER_SYS_DEFAULT_EXPORT_REPORT, m_pPdm);

				CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));
				int iDefaultiIndex(0);
				int iIndex;

				iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"VSI Report File (*.csv)");
				wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_COMBINED_REPORT);
				if (VSI_COMBINED_REPORT == dwExportTypeReport)
				{
					iDefaultiIndex = iIndex;
				}

				// Select a default file type
				wndExportType.SendMessage(CB_SETCURSEL, iDefaultiIndex, NULL);
			}

			// Options
			{
				CVsiBitfield<ULONG> pReportOutput;
				hr = m_pPdm->GetParameter(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
					VSI_PARAMETER_SYS_ANALYSIS_REPORT_OUTPUT, &pReportOutput);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (0 != pReportOutput.GetBits(VSI_SYS_ANALYSIS_REPORT_OUTPUT_MEASUREMENTS))
				{
					CheckDlgButton(IDC_IE_REPORT_MEASUREMENTS, BST_CHECKED);
				}
				if (0 != pReportOutput.GetBits(VSI_SYS_ANALYSIS_REPORT_OUTPUT_CALCULATIONS))
				{
					CheckDlgButton(IDC_IE_REPORT_CALCULATIONS, BST_CHECKED);
				}
			}
		}
		VSI_CATCH(hr);

		return 0;
	}
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
	{
		return 0;
	}
	LRESULT OnFileNameChanged(WORD, WORD, HWND, BOOL&)
	{
		m_pExportImage->UpdateUI();

		return 0;
	}
};