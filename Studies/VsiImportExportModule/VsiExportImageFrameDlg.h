/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportImageFrameDlg.h
**
**	Description:
**		Declaration of the CVsiExportImageFrameDlg
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiPdmModule.h>
#include "VsiExportImage.h"


class CVsiExportImageFrameDlg :
	public CDialogImpl<CVsiExportImageFrameDlg>
{
private:
	CVsiExportImage *m_pExportImage;
	CComPtr<IVsiPdm> m_pPdm;

public:
	enum { IDD = IDD_EXPORT_IMAGE_VIEW_FRAME };

public:
	CVsiExportImageFrameDlg(CVsiExportImage *pExportImage) :
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
			CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));
			int iIndex = (int)wndExportType.SendMessage(CB_GETCURSEL, NULL, NULL);
			DWORD dwType = (DWORD)wndExportType.SendMessage(CB_GETITEMDATA, iIndex, NULL);

			VsiSetRangeValue<DWORD>(dwType,
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_DEFAULT_EXPORT_FRAME, m_pPdm);

			BOOL bHideMsmnts = BST_CHECKED == IsDlgButtonChecked(IDC_IE_HIDE_MEASUREMENTS);
			VsiSetBooleanValue<BOOL>(bHideMsmnts,
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_MSMNT_EXPORT_HIDE, m_pPdm);
		}
		VSI_CATCH(hr);

		return hr;
	}

	BOOL UpdateUI()
	{
		CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));
		int iIndex = (int)wndExportType.SendMessage(CB_GETCURSEL, NULL, NULL);
		DWORD dwType = (DWORD)wndExportType.SendMessage(CB_GETITEMDATA, iIndex, NULL);

		BOOL bShowExportMsmntOption = FALSE;
		
		switch (dwType)
		{
		case VSI_IMAGE_TIFF_FULL:
		case VSI_IMAGE_BMP_FULL:
		case VSI_IMAGE_TIFF_IMAGE:
		case VSI_IMAGE_BMP_IMAGE:
			bShowExportMsmntOption = TRUE;
			break;
		}

		GetDlgItem(IDC_IE_HIDE_MEASUREMENTS).ShowWindow(bShowExportMsmntOption ? SW_SHOWNA : SW_HIDE);
		if (bShowExportMsmntOption)
		{
			CheckDlgButton(IDC_IE_HIDE_MEASUREMENTS, m_pExportImage->m_bHideMsmnts ? BST_CHECKED : BST_UNCHECKED);
		}

		return TRUE;
	}

protected:
	BEGIN_MSG_MAP(CVsiExportImageFrameDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		COMMAND_HANDLER(IDC_IE_SAVEAS, EN_CHANGE, OnFileNameChanged)
		COMMAND_HANDLER(IDC_IE_TYPE, CBN_SELCHANGE, OnFileTypeChanged)
		COMMAND_HANDLER(IDC_IE_HIDE_MEASUREMENTS, BN_CLICKED, OnBnClickedExportMeasurement)
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

			if (m_pExportImage->m_listImage.size() == 1)
			{
				SetDlgItemText(IDC_IE_FILE_LABEL, L"Save As");
			}

			// Fill types
			{
				DWORD dwExportTypeImage = VsiGetRangeValue<DWORD>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
					VSI_PARAMETER_SYS_DEFAULT_EXPORT_FRAME, m_pPdm);

				CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));
				int iDefaultiIndex(0);
				int iIndex;

				if (m_pExportImage->m_i3Ds > 0)
				{
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"TIFF for 3D Volume Slices (*.tif)");
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_TIFF_LOOP);
					if (VSI_IMAGE_TIFF_LOOP == dwExportTypeImage)
						iDefaultiIndex = iIndex;
				}
				iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"TIFF Full Screen (*.tif)");
				wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_TIFF_FULL);
				if (VSI_IMAGE_TIFF_FULL == dwExportTypeImage)
				{
					iDefaultiIndex = iIndex;
				}

				iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"BMP Full Screen (*.bmp)");
				wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_BMP_FULL);
				if (VSI_IMAGE_BMP_FULL == dwExportTypeImage)
				{
					iDefaultiIndex = iIndex;
				}

				iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"TIFF Image Area (*.tif)");
				wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_TIFF_IMAGE);
				if (VSI_IMAGE_TIFF_IMAGE == dwExportTypeImage)
				{
					iDefaultiIndex = iIndex;
				}

				iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"BMP Image Area (*.bmp)");
				wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_BMP_IMAGE);
				if (VSI_IMAGE_BMP_IMAGE == dwExportTypeImage)
				{
					iDefaultiIndex = iIndex;
				}

				// Add RF if we have RF images
				if (m_pExportImage->IsAllOfRf(VSI_RF_EXPORT_IQ))
				{
					CString strLabel(L"IQ Data (*." VSI_RF_FILE_EXTENSION_IQXML L")");
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)(LPCWSTR)strLabel);
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_IQ_DATA_FRAME);
			
					if (VSI_IMAGE_IQ_DATA_FRAME == dwExportTypeImage)
					{
						iDefaultiIndex = iIndex;
					}
				}
				if (m_pExportImage->IsAllOfRf(VSI_RF_EXPORT_RF))
				{
					CString strLabel(L"RF Data (*." VSI_RF_FILE_EXTENSION_RFXML L")");
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)(LPCWSTR)strLabel);
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_RF_DATA_FRAME);
			
					if (VSI_IMAGE_RF_DATA_FRAME == dwExportTypeImage)
					{
						iDefaultiIndex = iIndex;
					}
				}
				if (m_pExportImage->IsAllOfRf(VSI_RF_EXPORT_RAW))
				{
					CString strLabel(L"RAW Data (*." VSI_RF_FILE_EXTENSION_RAWXML L")");
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)(LPCWSTR)strLabel);
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_RAW_DATA_FRAME);
			
					if (VSI_IMAGE_RAW_DATA_FRAME == dwExportTypeImage)
					{
						iDefaultiIndex = iIndex;
					}
				}

				// Select a default file type
				wndExportType.SendMessage(CB_SETCURSEL, iDefaultiIndex, NULL);
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
	LRESULT OnFileTypeChanged(WORD, WORD, HWND, BOOL&)
	{
		m_pExportImage->UpdateRequiredSize();

		return 0;
	}
	LRESULT OnBnClickedExportMeasurement(WORD, WORD, HWND, BOOL&)
	{
		m_pExportImage->m_bHideMsmnts = BST_CHECKED == IsDlgButtonChecked(IDC_IE_HIDE_MEASUREMENTS);

		return 0;
	}
};