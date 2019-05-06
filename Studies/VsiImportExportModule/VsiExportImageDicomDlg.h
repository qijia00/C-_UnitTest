/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportImageDicomDlg.h
**
**	Description:
**		Declaration of the CVsiExportImageDicomDlg
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiPdmModule.h>
#include "VsiExportImage.h"


class CVsiExportImageDicomDlg :
	public CDialogImpl<CVsiExportImageDicomDlg>
{
private:
	CVsiExportImage *m_pExportImage;
	CComPtr<IVsiPdm> m_pPdm;

public:
	enum { IDD = IDD_EXPORT_IMAGE_VIEW_DICOM };

public:
	CVsiExportImageDicomDlg(CVsiExportImage *pExportImage) :
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
				VSI_PARAMETER_SYS_DEFAULT_EXPORT_DICOM, m_pPdm);

			BOOL bHideMsmnts = BST_CHECKED == IsDlgButtonChecked(IDC_IE_HIDE_MEASUREMENTS);
			VsiSetBooleanValue<BOOL>(bHideMsmnts,
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_MSMNT_EXPORT_HIDE, m_pPdm);

			BOOL bEncodeRegions = BST_CHECKED == IsDlgButtonChecked(IDC_IE_REGIONS);
			VsiSetBooleanValue<BOOL>(bEncodeRegions,
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_DICOM_ENCODE_REGIONS, m_pPdm);
		}
		VSI_CATCH(hr);

		return hr;
	}

	BOOL UpdateUI()
	{
		CheckDlgButton(IDC_IE_HIDE_MEASUREMENTS, m_pExportImage->m_bHideMsmnts ? BST_CHECKED : BST_UNCHECKED);

		return TRUE;
	}

protected:
	BEGIN_MSG_MAP(CVsiExportImageDicomDlg)
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
				DWORD dwExportTypeDicom = VsiGetRangeValue<DWORD>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
					VSI_PARAMETER_SYS_DEFAULT_EXPORT_DICOM, m_pPdm);

				CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));
				int iDefaultiIndex(0);
				int iIndex;

				iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL,
					(LPARAM)L"Implicit VR Little Endian (*.dcm)");
				wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_DICOM_IMPLICIT_VR_LITTLE_ENDIAN);
				if (VSI_DICOM_IMPLICIT_VR_LITTLE_ENDIAN == dwExportTypeDicom)
				{
					iDefaultiIndex = iIndex;
				}

				iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL,
					(LPARAM)L"Explicit VR Little Endian (*.dcm)");
				wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_DICOM_EXPLICIT_VR_LITTLE_ENDIAN);
				if (VSI_DICOM_EXPLICIT_VR_LITTLE_ENDIAN == dwExportTypeDicom)
				{
					iDefaultiIndex = iIndex;
				}

				if (! m_pExportImage->IsAllLoopsOf3dModes())
				{
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL,
						(LPARAM)L"JPEG Baseline (*.dcm)");
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_DICOM_JPEG_BASELINE);
					if (VSI_DICOM_JPEG_BASELINE == dwExportTypeDicom)
					{
						iDefaultiIndex = iIndex;
					}

					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL,
						(LPARAM)L"RLE Lossless (*.dcm)");
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_DICOM_RLE_LOSSLESS);
					if (VSI_DICOM_RLE_LOSSLESS == dwExportTypeDicom)
					{
						iDefaultiIndex = iIndex;
					}

					// Only show JPEG Lossless if Eng Mode enabled
					if (VsiModeUtils::GetIsEngineerMode(m_pPdm))
					{
						iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL,
							(LPARAM)L"JPEG Lossless (*.dcm)");
						wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_DICOM_JPEG_LOSSLESS);
						if (VSI_DICOM_JPEG_LOSSLESS == dwExportTypeDicom)
						{
							iDefaultiIndex = iIndex;
						}
					}
				}

				// Select a default file type
				wndExportType.SendMessage(CB_SETCURSEL, iDefaultiIndex, NULL);
			}

			BOOL bDicomEncodeRegions = VsiGetBooleanValue<BOOL>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_DICOM_ENCODE_REGIONS, m_pPdm);

			CheckDlgButton(IDC_IE_REGIONS, bDicomEncodeRegions ? BST_CHECKED : BST_UNCHECKED);

			SendDlgItemMessage(IDC_IE_3D_DICOM_INFO, RegisterWindowMessage(VSI_THEME_COMMAND), 
				VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_DYELLOW);

			if (VsiModeUtils::GetIsEngineerMode(m_pPdm))
			{
				GetDlgItem(IDC_IE_3D_DICOM_INFO).ShowWindow(m_pExportImage->m_i3Ds > 0);
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