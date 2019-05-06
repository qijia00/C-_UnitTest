/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportImageCineDlg.h
**
**	Description:
**		Declaration of the CVsiExportImageCineDlg
**
*******************************************************************************/

#pragma once

#include <vfw.h>
#include <VsiRes.h>
#include <VsiCommUtlLib.h>
#include <VsiPdmModule.h>
#include "VsiExportImage.h"


class CVsiExportImageCineDlg :
	public CDialogImpl<CVsiExportImageCineDlg>
{
private:
	CVsiExportImage *m_pExportImage;
	CComPtr<IVsiPdm> m_pPdm;

public:
	// List of FCC handler types for each AVI format
	DWORD m_pdwFccHandlerType[512];

	enum { IDD = IDD_EXPORT_IMAGE_VIEW_CINE };

public:
	CVsiExportImageCineDlg(CVsiExportImage *pExportImage) :
		m_pExportImage(pExportImage)
	{
		ZeroMemory(m_pdwFccHandlerType, sizeof(m_pdwFccHandlerType));
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
				VSI_PARAMETER_SYS_DEFAULT_EXPORT_LOOP, m_pPdm);

			if (dwType != VSI_IMAGE_IQ_DATA_LOOP)
			{
				// Don't modify FCC handler if we are in RF as its not required
				DWORD dwFccHandler = m_pdwFccHandlerType[iIndex];
				VsiSetRangeValue<DWORD>(dwFccHandler,
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
					VSI_PARAMETER_SYS_DEFAULT_EXPORT_LOOP_HANDLER, m_pPdm);

				BOOL bHideMsmnts = BST_CHECKED == IsDlgButtonChecked(IDC_IE_HIDE_MEASUREMENTS);
				VsiSetBooleanValue<BOOL>(bHideMsmnts,
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
					VSI_PARAMETER_MSMNT_EXPORT_HIDE, m_pPdm);
			}

			// AVI quality
			if (dwType != VSI_IMAGE_IQ_DATA_LOOP &&
				dwType != VSI_IMAGE_IQ_DATA_FRAME)
			{
				DWORD dwAviQuality = VSI_SYS_EXPORT_AVI_QUALITY_HIGH;
				if (BST_CHECKED == IsDlgButtonChecked(IDC_IE_QUALITY_HIGH))
				{
					dwAviQuality = VSI_SYS_EXPORT_AVI_QUALITY_HIGH;
				}
				else if (BST_CHECKED == IsDlgButtonChecked(IDC_IE_QUALITY_MED))
				{
					dwAviQuality = VSI_SYS_EXPORT_AVI_QUALITY_MED;
				}
				VsiSetDiscreteValue<DWORD>(
					dwAviQuality,
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_SYS_EXPORT_AVI_QUALITY, m_pPdm);
			}
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
		case VSI_IMAGE_AVI_UNCOMP:
		case VSI_IMAGE_AVI_VIDEO1:
		case VSI_IMAGE_AVI_MSMEDIAVIDEO9:
		case VSI_IMAGE_AVI_CUSTOM:
			bShowExportMsmntOption = TRUE;
			break;
		}

		GetDlgItem(IDC_IE_HIDE_MEASUREMENTS).ShowWindow(bShowExportMsmntOption ? SW_SHOWNA : SW_HIDE);
		if (bShowExportMsmntOption)
		{
			CheckDlgButton(IDC_IE_HIDE_MEASUREMENTS, m_pExportImage->m_bHideMsmnts ? BST_CHECKED : BST_UNCHECKED);
		}

		// Show quality settings
		BOOL bAviQuality = VSI_IMAGE_WAV != dwType;

		// Do not show quality if we are exporting RF data
		if (dwType == VSI_IMAGE_IQ_DATA_LOOP ||
			dwType == VSI_IMAGE_RF_DATA_LOOP ||
			dwType == VSI_IMAGE_RAW_DATA_LOOP)
		{
			bAviQuality = FALSE;
		}

		// Do not show quality if we exporting to TIFF
		if (dwType == VSI_IMAGE_TIFF_LOOP)
		{
			bAviQuality = FALSE;
		}

		if (bAviQuality)
		{
			GetDlgItem(IDC_IE_QUALITY_HIGH).ShowWindow(SW_SHOW);
			GetDlgItem(IDC_IE_QUALITY_MED).ShowWindow(SW_SHOW);
			GetDlgItem(IDC_IE_QUALITY).ShowWindow(SW_SHOW);
		}
		else
		{
			GetDlgItem(IDC_IE_QUALITY_HIGH).ShowWindow(SW_HIDE);
			GetDlgItem(IDC_IE_QUALITY_MED).ShowWindow(SW_HIDE);
			GetDlgItem(IDC_IE_QUALITY).ShowWindow(SW_HIDE);
		}

		// Check to see if we are still using cine
		if (VsiModeUtils::GetIsEngineerMode(m_pPdm))
		{
			if (!bAviQuality)
			{
				GetDlgItem(IDC_CONFIG_CODEC_BUTTON).ShowWindow(SW_HIDE);
				GetDlgItem(IDC_ABOUT_CODEC_BUTTON).ShowWindow(SW_HIDE);
			}
			else
			{
				CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));
				int iIndex = (int)wndExportType.SendMessage(CB_GETCURSEL, NULL, NULL);
				if (iIndex >= 0 && iIndex < _countof(m_pdwFccHandlerType))
				{
					DWORD dwFccHandler = m_pdwFccHandlerType[iIndex];
					BOOL bAbout = VsiDoesAviCodecHaveAbout(dwFccHandler);
					BOOL bConfig = VsiDoesAviCodecHaveConfig(dwFccHandler);

					GetDlgItem(IDC_CONFIG_CODEC_BUTTON).EnableWindow(bConfig);
					GetDlgItem(IDC_ABOUT_CODEC_BUTTON).EnableWindow(bAbout);
				}

				GetDlgItem(IDC_CONFIG_CODEC_BUTTON).ShowWindow(SW_SHOW);
				GetDlgItem(IDC_ABOUT_CODEC_BUTTON).ShowWindow(SW_SHOW);
			}
		}
		else
		{
			GetDlgItem(IDC_CONFIG_CODEC_BUTTON).ShowWindow(SW_HIDE);
			GetDlgItem(IDC_ABOUT_CODEC_BUTTON).ShowWindow(SW_HIDE);
		}

		return TRUE;
	}

protected:
	BEGIN_MSG_MAP(CVsiExportImageCineDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		COMMAND_HANDLER(IDC_IE_SAVEAS, EN_CHANGE, OnFileNameChanged)
		COMMAND_HANDLER(IDC_IE_TYPE, CBN_SELCHANGE, OnFileTypeChanged)
		COMMAND_HANDLER(IDC_IE_HIDE_MEASUREMENTS, BN_CLICKED, OnBnClickedExportMeasurement)
		COMMAND_HANDLER(IDC_CONFIG_CODEC_BUTTON, BN_CLICKED, OnBnClickedConfigCodec)
		COMMAND_HANDLER(IDC_ABOUT_CODEC_BUTTON, BN_CLICKED, OnBnClickedAboutCodec)
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
				DWORD dwExportTypeCine = VsiGetRangeValue<DWORD>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
					VSI_PARAMETER_SYS_DEFAULT_EXPORT_LOOP, m_pPdm);

				DWORD dwExportTypeCineFccHandler = VsiGetRangeValue<DWORD>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
					VSI_PARAMETER_SYS_DEFAULT_EXPORT_LOOP_HANDLER, m_pPdm);

				CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));
				int iDefaultiIndex(0);
				int iIndex;

				// Always add uncompressed type
				{
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"Uncompressed AVI (*.avi)");
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_AVI_CUSTOM);

					// Uncompressed AVI
					m_pdwFccHandlerType[iIndex] = VSI_4CC_UNCOMP;
					if (dwExportTypeCine == VSI_IMAGE_AVI_CUSTOM &&
						m_pdwFccHandlerType[iIndex] == dwExportTypeCineFccHandler)
					{
						iDefaultiIndex = iIndex;
					}
				}

				// Always add animated gif type if this feature has been enabled
				{
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"Animated GIF (*.gif)");
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_AVI_CUSTOM);

					// Uncompressed AVI
					m_pdwFccHandlerType[iIndex] = VSI_4CC_ANIMGIF;
					if (dwExportTypeCine == VSI_IMAGE_AVI_CUSTOM &&
						m_pdwFccHandlerType[iIndex] == dwExportTypeCineFccHandler)
					{
						iDefaultiIndex = iIndex;
					}
				}

				// If we have eng mode enumerate all
				if (VsiModeUtils::GetIsEngineerMode(m_pPdm))
				{
					ICINFO picInfo[256];
					DWORD dwNumCodecs = VsiEnumerateAviCodecs(picInfo, _countof(picInfo));

					for (DWORD i=0; i<dwNumCodecs; ++i)
					{
						CString szExportType(picInfo[i].szDescription);
						szExportType += L" (*.avi)";

						iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)(szExportType.GetBuffer()));

						// Add FCC handler as type
						wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_AVI_CUSTOM);

						m_pdwFccHandlerType[iIndex] = picInfo[i].fccHandler;
						if (dwExportTypeCine == VSI_IMAGE_AVI_CUSTOM &&
							m_pdwFccHandlerType[iIndex] == dwExportTypeCineFccHandler)
						{
							iDefaultiIndex = iIndex;
						}
					}
				}
				else
				{
					BOOL bAvailable = VsiIsAviCodecAvailable(VSI_4CC_COMP_MSVIDEO9);
					if (bAvailable)
					{
						iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"Compressed AVI MS Media Video 9 (*.avi)");
						wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_AVI_CUSTOM);

						m_pdwFccHandlerType[iIndex] = VSI_4CC_COMP_MSVIDEO9;
						if (dwExportTypeCine == VSI_IMAGE_AVI_CUSTOM &&
							m_pdwFccHandlerType[iIndex] == dwExportTypeCineFccHandler)
						{
							iDefaultiIndex = iIndex;
						}
					}

					bAvailable = VsiIsAviCodecAvailable(VSI_4CC_COMP_MSVIDEO1);
					if (bAvailable)
					{
						iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"Compressed AVI MS Video 1 (*.avi)");
						wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_AVI_CUSTOM);

						m_pdwFccHandlerType[iIndex] = VSI_4CC_COMP_MSVIDEO1;
						if (dwExportTypeCine == VSI_IMAGE_AVI_CUSTOM &&
							m_pdwFccHandlerType[iIndex] == dwExportTypeCineFccHandler)
						{
							iDefaultiIndex = iIndex;
						}
					}
				}

				// Add wave format
				if (m_pExportImage->IsAllLoopsOfModes(VSI_SETTINGS_MODE_PW_DOPPLER) ||
					m_pExportImage->IsAllLoopsOfModes(VSI_SETTINGS_MODE_PW_TISSUE_DOPPLER))
				{
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL,(LPARAM)L"Windows Audio Wave File (*.wav)");
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_WAV);
					if (VSI_IMAGE_WAV == dwExportTypeCine)
					{
						iDefaultiIndex = iIndex;
					}
				}

				// Add RF if we have RF images
				if (m_pExportImage->IsAllOfRf(VSI_RF_EXPORT_IQ))
				{
					CString strLabel(L"IQ Data (*." VSI_RF_FILE_EXTENSION_IQXML L")");
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)(LPCWSTR)strLabel);
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_IQ_DATA_LOOP);

					if (dwExportTypeCine == VSI_IMAGE_IQ_DATA_LOOP)
					{
						iDefaultiIndex = iIndex;
					}
				}
				if (m_pExportImage->IsAllOfRf(VSI_RF_EXPORT_RF))
				{
					CString strLabel(L"RF Data (*." VSI_RF_FILE_EXTENSION_RFXML L")");
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)(LPCWSTR)strLabel);
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_RF_DATA_LOOP);

					if (dwExportTypeCine == VSI_IMAGE_RF_DATA_LOOP)
					{
						iDefaultiIndex = iIndex;
					}
				}
				if (m_pExportImage->IsAllOfRf(VSI_RF_EXPORT_RAW))
				{
					CString strLabel(L"RAW Data (*." VSI_RF_FILE_EXTENSION_RAWXML L")");
					iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)(LPCWSTR)strLabel);
					wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_IMAGE_RAW_DATA_LOOP);

					if (dwExportTypeCine == VSI_IMAGE_RAW_DATA_LOOP)
					{
						iDefaultiIndex = iIndex;
					}
				}

				// Select a default file type
				wndExportType.SendMessage(CB_SETCURSEL, iDefaultiIndex, NULL);
			}

			// Check quality settings
			DWORD dwAviQuality = VsiGetDiscreteValue<DWORD>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_SYS_EXPORT_AVI_QUALITY, m_pPdm);

			if (dwAviQuality == VSI_SYS_EXPORT_AVI_QUALITY_HIGH)
			{
				CheckDlgButton(IDC_IE_QUALITY_HIGH, BST_CHECKED);
			}
			else
			{
				CheckDlgButton(IDC_IE_QUALITY_MED, BST_CHECKED);
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
	// Show the codec configuration if it exists
	LRESULT OnBnClickedConfigCodec(WORD, WORD, HWND, BOOL&)
	{
		CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));

		int iIndex = (int)wndExportType.SendMessage(CB_GETCURSEL, NULL, NULL);
		DWORD dwType = (DWORD)wndExportType.SendMessage(CB_GETITEMDATA, iIndex, NULL);

		if (dwType == VSI_IMAGE_AVI_CUSTOM)
		{
			DWORD dwFccHandler = m_pdwFccHandlerType[iIndex];

			if (dwFccHandler != VSI_4CC_ANIMGIF)
			{
				// Don't bother for gif type, custom
				HRESULT hr = VsiShowAviCodecConfig(dwFccHandler, m_hWnd);
				if (FAILED(hr))
				{
					VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure showing configuration dialog");
				}
			}
		}

		return 0;
	}
	// Show the codec about if it exists
	LRESULT OnBnClickedAboutCodec(WORD, WORD, HWND, BOOL&)
	{
		CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));

		int iIndex = (int)wndExportType.SendMessage(CB_GETCURSEL, NULL, NULL);
		DWORD dwType = (DWORD)wndExportType.SendMessage(CB_GETITEMDATA, iIndex, NULL);

		if (dwType == VSI_IMAGE_AVI_CUSTOM)
		{
			DWORD dwFccHandler = m_pdwFccHandlerType[iIndex];

			if (dwFccHandler != VSI_4CC_ANIMGIF)
			{
				// Don't bother for gif type, custom
				HRESULT hr = VsiShowAviCodecAbout(dwFccHandler, m_hWnd);
				if (FAILED(hr))
				{
					VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure showing about dialog");
				}
			}
		}

		return 0;
	}
};