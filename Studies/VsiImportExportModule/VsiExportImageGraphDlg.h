/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportImageGraphDlg.h
**
**	Description:
**		Declaration of the CVsiExportImageGraphDlg
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiPdmModule.h>
#include "VsiExportImage.h"


class CVsiExportImageGraphDlg :
	public CDialogImpl<CVsiExportImageGraphDlg>
{
private:
	CVsiExportImage *m_pExportImage;
	CComPtr<IVsiPdm> m_pPdm;

public:
	enum { IDD = IDD_EXPORT_IMAGE_VIEW_GRAPH };

public:
	CVsiExportImageGraphDlg(CVsiExportImage *pExportImage) :
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
		}
		VSI_CATCH(hr);

		return hr;
	}

	BOOL UpdateUI()
	{
		// Disable the OK button in case all are not checked
		BOOL bOK = BST_CHECKED == IsDlgButtonChecked(IDC_IE_EXPORT_CSV) ||
			BST_CHECKED == IsDlgButtonChecked(IDC_IE_EXPORT_BMP) ||
			BST_CHECKED == IsDlgButtonChecked(IDC_IE_EXPORT_TIFF);
		return bOK;
	}

protected:
	BEGIN_MSG_MAP(CVsiExportImageGraphDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		COMMAND_HANDLER(IDC_IE_SAVEAS, EN_CHANGE, OnFileNameChanged)
		COMMAND_HANDLER(IDC_IE_EXPORT_CSV, BN_CLICKED, OnFileTypeChanged)
		COMMAND_HANDLER(IDC_IE_EXPORT_BMP, BN_CLICKED, OnFileTypeChanged)
		COMMAND_HANDLER(IDC_IE_EXPORT_TIFF, BN_CLICKED, OnFileTypeChanged)
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

			// Fill types (Dummy to keep VsiExportImage happy)
			{
				CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));
				int iDefaultiIndex(0);
				int iIndex;

				wndExportType.SendMessage(CB_RESETCONTENT);
				iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"Graph");
				wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_EXPORT_GRAPH);

				// Select a default file type
				wndExportType.SendMessage(CB_SETCURSEL, iDefaultiIndex, NULL);
			}

			if (VSI_EXPORT_IMAGE_FLAG_GRAPH_PARAM & m_pExportImage->m_dwFlags)
			{
				GetDlgItem(IDC_IE_EXPORT_GRAPH_PARAM).EnableWindow(TRUE);
				GetDlgItem(IDC_IE_EXPORT_GRAPH_PARAM).ShowWindow(SW_SHOW);
			}
			else
			{
				GetDlgItem(IDC_IE_EXPORT_GRAPH_PARAM).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_EXPORT_GRAPH_PARAM).ShowWindow(SW_HIDE);
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
};