/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportImageTableDlg.h
**
**	Description:
**		Declaration of the CVsiExportImageTableDlg
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiPdmModule.h>
#include "VsiExportImage.h"


class CVsiExportImageTableDlg :
	public CDialogImpl<CVsiExportImageTableDlg>
{
private:
	CVsiExportImage *m_pExportImage;
	CComPtr<IVsiPdm> m_pPdm;

public:
	enum { IDD = IDD_EXPORT_IMAGE_VIEW_TABLE };

public:
	CVsiExportImageTableDlg(CVsiExportImage *pExportImage) :
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
		return TRUE;
	}

protected:
	BEGIN_MSG_MAP(CVsiExportImageTableDlg)
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
				CWindow wndExportType(GetDlgItem(IDC_IE_TYPE));
				int iDefaultiIndex(0);
				int iIndex;

				wndExportType.SendMessage(CB_RESETCONTENT);
				iIndex = (int)wndExportType.SendMessage(CB_ADDSTRING, NULL, (LPARAM)L"VSI Table File (*.txt)");
				wndExportType.SendMessage(CB_SETITEMDATA, iIndex, VSI_EXPORT_TABLE);

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
};