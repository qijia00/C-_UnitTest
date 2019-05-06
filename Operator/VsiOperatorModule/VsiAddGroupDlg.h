/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudy.h
**
**	Description:
**		Declaration and Implementation of the CVsiAddGroupDlg
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiCommCtlLib.h>
#include <VsiPathNameEditWnd.h>
#include "resource.h"       // main symbols

class CVsiAddGroupDlg :
	public CDialogImpl<CVsiAddGroupDlg>
{
private:
	CString* m_pstrGroupName;

public:
	enum { IDD = IDD_NEW_GROUP };

	CVsiAddGroupDlg(CString &strGroupName)
	{
		m_pstrGroupName = &strGroupName;
	}

	~CVsiAddGroupDlg()
	{
	}

protected:
	// Message map and handlers
	BEGIN_MSG_MAP(CVsiNewFolder)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)

		COMMAND_ID_HANDLER(IDOK, OnCmdOK)
		COMMAND_ID_HANDLER(IDCANCEL, OnCmdCancel)

		COMMAND_CODE_HANDLER(EN_UPDATE, OnEnChange)
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
	{
		HRESULT hr(S_OK);

		try
		{
			// Sets font
			HFONT hFont;
			VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
			VsiThemeRecurSetFont(m_hWnd, hFont);

			// Sets header and footer
			HIMAGELIST hilHeader(NULL);
			VsiThemeGetImageList(VSI_THEME_IL_HEADER_NORMAL, &hilHeader);
			VsiThemeDialogBox(*this, TRUE, TRUE, hilHeader);

			GetDlgItem(IDC_NG_NAME).SendMessage(CB_SETCUEBANNER, FALSE, (LPARAM)L"Type a Name");

			PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)GetDlgItem(IDC_NG_NAME), TRUE);
		}
		VSI_CATCH(hr);

		if (FAILED(hr))
		{
			EndDialog(IDCANCEL);
		}

		return 0;
	}

	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
	{
		return 0;
	}

	LRESULT OnCmdOK(WORD, WORD, HWND, BOOL&)
	{
		VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"New Group - OK");

		GetDlgItem(IDC_NG_NAME).GetWindowText(*m_pstrGroupName);

		EndDialog(IDOK);

		return 0;
	}

	LRESULT OnCmdCancel(WORD, WORD, HWND, BOOL&)
	{
		VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"New Group - Cancel");

		EndDialog(IDCANCEL);

		return 0;
	}

	LRESULT OnEnChange(WORD, WORD, HWND, BOOL& bHandled)
	{
		CString strName;
		GetDlgItem(IDC_NG_NAME).GetWindowText(strName);
		strName.Trim();

		GetDlgItem(IDOK).EnableWindow(!strName.IsEmpty());

		bHandled = FALSE;

		return 0;
	}
};