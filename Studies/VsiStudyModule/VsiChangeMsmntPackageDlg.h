/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiChangeMsmntPackageDlg.h
**
**	Description:
**		Declaration of the CVsiChangeMsmntPackageDlg
**
*******************************************************************************/

#pragma once


#include <VsiRes.h>
#include <VsiAppModule.h>
#include <VsiStudyModule.h>


class CVsiChangeMsmntPackageDlg :
	public CDialogImpl<CVsiChangeMsmntPackageDlg>
{
private:
	CComPtr<IVsiApp> m_pApp;
	int m_iImageCount;
	CComPtr<IVsiEnumImage> m_pEnumImages;

public:
	enum { IDD = IDD_CHANGE_MEASUREMENT_PACKAGE };

	CVsiChangeMsmntPackageDlg(IVsiApp *pApp, int iImages, IVsiEnumImage *pEnumImages);
	~CVsiChangeMsmntPackageDlg();

protected:
	BEGIN_MSG_MAP(CVsiChangeMsmntPackageDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		COMMAND_ID_HANDLER(IDOK, OnCmdOk)
		COMMAND_ID_HANDLER(IDCANCEL, OnCmdCancel)
		COMMAND_HANDLER(IDC_CMPKG_PACKAGES, CBN_SELCHANGE, OnPackageSelChange)
	END_MSG_MAP()

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wp, LPARAM lp, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnClose(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnCmdOk(WORD, WORD, HWND, BOOL&);
	LRESULT OnCmdCancel(WORD, WORD, HWND, BOOL&);
	LRESULT OnPackageSelChange(WORD, WORD, HWND, BOOL&);
};

