/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  CVsiOperatorProp.h
**
**	Description:
**		Declaration of the CVsiOperatorProp
**
*******************************************************************************/

#pragma once

#include <VsiStl.h>
#include <VsiRes.h>
#include <VsiCommCtl.h>
#include <VsiOperatorModule.h>
#include <VsiAppModule.h>
#include <VsiToolTipHelper.h>
#include <VsiPathNameEditWnd.h>
#include <VsiSelectOperator.h>
#include <VsiStudyAccessCombo.h>


class CVsiOperatorProp :
	public CDialogImpl<CVsiOperatorProp>,
	public CVsiToolTipHelper<CVsiOperatorProp>
{
public:
	typedef enum
	{
		VSI_OPERATION_NONE = 0,
		VSI_OPERATION_ADD,
		VSI_OPERATION_MODIFY,
	} VSI_OPERATION;

	CImageList m_ilGroupAdd;
	CString m_strName;
	CString m_strPassword;
	BOOL m_bUpdatePassword;
	CString m_strGroup;
	VSI_OPERATOR_TYPE m_dwType;
	VSI_OPERATOR_DEFAULT_STUDY_ACCESS m_dwDefaultStudyAccess;
	CString m_strIdCopyFrom;

private:
	CComQIPtr<IVsiApp> m_pApp;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;
	CComPtr<IVsiOperator> m_pOperator;

	CVsiStudyAccessCombo m_studyAccess;

	bool m_bSecured;
	bool m_bSelectedOpLoggedIn;
	bool m_bSecuredAdminLoggedIn;
	bool m_bHasPassword;
	bool m_bPasswordModified;

	VSI_OPERATION m_operation;

	CVsiPathNameEditWnd m_wndOperatorName;
	CVsiSelectOperator m_wndOperatorListCombo;

	CWindow m_wndGroupList;

	VSI_OPERATOR_TYPE m_dwOriginalType;

	typedef enum
	{
		VSI_TIPS_NONE = 0,
		VSI_TIPS_NAME,
		VSI_TIPS_NAME_EXISTS,
		VSI_TIPS_PASSWORD_MUST_EXISTS,
		VSI_TIPS_PASSWORD,
		VSI_TIPS_CURRENT_PASSWORD,
	} VSI_TIPS;  // Tips state

	VSI_TIPS m_TipsState;

public:
	DWORD IDD;

	CVsiOperatorProp(
		VSI_OPERATION operation,
		IVsiApp *pApp,
		IVsiOperator *pOperator,
		bool bSecureMode = false,
		bool bSelectedOpLoggedIn = false);
	~CVsiOperatorProp();

protected:
	BEGIN_MSG_MAP(CVsiOperatorProp)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)

		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)
		COMMAND_HANDLER(IDC_OPERATOR_GROUP_NEW, BN_CLICKED, OnBnClickedGroupNew)
		COMMAND_HANDLER(IDC_OPERATOR_TYPE_NORMAL, BN_CLICKED, OnBnClickedStandard)
		COMMAND_HANDLER(IDC_OPERATOR_TYPE_ADMIN, BN_CLICKED, OnBnClickedAdmin)
		COMMAND_HANDLER(IDC_OPERATOR_PROTECTED, BN_CLICKED, OnBnClickedProtected)
		COMMAND_HANDLER(IDC_OPERATOR_PASSWORD_CHANGE, BN_CLICKED, OnBnClickedChangePassword)
		COMMAND_CODE_HANDLER(EN_UPDATE, OnEnUpdate);
		COMMAND_HANDLER(IDC_OPERATOR_PASSWORD, EN_SETFOCUS, OnEnSetFocus);
		COMMAND_HANDLER(IDC_OPERATOR_PASSWORD2, EN_SETFOCUS, OnEnSetFocus);
		COMMAND_CODE_HANDLER(EN_KILLFOCUS, OnEnKillFocus);

		NOTIFY_CODE_HANDLER(TTN_GETDISPINFO, OnTipsGetDisplayInfo);

		CHAIN_MSG_MAP(CVsiToolTipHelper<CVsiOperatorProp>)
		CHAIN_MSG_MAP_MEMBER(m_wndOperatorListCombo)
	END_MSG_MAP()

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wp, LPARAM lp, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedGroupNew(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedStandard(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedAdmin(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedProtected(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedChangePassword(WORD, WORD, HWND, BOOL&);
	LRESULT OnEnUpdate(WORD, WORD, HWND, BOOL&);
	LRESULT OnEnSetFocus(WORD, WORD, HWND, BOOL&);
	LRESULT OnEnKillFocus(WORD, WORD, HWND, BOOL&);

	LRESULT OnTipsGetDisplayInfo(int, LPNMHDR, BOOL&);

private:
	void UpdateUI();
	void ActivateTips(VSI_TIPS state);
	void DeactivateTips();
	void FillGroupCombo();
};

