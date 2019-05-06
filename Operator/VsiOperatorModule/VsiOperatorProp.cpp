/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  CVsiOperatorProp.cpp
**
**	Description:
**		Implementation of CVsiOperatorProp
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiResProduct.h>
#include <VsiServiceProvider.h>
#include <VsiCommUtl.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiThemeColor.h>
#include <VsiServiceKey.h>
#include "VsiAddGroupDlg.h"
#include "VsiOperatorProp.h"


#define VSI_DUMMY_PASSWORD L"*=D<a_*={}`=Q]"


CVsiOperatorProp::CVsiOperatorProp(
	VSI_OPERATION operation,
	IVsiApp *pApp,
	IVsiOperator *pOperator,
	bool bSecureMode,
	bool bSelectedOpLoggedIn)
:
	IDD(0),
	m_operation(operation),
	m_pOperator(pOperator),
	m_bUpdatePassword(FALSE),
	m_dwType(VSI_OPERATOR_TYPE_STANDARD),
	m_dwOriginalType(VSI_OPERATOR_TYPE_STANDARD),
	m_dwDefaultStudyAccess(VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE),
	m_TipsState(VSI_TIPS_NONE),
	m_bSecured(bSecureMode),
	m_bSelectedOpLoggedIn(bSelectedOpLoggedIn),
	m_bSecuredAdminLoggedIn(false),
	m_bHasPassword(false),
	m_bPasswordModified(false)
{
	HRESULT hr;

	try
	{
		m_pApp = pApp;

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_OPERATOR_MGR,
			__uuidof(IVsiOperatorManager),
			(IUnknown**)&m_pOperatorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (NULL != pOperator)
		{
			CComHeapPtr<OLECHAR> pszName;
			hr = pOperator->GetName(&pszName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			m_strName = pszName;

			m_strPassword = VSI_DUMMY_PASSWORD;

			VSI_OPERATOR_TYPE dwType;
			hr = pOperator->GetType(&dwType);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_dwType = dwType;
			m_dwOriginalType = dwType;

			VSI_OPERATOR_DEFAULT_STUDY_ACCESS dwStudyAccess;
			hr = pOperator->GetDefaultStudyAccess(&dwStudyAccess);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_dwDefaultStudyAccess = dwStudyAccess;

			CComHeapPtr<OLECHAR> pszGroup;
			hr = pOperator->GetGroup(&pszGroup);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_strGroup = pszGroup;

			hr = pOperator->HasPassword();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			if (S_OK == hr)
			{
				m_bHasPassword = true;
			}
		}

		if (m_bSecured)
		{
			VSI_OPERATOR_TYPE dwType = VSI_OPERATOR_TYPE_NONE;

			CComPtr<IVsiOperator> pOperatorCurrent;
			hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorCurrent);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (NULL != pOperatorCurrent)
			{
				hr = pOperatorCurrent->GetType(&dwType);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				m_bSecuredAdminLoggedIn = VSI_OPERATOR_TYPE_ADMIN == (VSI_OPERATOR_TYPE_BASE_MASK & dwType);
			}

			switch (m_operation)
			{
			case VSI_OPERATION_ADD:
				{
					IDD = IDD_OPERATOR_ADD;
				}
				break;
			case VSI_OPERATION_MODIFY:
				{
					IDD = IDD_OPERATOR_MODIFY;
				}
				break;
			default:
				{
					_ASSERT(0);
				}
				break;
			}
		}
		else
		{
			IDD = IDD_OPERATOR_SET;
		}
	}
	VSI_CATCH(hr);
}

CVsiOperatorProp::~CVsiOperatorProp()
{
}

LRESULT CVsiOperatorProp::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr(S_OK);

	try
	{
		m_wndOperatorName.SubclassWindow(GetDlgItem(IDC_OPERATOR_NAME));
		m_wndOperatorName.SendMessage(EM_SETCUEBANNER, FALSE, (LPARAM)L"Type a Name");

		// Tips
		InsertTips(GetDlgItem(IDC_OPERATOR_NAME), L"");
		InsertTips(GetDlgItem(IDC_OPERATOR_PASSWORD), L"");

		// Set font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
		VsiThemeRecurSetFont(m_hWnd, hFont);

		// Sets header and footer
		HIMAGELIST hilHeader(NULL);
		VsiThemeGetImageList(VSI_THEME_IL_HEADER_NORMAL, &hilHeader);
		VsiThemeDialogBox(*this, TRUE, TRUE, hilHeader);

		// Sets limits
		SendDlgItemMessage(IDC_OPERATOR_NAME, EM_SETLIMITTEXT, VSI_OPERATOR_NAME_MAX);
		SendDlgItemMessage(IDC_OPERATOR_PASSWORD, EM_SETLIMITTEXT, VSI_OPERATOR_PASSWORD_MAX);
		SendDlgItemMessage(IDC_OPERATOR_PASSWORD2, EM_SETLIMITTEXT, VSI_OPERATOR_PASSWORD_MAX);
		if (m_bSecured && VSI_OPERATION_MODIFY == m_operation)
		{
			SendDlgItemMessage(IDC_OPERATOR_PASSWORD_OLD, EM_SETLIMITTEXT, VSI_OPERATOR_PASSWORD_MAX);
			if (m_bSecuredAdminLoggedIn && !m_bSelectedOpLoggedIn)
			{
				GetDlgItem(IDC_OPERATOR_PASSWORD_OLD_LABEL).SetWindowText(L"&Your Password");
			}
		}

		// Sets value
		SetDlgItemText(IDC_OPERATOR_NAME, m_strName);
		if (VSI_OPERATION_ADD == m_operation)
		{
			GetDlgItem(IDC_OPERATOR_NAME).SendMessage(EM_SETREADONLY, FALSE);
		}

		// Set Protected check box state
		if (!m_bSecured)
		{
			CheckDlgButton(IDC_OPERATOR_PROTECTED, m_bHasPassword ? BST_CHECKED : BST_UNCHECKED);

			if (VSI_OPERATOR_TYPE_STANDARD == m_dwType || !m_bHasPassword)
			{
				GetDlgItem(IDC_OPERATOR_PROTECTED).EnableWindow(TRUE);
			}
		}

		if (!m_bHasPassword)
		{
			m_strPassword.Empty();
		}

		BOOL bAllowType(FALSE);
		if (m_bSecured)
		{
			m_ilGroupAdd.CreateFromImage(
				IDB_STUDY_ACCESS_GROUP_ADD,
				29,
				1,
				0,
				IMAGE_BITMAP,
				LR_CREATEDIBSECTION);

			VSI_SETIMAGE si = { 0 };
			for (int i = 0; i < _countof(si.uStyleFlags); ++i)
			{
				si.uStyleFlags[i] = VSI_BTN_IMAGE_STYLE_DRAWTHEME;
			}
			si.hImageLists[VSI_BTN_IMAGE_UNPUSHED] = m_ilGroupAdd;
			si.iImageIndex = 0;

			UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);
			GetDlgItem(IDC_OPERATOR_GROUP_NEW).SendMessage(uiThemeCmd, VSI_BT_CMD_SETIMAGE, (LPARAM)&si);
			GetDlgItem(IDC_OPERATOR_GROUP_NEW).SendMessage(uiThemeCmd, VSI_BT_CMD_SETTIP, (LPARAM)L"Add Group");

			if (VSI_OPERATION_ADD == m_operation)
			{
				GetDlgItem(IDC_OPERATOR_GROUP_LABEL).EnableWindow(TRUE);
				GetDlgItem(IDC_OPERATOR_STUDY_GROUP_COMBO).EnableWindow(TRUE);
				GetDlgItem(IDC_OPERATOR_GROUP_NEW).EnableWindow(TRUE);
			}
			else
			{
				GetDlgItem(IDC_OPERATOR_GROUP_LABEL).EnableWindow(m_bSecuredAdminLoggedIn);
				GetDlgItem(IDC_OPERATOR_STUDY_GROUP_COMBO).EnableWindow(m_bSecuredAdminLoggedIn);
				GetDlgItem(IDC_OPERATOR_STUDY_GROUP_COMBO).SendMessage(
					uiThemeCmd, VSI_CB_CMD_SETREADONLY, m_bSecuredAdminLoggedIn ? 0 : 1);
				GetDlgItem(IDC_OPERATOR_GROUP_NEW).EnableWindow(m_bSecuredAdminLoggedIn);
			}

			// Admin cannot change self to standard
			bAllowType = (VSI_OPERATION_ADD == m_operation) || (m_bSecuredAdminLoggedIn && !m_bSelectedOpLoggedIn);

			GetDlgItem(IDC_OPERATOR_TYPE_LABEL).EnableWindow(bAllowType);
			GetDlgItem(IDC_OPERATOR_TYPE_ADMIN).EnableWindow(bAllowType);

			CheckDlgButton(IDC_OPERATOR_TYPE_ADMIN, (VSI_OPERATOR_TYPE_ADMIN == m_dwType) ? BST_CHECKED : BST_UNCHECKED);
		}
		else
		{
			GetDlgItem(IDC_OPERATOR_TYPE_NORMAL).EnableWindow(TRUE);
			GetDlgItem(IDC_OPERATOR_TYPE_ADMIN).EnableWindow(TRUE);

			CheckDlgButton(
				(VSI_OPERATOR_TYPE_STANDARD == m_dwType) ?
					IDC_OPERATOR_TYPE_NORMAL : IDC_OPERATOR_TYPE_ADMIN,
				BST_CHECKED);
		}

		if (m_bSecured && VSI_OPERATION_MODIFY == m_operation)
		{
			SetDlgItemText(IDC_OPERATOR_PASSWORD_OLD, m_strPassword);
		}
		SetDlgItemText(IDC_OPERATOR_PASSWORD, m_strPassword);
		SetDlgItemText(IDC_OPERATOR_PASSWORD2, m_strPassword);

		// Set study access radio buttons and label state
		if (m_bSecured)
		{
			m_studyAccess.Initialize(GetDlgItem(IDC_OPERATOR_STUDY_ACCESS));

			CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE type(CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_NONE);

			switch (m_dwDefaultStudyAccess)
			{
			case VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE:
				type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PRIVATE;
				break;
			case VSI_OPERATOR_DEFAULT_STUDY_ACCESS_GROUP:
				type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_GROUP;
				break;
			case VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PUBLIC:
				type = CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PUBLIC;
				break;
			}

			m_studyAccess.SetStudyAccess(type);

			// Fill Group list
			m_wndGroupList = GetDlgItem(IDC_OPERATOR_STUDY_GROUP_COMBO);

			FillGroupCombo();
		}

		if (m_bSecured && VSI_OPERATION_ADD == m_operation)
		{
			// Fill operator list
			hr = m_wndOperatorListCombo.Initialize(
				FALSE,
				FALSE,
				TRUE,
				GetDlgItem(IDC_OPERATOR_CLONE),
				m_pOperatorMgr,
				NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_wndOperatorListCombo.Fill(NULL);

			// Mandatory fields
			UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);
			
			SendDlgItemMessage(
				IDC_OPERATOR_NAME_LABEL,
				uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
			SendDlgItemMessage(
				IDC_OPERATOR_PASSWORD_LABEL,
				uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
			SendDlgItemMessage(
				IDC_OPERATOR_PASSWORD2_LABEL,
				uiThemeCmd, VSI_STC_CMD_SETTEXTCOLOR, VSI_COLOR_MANDATORY);
		}

		UpdateUI();

		int iFocus(IDC_OPERATOR_NAME);
		if (VSI_OPERATION_MODIFY == m_operation)
		{
			if (m_bSecured)
			{
				iFocus = IDC_OPERATOR_STUDY_ACCESS;
			}
			else
			{
				if (GetDlgItem(IDC_OPERATOR_PROTECTED).IsWindowEnabled())
				{
					iFocus = IDC_OPERATOR_PROTECTED;
				}
				else 
				{
					iFocus = ((VSI_OPERATOR_TYPE_ADMIN == m_dwType) ?
						IDC_OPERATOR_TYPE_NORMAL : IDC_OPERATOR_TYPE_ADMIN);
				}
			}
		}

		PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)GetDlgItem(iFocus), TRUE);
	}
	VSI_CATCH(hr);

	if (FAILED(hr))
	{
		EndDialog(IDCANCEL);
	}

	return 0;
}

LRESULT CVsiOperatorProp::OnDestroy(UINT, WPARAM, LPARAM, BOOL& bHandled)
{
	bHandled = FALSE;

	m_ilGroupAdd.Destroy();
	m_wndOperatorListCombo.Uninitialize();

	return 0;
}

LRESULT CVsiOperatorProp::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Operator Properties - OK");

	HRESULT hr = S_OK;

	WCHAR szPassword[VSI_OPERATOR_PASSWORD_MAX + 1];
	WCHAR szPassword2[VSI_OPERATOR_PASSWORD_MAX + 1];
	WCHAR szPasswordOld[VSI_OPERATOR_PASSWORD_MAX + 1];

	DeactivateTips();

	CString strName;
	GetDlgItemText(IDC_OPERATOR_NAME, strName);
	strName.Trim();

	if (strName.IsEmpty())
	{
		GetDlgItem(IDC_OPERATOR_NAME).SetFocus();

		ActivateTips(VSI_TIPS_NAME);

		return 0;
	}

	if (VSI_OPERATION_ADD == m_operation)
	{
		CComPtr<IVsiOperator> pOperator;
		HRESULT hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, strName, &pOperator);
		BOOL bNameExists = S_OK == hr;
		if (bNameExists)
		{
			GetDlgItem(IDC_OPERATOR_NAME).SetFocus();
			SendDlgItemMessage(IDC_OPERATOR_NAME, EM_SETSEL, 0, -1);

			ActivateTips(VSI_TIPS_NAME_EXISTS);

			return 0;
		}
	}

	if (m_bSecured)
	{
		CComPtr<IVsiOperator> pOperatorCurrent;
		hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorCurrent);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (NULL == pOperatorCurrent)
		{
			// Abnormal for secure mode
			_ASSERT(0);
			return 0;
		}

		BOOL bAdmin = BST_CHECKED == IsDlgButtonChecked(IDC_OPERATOR_TYPE_ADMIN);

		// Password
		BOOL bModifyPasswordChecked = TRUE;

		if (VSI_OPERATION_MODIFY == m_operation)
		{
			bModifyPasswordChecked = BST_CHECKED == IsDlgButtonChecked(IDC_OPERATOR_PASSWORD_CHANGE);
		}

		if (bModifyPasswordChecked)
		{
			// Check old password when changing password
			if (VSI_OPERATION_MODIFY == m_operation)
			{
				GetDlgItemText(IDC_OPERATOR_PASSWORD_OLD, szPasswordOld, _countof(szPasswordOld));

				if (NULL != pOperatorCurrent)
				{
					if (S_OK != pOperatorCurrent->TestPassword(szPasswordOld))
					{
						GetDlgItem(IDC_OPERATOR_PASSWORD_OLD).SetFocus();
						SendDlgItemMessage(IDC_OPERATOR_PASSWORD_OLD, EM_SETSEL, 0, -1);

						ActivateTips(VSI_TIPS_CURRENT_PASSWORD);

						return 0;
					}
				}
			}

			GetDlgItemText(IDC_OPERATOR_PASSWORD, szPassword, _countof(szPassword));
			GetDlgItemText(IDC_OPERATOR_PASSWORD2, szPassword2, _countof(szPassword2));

			// Password is required when adding adding new operator
			BOOL bCheckPassword(TRUE);

			if (VSI_OPERATION_MODIFY == m_operation)
			{
				// For admin, he can disable other account with an empty password
				// For admin, he cannot disable any administrator account(including himself)
				if (m_bSecuredAdminLoggedIn && !bAdmin)
				{
					bCheckPassword = FALSE;
				}
			}

			if (bCheckPassword && 0 == *szPassword)
			{
				GetDlgItem(IDC_OPERATOR_PASSWORD).SetFocus();

				ActivateTips(VSI_TIPS_PASSWORD_MUST_EXISTS);

				return 0;
			}

			// Make sure both passwords are the same in all cases
			if (0 != wcscmp(szPassword, szPassword2))
			{
				GetDlgItem(IDC_OPERATOR_PASSWORD).SetFocus();
				SendDlgItemMessage(IDC_OPERATOR_PASSWORD, EM_SETSEL, 0, -1);

				ActivateTips(VSI_TIPS_PASSWORD);

				return 0;
			}
		}
		else
		{
			*szPassword = 0;
			*szPassword2 = 0;
		}

		m_strName = strName;
		m_dwType = bAdmin ?	VSI_OPERATOR_TYPE_ADMIN : VSI_OPERATOR_TYPE_STANDARD;
		m_bUpdatePassword = FALSE;

		if (bModifyPasswordChecked)
		{
			m_strPassword = szPassword;

			// Make sure the password is real
			if (0 != wcscmp(szPassword, VSI_DUMMY_PASSWORD))
			{
				m_bUpdatePassword = TRUE;
			}
		}

		// Group info
		CString strGroup;
		m_wndGroupList.GetWindowText(strGroup);
		m_strGroup = strGroup.Trim();

		// Clone Settings
		if (VSI_OPERATION_ADD == m_operation)
		{
			int index = (int)GetDlgItem(IDC_OPERATOR_CLONE).SendMessage(CB_GETCURSEL);
			if (CB_ERR != index)
			{
				LPCWSTR pszId = (LPCWSTR)GetDlgItem(IDC_OPERATOR_CLONE).SendMessage(CB_GETITEMDATA, index);
				if (! IS_INTRESOURCE(pszId))
				{
					m_strIdCopyFrom = pszId;
				}
			}
		}

		// Access info
		CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE type = m_studyAccess.GetStudyAccess();
		switch (type)
		{
		case CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PRIVATE:
			m_dwDefaultStudyAccess = VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PRIVATE;
			break;
		case CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_GROUP:
			m_dwDefaultStudyAccess = VSI_OPERATOR_DEFAULT_STUDY_ACCESS_GROUP;
			break;
		case CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PUBLIC:
			m_dwDefaultStudyAccess = VSI_OPERATOR_DEFAULT_STUDY_ACCESS_PUBLIC;
			break;
		default:
			_ASSERT(0);
			m_dwDefaultStudyAccess = VSI_OPERATOR_DEFAULT_STUDY_ACCESS_NONE;
			break;
		}

		EndDialog(IDOK);
	}
	else
	{
		BOOL bProtected = (BST_CHECKED == IsDlgButtonChecked(IDC_OPERATOR_PROTECTED));
		if (bProtected)
		{
			GetDlgItemText(IDC_OPERATOR_PASSWORD, szPassword, _countof(szPassword));
			GetDlgItemText(IDC_OPERATOR_PASSWORD2, szPassword2, _countof(szPassword2));
		}
		else
		{
			*szPassword = 0;
			*szPassword2 = 0;
		}

		if (bProtected && (0 == *szPassword))
		{
			GetDlgItem(IDC_OPERATOR_PASSWORD).SetFocus();
			SendDlgItemMessage(IDC_OPERATOR_PASSWORD, EM_SETSEL, 0, -1);

			ActivateTips(VSI_TIPS_PASSWORD_MUST_EXISTS);

			return 0;
		}
		else if (lstrcmp(szPassword, szPassword2) != 0)
		{
			GetDlgItem(IDC_OPERATOR_PASSWORD).SetFocus();
			SendDlgItemMessage(IDC_OPERATOR_PASSWORD, EM_SETSEL, 0, -1);

			ActivateTips(VSI_TIPS_PASSWORD);

			return 0;
		}

		m_strName = strName;
		m_strPassword = szPassword;
		// Make sure the password is real
		if (0 != wcscmp(szPassword, VSI_DUMMY_PASSWORD))
		{
			m_bUpdatePassword = TRUE;
		}

		if (BST_CHECKED == IsDlgButtonChecked(IDC_OPERATOR_TYPE_ADMIN))
		{
			m_dwType = VSI_OPERATOR_TYPE_ADMIN;
		}
		else
		{
			m_dwType = VSI_OPERATOR_TYPE_STANDARD;
		}

		DWORD dwCount = 0;
		m_pOperatorMgr->GetOperatorCount(
			VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_PASSWORD,
			&dwCount);
		if (0 == dwCount)
		{
			EndDialog(IDOK);
		}
		else if ((VSI_OPERATION_MODIFY == m_operation) && (m_dwType == m_dwOriginalType) &&
			!m_bHasPassword)
		{
			// If the operator has no password, and changes something that does not require
			//  administrator rights, let it through
			EndDialog(IDOK);
		}
		else
		{
			HRESULT hr = S_OK;

			try
			{
				CComPtr<IVsiOperatorLogon> pLogon;
				hr = pLogon.CoCreateInstance(__uuidof(VsiOperatorLogon));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				WCHAR szTitle[] = L"Verification";
				CString szMsg;
				if ((VSI_OPERATION_MODIFY == m_operation) && (m_dwType == m_dwOriginalType))
				{
					szMsg = L"Only the user himself or administrator can perform this task.";
				}
				else
				{
					szMsg = L"Only the administrator can perform this task.";
				}
				hr = pLogon->Activate(
					m_pApp,
					GetParent(),
					szTitle,
					szMsg,
					VSI_OPERATOR_TYPE_ADMIN,
					FALSE,
					(m_dwType == m_dwOriginalType) ? ((NULL == m_pOperator) ? NULL : &(m_pOperator.p)) : NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					EndDialog(IDOK);
				}

				ShowWindow(SW_SHOW);
			}
			VSI_CATCH(hr);
		}
	}

	return 0;
}

LRESULT CVsiOperatorProp::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Operator Properties - Cancel");

	DeactivateTips();

	EndDialog(IDCANCEL);

	return 0;
}

LRESULT CVsiOperatorProp::OnBnClickedGroupNew(WORD, WORD, HWND, BOOL&)
{
	CString strGroupName;
	CVsiAddGroupDlg group(strGroupName);
	INT_PTR iRet = group.DoModal(*this);

	if (IDOK == iRet)
	{
		int iIndex = (int)m_wndGroupList.SendMessage(CB_ADDSTRING, 0, (LPARAM)(LPCWSTR)strGroupName);
		VSI_VERIFY(CB_ERR != iIndex, VSI_LOG_ERROR, NULL);
		m_wndGroupList.SendMessage(CB_SETCURSEL, iIndex);
	}

	return 0;
}

LRESULT CVsiOperatorProp::OnBnClickedStandard(WORD, WORD, HWND, BOOL&)
{
	if (!m_bSecured)
	{
		GetDlgItem(IDC_OPERATOR_PROTECTED).EnableWindow(TRUE);

		UpdateUI();
	}

	return 0;
}

LRESULT CVsiOperatorProp::OnBnClickedAdmin(WORD, WORD, HWND, BOOL&)
{
	if (m_bSecured)
	{		
		if (VSI_OPERATION_MODIFY == m_operation)
		{
			if (!m_bHasPassword)
			{
				bool bIsAdmin = BST_CHECKED == IsDlgButtonChecked(IDC_OPERATOR_TYPE_ADMIN);
				bool bChangePassword = BST_CHECKED == IsDlgButtonChecked(IDC_OPERATOR_PASSWORD_CHANGE);

				CheckDlgButton(IDC_OPERATOR_PASSWORD_CHANGE, bIsAdmin ? BST_CHECKED : BST_UNCHECKED);
				GetDlgItem(IDC_OPERATOR_PASSWORD_CHANGE).EnableWindow(!bIsAdmin);
				GetDlgItem(IDC_OPERATOR_PASSWORD_OLD_LABEL).EnableWindow(bIsAdmin);
				GetDlgItem(IDC_OPERATOR_PASSWORD_LABEL).EnableWindow(bIsAdmin);
				GetDlgItem(IDC_OPERATOR_PASSWORD2_LABEL).EnableWindow(bIsAdmin);
				GetDlgItem(IDC_OPERATOR_PASSWORD_OLD).EnableWindow(bIsAdmin);
				GetDlgItem(IDC_OPERATOR_PASSWORD).EnableWindow(bIsAdmin);
				GetDlgItem(IDC_OPERATOR_PASSWORD2).EnableWindow(bIsAdmin);

				if (bIsAdmin && !bChangePassword)
				{
					SetDlgItemText(IDC_OPERATOR_PASSWORD_OLD, L"");
					SetDlgItemText(IDC_OPERATOR_PASSWORD, L"");
					SetDlgItemText(IDC_OPERATOR_PASSWORD2, L"");
				}
				if (bIsAdmin)
				{
					GetDlgItem(IDC_OPERATOR_PASSWORD_OLD).SetFocus();
				}
			}
		}
	}
	else
	{
		CheckDlgButton(IDC_OPERATOR_PROTECTED, BST_CHECKED);
		GetDlgItem(IDC_OPERATOR_PROTECTED).EnableWindow(FALSE);

		UpdateUI();

		if (!m_bHasPassword)
		{
			GetDlgItem(IDC_OPERATOR_PASSWORD).SetFocus();
		}
	}

	return 0;
}

LRESULT CVsiOperatorProp::OnBnClickedProtected(WORD, WORD, HWND, BOOL&)
{
	UpdateUI();

	if (BST_CHECKED != IsDlgButtonChecked(IDC_OPERATOR_PROTECTED))
	{
		SetDlgItemText(IDC_OPERATOR_PASSWORD, L"");
		SetDlgItemText(IDC_OPERATOR_PASSWORD2, L"");
	}
	else
	{
		GetDlgItem(IDC_OPERATOR_PASSWORD).SetFocus();
	}

	return 0;
}

LRESULT CVsiOperatorProp::OnBnClickedChangePassword(WORD, WORD, HWND, BOOL&)
{
	UpdateUI();

	if (BST_CHECKED == IsDlgButtonChecked(IDC_OPERATOR_PASSWORD_CHANGE))
	{
		SetDlgItemText(IDC_OPERATOR_PASSWORD_OLD, L"");
		SetDlgItemText(IDC_OPERATOR_PASSWORD, L"");
		SetDlgItemText(IDC_OPERATOR_PASSWORD2, L"");

		GetDlgItem(IDC_OPERATOR_PASSWORD_OLD).SetFocus();
	}
	else
	{
		if (m_bHasPassword)
		{
			SetDlgItemText(IDC_OPERATOR_PASSWORD_OLD, VSI_DUMMY_PASSWORD);
			SetDlgItemText(IDC_OPERATOR_PASSWORD, VSI_DUMMY_PASSWORD);
			SetDlgItemText(IDC_OPERATOR_PASSWORD2, VSI_DUMMY_PASSWORD);
		}
		else
		{
			SetDlgItemText(IDC_OPERATOR_PASSWORD_OLD, L"");
			SetDlgItemText(IDC_OPERATOR_PASSWORD, L"");
			SetDlgItemText(IDC_OPERATOR_PASSWORD2, L"");
		}
	}

	return 0;
}

LRESULT CVsiOperatorProp::OnEnUpdate(WORD, WORD, HWND, BOOL& bHandled)
{
	DeactivateTips();

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiOperatorProp::OnEnSetFocus(WORD, WORD, HWND, BOOL& bHandled)
{
	if (VSI_OPERATION_MODIFY == m_operation)
	{
		if (!m_bSecured && !m_bPasswordModified)
		{
			SetDlgItemText(IDC_OPERATOR_PASSWORD, L"");
			SetDlgItemText(IDC_OPERATOR_PASSWORD2, L"");
			m_bPasswordModified = true;
		}
	}

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiOperatorProp::OnEnKillFocus(WORD, WORD, HWND, BOOL& bHandled)
{
	DeactivateTips();

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiOperatorProp::OnTipsGetDisplayInfo(int, LPNMHDR pnmh, BOOL&)
{
	LPNMTTDISPINFO pttd = (LPNMTTDISPINFO)pnmh;

	::SendMessage(pnmh->hwndFrom, TTM_SETMAXTIPWIDTH, 0, 300);

	if (VSI_TIPS_NAME == m_TipsState)
		pttd->lpszText = L"The name is not valid.";
	else if (VSI_TIPS_NAME_EXISTS == m_TipsState)
		pttd->lpszText = L"The name already exists.";
	else if (VSI_TIPS_PASSWORD_MUST_EXISTS == m_TipsState)
		pttd->lpszText = L"The password cannot be left blank.";
	else if (VSI_TIPS_PASSWORD == m_TipsState)
		pttd->lpszText = L"The passwords do not match.\nRetype the new password in both fields.";
	else if (VSI_TIPS_CURRENT_PASSWORD == m_TipsState)
		pttd->lpszText = L"Password is invalid.\nRetype the password.";
	else
		pttd->lpszText = L"";

	return 0;
}

void CVsiOperatorProp::UpdateUI()
{
	if (m_bSecured)
	{
		if (VSI_OPERATION_MODIFY == m_operation)
		{
			BOOL bChecked = BST_CHECKED == IsDlgButtonChecked(IDC_OPERATOR_PASSWORD_CHANGE);
			GetDlgItem(IDC_OPERATOR_PASSWORD_OLD_LABEL).EnableWindow(bChecked);
			GetDlgItem(IDC_OPERATOR_PASSWORD_LABEL).EnableWindow(bChecked);
			GetDlgItem(IDC_OPERATOR_PASSWORD2_LABEL).EnableWindow(bChecked);
			GetDlgItem(IDC_OPERATOR_PASSWORD_OLD).EnableWindow(bChecked);
			GetDlgItem(IDC_OPERATOR_PASSWORD).EnableWindow(bChecked);
			GetDlgItem(IDC_OPERATOR_PASSWORD2).EnableWindow(bChecked);
		}
	}
	else
	{
		BOOL bEanble = IsDlgButtonChecked(IDC_OPERATOR_PROTECTED);

		WORD pwIds[] =
		{
			IDC_OPERATOR_PASSWORD_LABEL,
			IDC_OPERATOR_PASSWORD,
			IDC_OPERATOR_PASSWORD2_LABEL,
			IDC_OPERATOR_PASSWORD2
		};

		for (int i = 0; i < _countof(pwIds); ++i)
		{
			GetDlgItem(pwIds[i]).EnableWindow(bEanble);
		}
	}
}

void CVsiOperatorProp::ActivateTips(VSI_TIPS state)
{
	m_TipsState = state;

	WORD wId = 0;
	switch (state)
	{
	case VSI_TIPS_NAME:
	case VSI_TIPS_NAME_EXISTS:
		wId = IDC_OPERATOR_NAME;
		break;
	case VSI_TIPS_PASSWORD_MUST_EXISTS:
		wId = IDC_OPERATOR_PASSWORD;
		break;
	case VSI_TIPS_PASSWORD:
		wId = IDC_OPERATOR_PASSWORD;
		break;
	case VSI_TIPS_CURRENT_PASSWORD:
		wId = IDC_OPERATOR_PASSWORD_OLD;
		break;
	default:
		_ASSERT(0);
	}

	CVsiToolTipHelper<CVsiOperatorProp>::ActivateTips(GetDlgItem(wId));
}

void CVsiOperatorProp::DeactivateTips()
{
	m_TipsState = VSI_TIPS_NONE;

	CVsiToolTipHelper<CVsiOperatorProp>::DeactivateTips();
}

void CVsiOperatorProp::FillGroupCombo()
{
	HRESULT hr = S_OK;

	try
	{
		// Empty list
		m_wndGroupList.SendMessage(CB_RESETCONTENT);

		int iIndex = (int)m_wndGroupList.SendMessage(CB_ADDSTRING, 0, (LPARAM)L"     ");
		VSI_VERIFY(CB_ERR != iIndex, VSI_LOG_ERROR, NULL);
		SendMessage(CB_SETCURSEL, iIndex);
		
		VSI_OPERATOR_FILTER_TYPE types = VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN;

		int iSelect(-1);
		int iCount(0);
		CComPtr<IVsiEnumOperator> pOperators;
		hr = m_pOperatorMgr->GetOperators(types, &pOperators);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		std::set<CString> setGroups;
		CComPtr<IVsiOperator> pOperator;
		while (pOperators->Next(1, &pOperator, NULL) == S_OK)
		{
			CComHeapPtr<OLECHAR> pszGroup;
			hr = pOperator->GetGroup(&pszGroup);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (0 != *(pszGroup.m_pData))
			{
				auto iter = setGroups.find(pszGroup.m_pData);
				if (setGroups.end() == iter)
				{
					setGroups.insert(pszGroup.m_pData);

					iIndex = (int)m_wndGroupList.SendMessage(CB_ADDSTRING, 0, (LPARAM)pszGroup.m_pData);
					if (CB_ERR != iIndex)
					{
						++iCount;
					}

					if (m_strGroup == pszGroup)
					{
						iSelect = iIndex;
						m_wndGroupList.SendMessage(CB_SETCURSEL, iSelect);
					}
				}
			}

			pOperator.Release();
		}

		if (0 == iCount)
		{
			// No operators, no group
		}
	}
	VSI_CATCH(hr);
}

