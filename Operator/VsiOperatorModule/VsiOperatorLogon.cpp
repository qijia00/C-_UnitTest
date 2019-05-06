/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorLogon.cpp
**
**	Description:
**		Implementation of CVsiOperatorLogon
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiServiceProvider.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiServiceKey.h>
#include <VsiParameterUtils.h>
#include <VsiParameterSystem.h>
#include <VsiParameterNamespace.h>
#include "VsiOperatorLogon.h"


CVsiOperatorLogon::CVsiOperatorLogon() :
	MsgHook((TMFP)&CVsiOperatorLogon::MsgProc, this),
	m_pszTitle(NULL),
	m_pszMsg(NULL),
	m_dwType(VSI_OPERATOR_TYPE_NONE),
	m_bUpdateCurrentOperator(FALSE)
{
}

CVsiOperatorLogon::~CVsiOperatorLogon()
{
}

STDMETHODIMP CVsiOperatorLogon::Activate(
	IVsiApp *pApp,
	HWND hwndParent,
	LPCOLESTR pszTitle,
	LPCOLESTR pszMsg,
	VSI_OPERATOR_TYPE dwType,
	BOOL bUpdateCurrentOperator,
	IVsiOperator **ppOperator)
{
	HRESULT hr = S_OK;

	try
	{
		m_pApp = pApp;
		m_pszTitle = pszTitle;
		m_pszMsg = pszMsg;
		m_dwType = dwType;
		m_bUpdateCurrentOperator = bUpdateCurrentOperator;

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_OPERATOR_MGR,
			__uuidof(IVsiOperatorManager),
			(IUnknown**)&m_pOperatorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"CVsiOperatorLogon::Activate failed");

		// Get PDM
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (ppOperator != NULL)
		{
			m_pOperatorSelected = *ppOperator;
		}

		INT_PTR iRet = DoModal(hwndParent);

		if (IDOK == iRet && NULL != ppOperator)
		{
			if (*ppOperator != NULL)
				(*ppOperator)->Release();

			*ppOperator = m_pOperatorSelected.Detach();
		}

		hr = (IDOK == iRet) ? S_OK : S_FALSE;
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
			VSI_ERROR_LOG(err);
	}

	return hr;
}

LRESULT CVsiOperatorLogon::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	LRESULT lRet = 1;

	try
	{
		// Tips
		InsertTips(GetDlgItem(IDC_LOGIN_NAME_COMBO),
			L"Select a user.");

		InsertTips(GetDlgItem(IDC_LOGIN_PASSWORD_EDIT),
			L"The password is incorrect.\r\nRetype the password and try again.");

		// Set font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
		VsiThemeRecurSetFont(m_hWnd, hFont);

		// Set header and footer
		HIMAGELIST hilHeader(NULL);
		VsiThemeGetImageList(VSI_THEME_IL_HEADER_NORMAL, &hilHeader);
		VsiThemeDialogBox(*this, TRUE, TRUE, hilHeader);

		if (NULL != m_pszTitle)
			SetWindowText(m_pszTitle);

		if (NULL != m_pszMsg)
			SetDlgItemText(IDC_LOGIN_DESCRIPTION, m_pszMsg);

		// Populate list
		m_wndName = GetDlgItem(IDC_LOGIN_NAME_COMBO);
		m_wndPasswordLabel = GetDlgItem(IDC_LOGIN_PASSWORD_LABEL);
		m_wndPassword = GetDlgItem(IDC_LOGIN_PASSWORD_EDIT);

		// Sets limits
		SendDlgItemMessage(IDC_LOGIN_PASSWORD_EDIT, EM_SETLIMITTEXT, VSI_OPERATOR_PASSWORD_MAX);

		FillNameCombo();

		UpdateUI();

		if (m_wndPassword.IsWindowEnabled())
		{
			m_wndPassword.SetFocus();
			lRet = 0;
		}

		m_hHook = SetWindowsHookEx(
			WH_MSGFILTER,
			(HOOKPROC)MsgHook::GetThunk(),
			_AtlBaseModule.GetModuleInstance(),
			GetCurrentThreadId());
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
			VSI_ERROR_LOG(err);

		EndDialog(IDCANCEL);
	}

	return lRet;
}

LRESULT CVsiOperatorLogon::OnDestroy(UINT, WPARAM, LPARAM, BOOL& bHandled)
{
	if (NULL != m_hHook)
	{
		UnhookWindowsHookEx(m_hHook);
	}

	Cleanup();

	bHandled = FALSE;

	m_pPdm.Release();

	m_pOperatorMgr.Release();

	m_pApp.Release();

	return 0;
}

LRESULT CVsiOperatorLogon::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Operator Logon - OK");

	DeactivateTips();
	
	int iIndex = (int)m_wndName.SendMessage(CB_GETCURSEL);

	if (CB_ERR != iIndex)
	{
		CComPtr<IVsiOperator> pOperator =
			(IVsiOperator*)m_wndName.SendMessage(CB_GETITEMDATA, iIndex);

		WCHAR szPassword[VSI_OPERATOR_PASSWORD_MAX + 1];
		m_wndPassword.GetWindowText(szPassword, _countof(szPassword));

		if (S_OK != pOperator->TestPassword(szPassword))
		{
			m_wndPassword.SetFocus();
			m_wndPassword.SendMessage(EM_SETSEL, 0, -1);

			ActivateTips(GetDlgItem(IDC_LOGIN_PASSWORD_EDIT));

			return 0;
		}

		m_pOperatorSelected = pOperator;

		if (m_bUpdateCurrentOperator)
		{
			CComHeapPtr<OLECHAR> pszOperId;
			m_pOperatorSelected->GetId(&pszOperId);

			// Set new current operator
			if (NULL != pszOperId)
			{
				m_pOperatorMgr->SetCurrentOperator(pszOperId);

				if (S_OK != pOperator->HasPassword())
				{
					// add operator to the set of logged in operators
					m_pOperatorMgr->SetIsAuthenticated(pszOperId, TRUE);
				}
			}

			// Set last logged on operator
			CComVariant vOperId(pszOperId);
			VsiSetRangeValue<LPCWSTR>(V_BSTR(&vOperId),
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_LAST_LOGGED_IN, m_pPdm);
		}
	}
	else
	{
		m_wndName.SetFocus();

		ActivateTips(GetDlgItem(IDC_LOGIN_NAME_COMBO));

		return 0;
	}

	EndDialog(IDOK);

	return 0;
}

LRESULT CVsiOperatorLogon::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Operator Logon - Cancel");

	DeactivateTips();

	EndDialog(IDCANCEL);

	return 0;
}

LRESULT CVsiOperatorLogon::OnSelChange(WORD, WORD, HWND, BOOL&)
{
	DeactivateTips();

	UpdateUI();

	return 0;
}

LRESULT CVsiOperatorLogon::OnCbCloseUp(WORD, WORD, HWND, BOOL&)
{
	if (m_wndPassword.IsWindowEnabled())
		m_wndPassword.SetFocus();

	return 0;
}

LRESULT CVsiOperatorLogon::OnCbChange(WORD, WORD, HWND, BOOL& bHandled)
{
	DeactivateTips();

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiOperatorLogon::OnEnChange(WORD, WORD, HWND, BOOL& bHandled)
{
	DeactivateTips();

	bHandled = FALSE;

	return 0;
}

void CVsiOperatorLogon::FillNameCombo(bool bServiceMode)
{
	HRESULT hr = S_OK;

	try
	{
		// UI
		Cleanup();

		CComHeapPtr<OLECHAR> pszIdSelected;
		if (m_pOperatorSelected != NULL)
		{
			hr = m_pOperatorSelected->GetId(&pszIdSelected);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Get all operators and filter them in the loop
		VSI_OPERATOR_FILTER_TYPE filter =
			VSI_OPERATOR_FILTER_TYPE_STANDARD |
			VSI_OPERATOR_FILTER_TYPE_ADMIN |
			VSI_OPERATOR_FILTER_TYPE_PASSWORD;

		if (bServiceMode)
		{
			filter |= VSI_OPERATOR_FILTER_TYPE_SERVICE_MODE;
		}

		CComPtr<IVsiEnumOperator> pOperators;
		hr = m_pOperatorMgr->GetOperators(
			filter,
			&pOperators);
		BOOL bSelected = FALSE;
		if (SUCCEEDED(hr))
		{
			CComPtr<IVsiOperator> pOperator;
			while (pOperators->Next(1, &pOperator, NULL) == S_OK)
			{
				VSI_OPERATOR_STATE dwState;
				pOperator->GetState(&dwState);
				if (VSI_OPERATOR_STATE_DISABLE == dwState)
				{
					pOperator.Release();
					continue;
				}

				CComHeapPtr<OLECHAR> pszName;
				hr = pOperator->GetName(&pszName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComHeapPtr<OLECHAR> pszId;
				hr = pOperator->GetId(&pszId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (0 != lstrcmp(pszIdSelected, pszId))
				{
					// Not selected - check type
					VSI_OPERATOR_TYPE dwType;
					hr = pOperator->GetType(&dwType);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (0 == (m_dwType & dwType))
					{
						pOperator.Release();
						continue;
					}
					else 
					{
						if ((m_dwType & VSI_OPERATOR_TYPE_PASSWORD_MANDATRORY))
						{
							CComHeapPtr<OLECHAR> pszPassword;
							if (S_OK != pOperator->HasPassword())
							{
								pOperator.Release();
								continue;
							}
						}
					}
				}

				int iIndex = (int)m_wndName.SendMessage(CB_ADDSTRING, 0, (LPARAM)pszName.m_pData);
				if (CB_ERR != iIndex)
				{
					// Set the config page as part of the item data
					int iRet = (int)m_wndName.SendMessage(CB_SETITEMDATA, iIndex, (LPARAM)pOperator.p);
					_ASSERT(CB_ERR != iRet);
					pOperator.Detach();
				}

				if (!bSelected)
				{
					if (0 == lstrcmp(pszIdSelected, pszId))
					{
						bSelected = TRUE;
						m_wndName.SendMessage(CB_SETCURSEL, iIndex);
					}
				}

				pOperator.Release();
			}

			// Select first operator if none passed in
			if (!bSelected || bServiceMode)
			{
				m_wndName.SendMessage(CB_SETCURSEL, 0);
			}
		}
	}
	VSI_CATCH(hr);
}

void CVsiOperatorLogon::UpdateUI()
{
	int index = (int)m_wndName.SendMessage(CB_GETCURSEL);
	if (CB_ERR != index)
	{
		IVsiOperator *pOperator = (IVsiOperator*)m_wndName.SendMessage(CB_GETITEMDATA, index);

		if (S_OK == pOperator->HasPassword())
		{
			m_wndPasswordLabel.EnableWindow(TRUE);
			m_wndPassword.EnableWindow(TRUE);
		}
		else
		{
			m_wndPasswordLabel.EnableWindow(FALSE);
			m_wndPassword.EnableWindow(FALSE);
		}
	}
	else
	{
		m_wndPasswordLabel.EnableWindow(FALSE);
		m_wndPassword.EnableWindow(FALSE);
	}
}

void CVsiOperatorLogon::Cleanup()
{
	int iCount = (int)m_wndName.SendMessage(CB_GETCOUNT, 0, 0);
	for (int i = 0; i < iCount; ++i)
	{
		IVsiOperator *pOperator = (IVsiOperator*)m_wndName.SendMessage(CB_GETITEMDATA, i);
		pOperator->Release();
	}

	m_wndName.SendMessage(CB_RESETCONTENT);
}

LRESULT CVsiOperatorLogon::MsgProc(int nCode, WPARAM wp, LPARAM lp)
{
	MSG *pMsg = reinterpret_cast<MSG*>(lp);

	if (WM_KEYDOWN == pMsg->message)
	{
		// Ctrl-Alt-S
		if (GetKeyState(VK_CONTROL) < 0  && GetKeyState(VK_MENU) < 0 && GetKeyState(L'S') < 0)
		{
			FillNameCombo(true);
			UpdateUI();
		}
	}

	return CallNextHookEx(m_hHook, nCode, wp, lp);
}
