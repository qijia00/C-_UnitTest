/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiSelectOperator.h
**
**	Description:
**		Declaration of the CVsiSelectOperator
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiResProduct.h>

class CVsiSelectOperator : public CWindow
{
protected:
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;
	CComPtr<IVsiCommandManager> m_pCmdMgr;
	WORD m_wId;
	BOOL m_bCheckPassword;
	BOOL m_bOperatorSetup;
	BOOL m_bAllowNone;
	CString m_strDefaultOperator;

public:
	CVsiSelectOperator() :
		m_wId(0)
	{
	}

	~CVsiSelectOperator()
	{
		_ASSERT(0 == m_wId);  // Must clean up memory in Uninitialize
	}
	
	HRESULT Initialize(
		BOOL bCheckPassword,
		BOOL bAllowOperatorSetup,
		BOOL bAllowNone,
		HWND hwnd,
		IVsiOperatorManager *pOperatorMgr,
		IVsiCommandManager *pCmdMgr)
	{
		// We still want to set this so we can control display
		m_hWnd = hwnd;
		m_bCheckPassword = bCheckPassword;
		m_bOperatorSetup = bAllowOperatorSetup;
		m_bAllowNone = bAllowNone;
		m_hWnd = hwnd;
		m_pOperatorMgr = pOperatorMgr;
		m_pCmdMgr = pCmdMgr;
		m_wId = (WORD)GetDlgCtrlID();

		CString strSelect(MAKEINTRESOURCE(IDS_SB_SELECT_OPERATOR));
		SendMessage(CB_SETCUEBANNER, 0, (LPARAM)strSelect.GetString());

		return S_OK;
	}

	HRESULT Uninitialize()
	{
		// Clean up operator list
		Clear();

		m_wId = 0;
		m_pCmdMgr.Release();
		m_pOperatorMgr.Release();
		m_hWnd = NULL;
		
		return S_OK;
	}
	
	HRESULT Clear()
	{
		// Clean up operator list
		if (IsWindow())
		{
			int iCount = (int)SendMessage(CB_GETCOUNT, 0, 0);
			for (int i = 0; i < iCount; ++i)
			{
				OLECHAR *pszOperator = (OLECHAR*)SendMessage(CB_GETITEMDATA, i);
				if (! IS_INTRESOURCE(pszOperator))
				{
					CComHeapPtr<OLECHAR> pszId;
					pszId.Attach(pszOperator);
					pszId.Free();
				}
			}

			SendMessage(CB_RESETCONTENT);
		}
		
		return S_OK;
	}
	
	HRESULT Fill(LPCWSTR pszDefault)
	{
		HRESULT hr(S_OK);

		// May not be initialized
		if (NULL == m_pOperatorMgr)
			return S_FALSE;
		
		try
		{
			BOOL bSelected = FALSE;

			if (NULL != pszDefault)
			{
				m_strDefaultOperator = pszDefault;
			}
			else
			{
				m_strDefaultOperator.Empty();

				int i = (int)SendMessage(CB_GETCURSEL);
				if (CB_ERR != i)
				{
					LPCWSTR pszId = (LPCWSTR)SendMessage(CB_GETITEMDATA, i, 0);
					if (! IS_INTRESOURCE(pszId))
					{
						GetWindowText(m_strDefaultOperator);
					}
				}
			}
			
			Clear();

			if (m_bAllowNone)
			{
				int iIndex = (int)SendMessage(CB_ADDSTRING, 0, (LPARAM)L"     ");
				VSI_VERIFY(CB_ERR != iIndex, VSI_LOG_ERROR, NULL);
				SendMessage(CB_SETCURSEL, iIndex);
			}

			CComPtr<IVsiEnumOperator> pOperators;
			hr = m_pOperatorMgr->GetOperators(
				VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN,
				&pOperators);
			if (SUCCEEDED(hr))
			{
				CComPtr<IVsiOperator> pOperator;
				while (S_OK == pOperators->Next(1, &pOperator, NULL))
				{
					VSI_OPERATOR_STATE dwState(VSI_OPERATOR_STATE_DISABLE);

					hr = pOperator->GetState(&dwState);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

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

					int iIndex = (int)SendMessage(CB_ADDSTRING, 0, (LPARAM)pszName.m_pData);
					VSI_VERIFY(CB_ERR != iIndex, VSI_LOG_ERROR, NULL);

					if (!bSelected)
					{
						if (m_strDefaultOperator == pszName)
						{
							bSelected = TRUE;
							SendMessage(CB_SETCURSEL, iIndex);
						}
					}

					int iRet = (int)SendMessage(CB_SETITEMDATA, iIndex, (LPARAM)pszId.m_pData);
					VSI_VERIFY(CB_ERR != iRet, VSI_LOG_ERROR, NULL);

					pszId.Detach();

					pOperator.Release();
				}
			}

			// Add commands
			if (m_bOperatorSetup)
			{
				CString strSetup(MAKEINTRESOURCE(IDS_OPERATOR_SETUP));
				int iIndex = (int)SendMessage(CB_ADDSTRING, 0, (LPARAM)(LPCWSTR)strSetup);
				VSI_VERIFY(CB_ERR != iIndex, VSI_LOG_ERROR, NULL);

				int iRet = (int)SendMessage(CB_SETITEMDATA, iIndex, (LPARAM)ID_CMD_OPERATOR);
				VSI_VERIFY(CB_ERR != iRet, VSI_LOG_ERROR, NULL);
			}
		}
		VSI_CATCH(hr);

		return hr;
	}

protected:
	BEGIN_MSG_MAP(CVsiSelectOperator)
		COMMAND_HANDLER(m_wId, CBN_CLOSEUP, OnOperatorCloseUp)
		COMMAND_HANDLER(m_wId, CBN_SELENDOK, OnOperatorSelenOk)
	END_MSG_MAP()


	LRESULT OnOperatorCloseUp(WORD, WORD, HWND, BOOL &bHandled)
	{
		HRESULT hr(S_OK);
		bHandled = FALSE;
	
		try
		{
			int i = (int)SendMessage(CB_GETCURSEL);
			if (CB_ERR == i)
			{
				SendMessage(CB_SETCURSEL, static_cast<WPARAM>(-1));

				CString strSelect(MAKEINTRESOURCE(IDS_SB_SELECT_OPERATOR));
				SendMessage(CB_SETCUEBANNER, 0, (LPARAM)strSelect.GetString());
			}
			else
			{
				BOOL bNoOperator = FALSE;

				LPCWSTR pszId = (LPCWSTR)SendMessage(CB_GETITEMDATA, i, 0);
				if (IS_INTRESOURCE(pszId))
				{
					if (!m_strDefaultOperator.IsEmpty())
					{
						int iSel = (int)SendMessage(
							CB_FINDSTRINGEXACT, (WPARAM)-1, (LPARAM)(LPCWSTR)m_strDefaultOperator);
						if (CB_ERR != iSel)
						{
							SendMessage(CB_SETCURSEL, iSel);
						}
						else
						{
							bNoOperator = TRUE;
						}
					}
					else
					{
						bNoOperator = TRUE;
					}
				}

				if (bNoOperator)
				{
					SendMessage(CB_SETCURSEL, static_cast<WPARAM>(-1));

					CString strSelect(MAKEINTRESOURCE(IDS_SB_SELECT_OPERATOR));
					SendMessage(CB_SETCUEBANNER, 0, (LPARAM)strSelect.GetString());
				}
			}
		}
		VSI_CATCH(hr);

		return 0;
	}


	LRESULT OnOperatorSelenOk(WORD, WORD, HWND, BOOL &bHandled)
	{
		HRESULT hr(S_OK);
		bHandled = FALSE;

		try
		{
			int i = (int)SendMessage(CB_GETCURSEL);
			if (CB_ERR != i)
			{
				BOOL bNoneSelected = FALSE;
				LPCWSTR pszId = (LPCWSTR)SendMessage(CB_GETITEMDATA, i, 0);

				if (m_bAllowNone)
				{
					if (NULL == pszId)
					{
						bNoneSelected = TRUE;
					}
				}

				if (!bNoneSelected)
				{
					if (IS_INTRESOURCE(pszId))
					{
						DWORD dwCmd = (DWORD)(ULONG_PTR)pszId;

						if (NULL != m_pCmdMgr)
						{
							m_pCmdMgr->InvokeCommandAsync(dwCmd, NULL, FALSE);
						}
					
						if (ID_CMD_LOG_OUT == dwCmd)
						{
							m_strDefaultOperator.Empty();
							SendMessage(CB_SETCURSEL, (WPARAM)-1);
						}
					}
					else
					{
						GetWindowText(m_strDefaultOperator);
					}
				}
			}
		}
		VSI_CATCH(hr);

		return 0;
	}
};
