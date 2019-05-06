/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorLogon.h
**
**	Description:
**		Declaration of the CVsiOperatorLogon
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiThunk.h>
#include <VsiToolTipHelper.h>
#include <VsiOperatorModule.h>
#include <VsiAppModule.h>
#include "resource.h"


class ATL_NO_VTABLE CVsiOperatorLogon :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiOperatorLogon, &CLSID_VsiOperatorLogon>,
	public CDialogImpl<CVsiOperatorLogon>,
	public CVsiToolTipHelper<CVsiOperatorLogon>,
	public CVsiThunkImpl<CVsiOperatorLogon>,
	public IVsiOperatorLogon
{
private:
	typedef CVsiThunkImpl<CVsiOperatorLogon> MsgHook;
	HHOOK m_hHook;

	CComQIPtr<IVsiApp> m_pApp;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;
	CComPtr<IVsiPdm> m_pPdm;

	CComPtr<IVsiOperator> m_pOperatorSelected;

	LPCWSTR m_pszTitle;
	LPCWSTR m_pszMsg;
	VSI_OPERATOR_TYPE m_dwType;
	BOOL m_bUpdateCurrentOperator;

	CWindow m_wndName;
	CWindow m_wndPasswordLabel;
	CWindow m_wndPassword;

public:
	enum { IDD = IDD_OPERATOR_LOGIN };

	CVsiOperatorLogon();
	~CVsiOperatorLogon();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSIOPERATORLOGON)

	BEGIN_COM_MAP(CVsiOperatorLogon)
		COM_INTERFACE_ENTRY(IVsiOperatorLogon)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

	// CDialogImplBaseT
	virtual void OnFinalMessage(HWND hWnd)
	{
		UNREFERENCED_PARAMETER(hWnd);
	}

	VSI_DECLARE_USE_MESSAGE_LOG()

public:
	// IVsiOperatorLogon
	STDMETHOD(Activate)(
		IVsiApp *pApp,
		HWND hwndParent,
		LPCOLESTR pszTitle,
		LPCOLESTR pszMsg,
		VSI_OPERATOR_TYPE dwType,
		BOOL bUpdateCurrentOperator,
		IVsiOperator **ppOperator);

protected:
	BEGIN_MSG_MAP(CVsiOperatorLogon)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)

		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)

		COMMAND_HANDLER(IDC_LOGIN_NAME_COMBO, CBN_SELENDOK, OnSelChange)
		COMMAND_HANDLER(IDC_LOGIN_NAME_COMBO, CBN_CLOSEUP, OnCbCloseUp)
		COMMAND_HANDLER(IDC_LOGIN_NAME_COMBO, CBN_DROPDOWN, OnCbChange)
		COMMAND_HANDLER(IDC_LOGIN_NAME_COMBO, CBN_KILLFOCUS, OnCbChange)

		COMMAND_CODE_HANDLER(EN_UPDATE, OnEnChange);
		COMMAND_CODE_HANDLER(EN_KILLFOCUS, OnEnChange);

		CHAIN_MSG_MAP(CVsiToolTipHelper<CVsiOperatorLogon>)
	END_MSG_MAP()

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wp, LPARAM lp, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnSelChange(WORD, WORD, HWND, BOOL&);
	LRESULT OnCbCloseUp(WORD, WORD, HWND, BOOL&);
	LRESULT OnCbChange(WORD, WORD, HWND, BOOL&);

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);

	LRESULT OnEnChange(WORD, WORD, HWND, BOOL&);

private:
	void Refresh();
	void UpdateUI();
	void FillNameCombo(bool bServiceMode = false);
	void Cleanup();

	LRESULT MsgProc(int nCode, WPARAM wp, LPARAM lp);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiOperatorLogon), CVsiOperatorLogon)
