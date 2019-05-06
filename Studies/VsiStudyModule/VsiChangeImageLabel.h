/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiChangeImageLabel.h
**
**	Description:
**		Declaration of the CVsiChangeImageLabel
**
*******************************************************************************/

#pragma once

#include <shldisp.h>
#include <VsiRes.h>
#include <VsiStudyModule.h>
#include <VsiMessageFilter.h>
#include <VsiModalDlgBase.h>
#include "resource.h"


class ATL_NO_VTABLE CVsiChangeImageLabel :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiChangeImageLabel, &CLSID_VsiChangeImageLabel>,
	public CDialogImpl<CVsiChangeImageLabel>,
	public CVsiModalDlgBase<CVsiChangeImageLabel>,
	public IVsiPreTranslateAccelFilter,
	public IVsiChangeImageLabel
{
private:
	CComPtr<IVsiApp> m_pApp;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IUnknown> m_pUnkItem;

	CComPtr<IVsiCursorManager> m_pCursorMgr;

	CComPtr<IAutoComplete> m_pAutoComplete;
	CComPtr<IVsiAutoCompleteSource> m_pAutoCompleteSource;

	BOOL m_bCommit;

public:
	enum { IDD = IDD_IMAGE_LABEL };

	CVsiChangeImageLabel();
	~CVsiChangeImageLabel();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSICHANGEIMAGELABEL)

	BEGIN_COM_MAP(CVsiChangeImageLabel)
		COM_INTERFACE_ENTRY(IVsiChangeImageLabel)
		COM_INTERFACE_ENTRY(IVsiPreTranslateAccelFilter)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

	VSI_DECLARE_USE_MESSAGE_LOG()

public:
	// IVsiChangeImageLabel
	STDMETHOD(Activate)(
		IVsiApp *pApp,
		HWND hwndParent,
		IUnknown *pUnkItem);
	STDMETHOD(Deactivate)();

	// IVsiPreTranslateAccelFilter
	STDMETHOD(PreTranslateAccel)(MSG *pMsg, BOOL *pbHandled);

protected:
	BEGIN_MSG_MAP(CVsiChangeImageLabel)
		CHAIN_MSG_MAP(CVsiModalDlgBase<CVsiChangeImageLabel>)

		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)

		COMMAND_ID_HANDLER(IDOK, OnCmdOk)
		COMMAND_ID_HANDLER(IDCANCEL, OnCmdCancel)
	END_MSG_MAP()

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wp, LPARAM lp, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnCmdOk(WORD, WORD, HWND, BOOL&);
	LRESULT OnCmdCancel(WORD, WORD, HWND, BOOL&);

private:
	HRESULT CommitLabel();
	void CloseWindow(BOOL bOk);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiChangeImageLabel), CVsiChangeImageLabel)
