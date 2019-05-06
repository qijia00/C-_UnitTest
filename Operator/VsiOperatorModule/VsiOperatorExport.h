/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorExport.h
**
**	Description:
**		Declaration of the CVsiOperatorExport
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiAppModule.h>
#include <VsiMessageFilter.h>
#include <VsiLogoViewBase.h>
#include <VsiAppViewModule.h>
#include <VsiOperatorModule.h>
#include <VsiImportExportModule.h>
#include <VsiPathNameEditWnd.h>
#include <VsiImportExportUIHelper.h>
#include "resource.h"


class ATL_NO_VTABLE CVsiOperatorExport :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiOperatorExport, &CLSID_VsiOperatorExport>,
	public CDialogImpl<CVsiOperatorExport>,
	public CVsiLogoViewBase<CVsiOperatorExport>,
	public CVsiImportExportUIHelper<CVsiOperatorExport>,
	public IVsiView,
	public IVsiPreTranslateMsgFilter,
	public IVsiOperatorExport
{
	friend CVsiImportExportUIHelper<CVsiOperatorExport>;

private:
	CComQIPtr<IVsiApp> m_pApp;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiCommandManager> m_pCmdMgr;

	std::vector<CString> m_vecExportFiles;

public:
	enum { IDD = IDD_COMMON_EXPORT };
	enum
	{
		WM_VSI_REFRESH_TREE = WM_USER + 100,
		WM_VSI_SET_SELECTION,
	};

	CVsiOperatorExport();
	~CVsiOperatorExport();

	DECLARE_NO_REGISTRY()

	BEGIN_COM_MAP(CVsiOperatorExport)
		COM_INTERFACE_ENTRY(IVsiOperatorExport)
		COM_INTERFACE_ENTRY(IVsiView)
		COM_INTERFACE_ENTRY(IVsiPreTranslateMsgFilter)
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
	// IVsiView
	STDMETHOD(Activate)(IVsiApp *pApp, HWND hwndParent);
	STDMETHOD(Deactivate)();
	STDMETHOD(GetWindow)(HWND *phWnd);
	STDMETHOD(GetAccelerator)(HACCEL *phAccel);
	STDMETHOD(GetIsBusy)(
		DWORD dwStateCurrent,
		DWORD dwState,
		BOOL bTryRelax,
		BOOL *pbBusy);

	// IVsiPreTranslateMsgFilter
	STDMETHOD(PreTranslateMessage)(MSG *pMsg, BOOL *pbHandled);

	// IVsiOperatorExport
	STDMETHOD(Initialize)(IUnknown *pPropBag);

protected:
	BEGIN_MSG_MAP(CVsiOperatorExport)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_VSI_REFRESH_TREE, OnRefreshTree)
		MESSAGE_HANDLER(WM_VSI_SET_SELECTION, OnSetSelection)

		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)

		NOTIFY_HANDLER(IDC_IE_FOLDER_TREE, TVN_SELCHANGED, OnTreeSelChanged);

		CHAIN_MSG_MAP(CVsiImportExportUIHelper<CVsiOperatorExport>)
		CHAIN_MSG_MAP(CVsiLogoViewBase<CVsiOperatorExport>)

		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnRefreshTree(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnSetSelection(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);

	LRESULT OnTreeSelChanged(int, LPNMHDR, BOOL&);

protected:
	// CVsiImportExportUIHelper
	BOOL GetAvailablePath(LPWSTR pszBuffer, int iBufferLength);
	int CheckForOverwrite();
	void UpdateUI();
};

OBJECT_ENTRY_AUTO(__uuidof(VsiOperatorExport), CVsiOperatorExport)
