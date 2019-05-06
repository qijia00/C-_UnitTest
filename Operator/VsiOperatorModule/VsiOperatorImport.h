/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorImport.h
**
**	Description:
**		Declaration of the CVsiOperatorImport
**
*******************************************************************************/

#pragma once


#include <VsiRes.h>
#include <VsiAppModule.h>
#include <VsiAppControllerModule.h>
#include <VsiLogoViewBase.h>
#include <VsiAppViewModule.h>
#include <VsiOperatorModule.h>
#include <VsiMessageFilter.h>
#include <VsiImportExportModule.h>
#include <VsiImportExportUIHelper.h>
#include "resource.h"


class ATL_NO_VTABLE CVsiOperatorImport :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiOperatorImport, &CLSID_VsiOperatorImport>,
	public CDialogImpl<CVsiOperatorImport>,
	public CVsiLogoViewBase<CVsiOperatorImport>,
	public CVsiImportExportUIHelper<CVsiOperatorImport>,
	public IVsiView,
	public IVsiPreTranslateMsgFilter,
	public IVsiOperatorImport
{
	friend CVsiImportExportUIHelper<CVsiOperatorImport>;

private:
	CComPtr<IVsiApp> m_pApp;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;
	CComPtr<IVsiCommandManager> m_pCmdMgr;

	CWindow m_wndRefresh;
	int m_msgRefresh;

public:
	enum { IDD = IDD_COMMON_IMPORT };
	enum
	{
		WM_VSI_REFRESH_TREE = WM_USER + 100,
		WM_VSI_SET_SELECTION,
	};

	CVsiOperatorImport();
	~CVsiOperatorImport();

	DECLARE_NO_REGISTRY()

	BEGIN_COM_MAP(CVsiOperatorImport)
		COM_INTERFACE_ENTRY(IVsiOperatorImport)
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

	// IVsiOperatorImport
	STDMETHOD(Initialize)(IUnknown *pPropBag);

protected:
	BEGIN_MSG_MAP(CVsiOperatorImport)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_VSI_REFRESH_TREE, OnRefreshTree)
		MESSAGE_HANDLER(WM_VSI_SET_SELECTION, OnSetSelection)

		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)

		NOTIFY_HANDLER(IDC_IE_FOLDER_TREE, TVN_SELCHANGED, OnTreeSelChanged);
		NOTIFY_HANDLER(IDC_IE_FOLDER_TREE, NM_VSI_FOLDER_GET_ITEM_INFO, OnCheckForBackup);

		CHAIN_MSG_MAP(CVsiImportExportUIHelper<CVsiOperatorImport>)
		CHAIN_MSG_MAP(CVsiLogoViewBase<CVsiOperatorImport>)

		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnRefreshTree(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnSetSelection(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);

	LRESULT OnTreeSelChanged(int, LPNMHDR, BOOL &bHandled);

	LRESULT OnCheckForBackup(int, LPNMHDR, BOOL &bHandled);

protected:
	// CVsiImportExportUIHelper
	BOOL GetAvailablePath(LPWSTR pszBuffer, int iBufferLength);
	int CheckForOverwrite();
	void UpdateUI();
	bool IsBackup(LPCWSTR pszPath);

	typedef struct tagVSI_OPERATOR_BACKUP_INFO
	{
		CString strName;
		SYSTEMTIME stBackup;
	} VSI_OPERATOR_BACKUP_INFO;
	HRESULT GetBackupInfo(LPCWSTR pszPath, VSI_OPERATOR_BACKUP_INFO *pInfo);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiOperatorImport), CVsiOperatorImport)
