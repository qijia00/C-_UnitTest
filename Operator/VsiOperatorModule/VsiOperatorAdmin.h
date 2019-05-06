/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorAdmin.h
**
**	Description:
**		Declaration of the CVsiOperatorAdmin
**
*******************************************************************************/

#pragma once

#include <IVsiPropertyPageImpl.h>
#include <VsiScrollImpl.h>
#include <VsiStl.h>
#include <VsiRes.h>
#include <VsiLogoViewBase.h>
#include <VsiAppViewModule.h>
#include <VsiStudyModule.h>
#include <VsiMessageFilter.h>
#include <VsiAppModule.h>
#include "resource.h"       // main symbols
#include "VsiOperatorModule.h"


class ATL_NO_VTABLE CVsiOperatorAdmin :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiOperatorAdmin, &CLSID_VsiOperatorAdmin>,
	public IVsiPropertyPageImpl<CVsiOperatorAdmin>,
	public CDialogImpl<CVsiOperatorAdmin>,
	public CVsiScrollImpl<CVsiOperatorAdmin>,
	public IVsiOperatorAdmin
{
private:
	CComQIPtr<IVsiApp> m_pApp;
	CComPtr<IVsiCommandManager> m_pCmdMgr;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;
	CComPtr<IVsiSession> m_pSession;

	bool m_bSecureMode;
	bool m_bSecureModeToChange;
	int m_iSortColumn;
	BOOL m_bSortDescending;
	CWindow m_wndOperatorList;

	CWindow m_wndUsageList;

public:
	enum { IDD = IDD_PREF_OPERATOR_ADMIN };
	enum
	{
		WM_VSI_REFRESH_OPERATOR_LIST = (WM_USER + 101)
	};

	CVsiOperatorAdmin();
	~CVsiOperatorAdmin();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSIOPERATORADMIN)

	BEGIN_COM_MAP(CVsiOperatorAdmin)
		COM_INTERFACE_ENTRY(IPropertyPage)
		COM_INTERFACE_ENTRY(IVsiOperatorAdmin)
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
	// IPropertyPage
	STDMETHOD(SetObjects)(ULONG nObjects, IUnknown **ppUnk);
	STDMETHOD(Apply)();

protected:
	BEGIN_MSG_MAP(CVsiOperatorAdmin)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_VSI_REFRESH_OPERATOR_LIST, OnRefreshOperatorList)

		COMMAND_HANDLER(IDC_OPERATOR_TURNON_SECURE_CHECK, BN_CLICKED, OnBnClickedSecure)
		COMMAND_HANDLER(IDC_OPERATOR_ADD, BN_CLICKED, OnBnClickedOperatorAdd)
		COMMAND_HANDLER(IDC_OPERATOR_DELETE, BN_CLICKED, OnBnClickedOperatorDelete)
		COMMAND_HANDLER(IDC_OPERATOR_MODIFY, BN_CLICKED, OnBnClickedOperatorModify)
		COMMAND_HANDLER(IDC_OPERATOR_EXPORT, BN_CLICKED, OnBnClickedOperatorExport)
		COMMAND_HANDLER(IDC_OPERATOR_IMPORT, BN_CLICKED, OnBnClickedOperatorImport)
		COMMAND_HANDLER(IDC_OPERATOR_TURNON_USAGE_CHECK, BN_CLICKED, OnBnClickedUsage)
		COMMAND_HANDLER(IDC_OPERATOR_USAGE_EXPORT, BN_CLICKED, OnBnClickedUsageExport)
		COMMAND_HANDLER(IDC_OPERATOR_USAGE_EXPORT, BN_CLICKED, OnBnClickedUsageClear)

		NOTIFY_HANDLER(IDC_OPERATOR_LIST, NM_CUSTOMDRAW, OnListCustomDraw)
		NOTIFY_HANDLER(IDC_OPERATOR_LIST, LVN_ITEMACTIVATE, OnListItemActivate)
		NOTIFY_HANDLER(IDC_OPERATOR_LIST, LVN_DELETEITEM, OnListDeleteItem)
		NOTIFY_HANDLER(IDC_OPERATOR_LIST, LVN_ITEMCHANGED, OnListItemChanged)
		NOTIFY_HANDLER(IDC_OPERATOR_LIST, LVN_COLUMNCLICK, OnListColumnClick)

		CHAIN_MSG_MAP(CVsiScrollImpl<CVsiOperatorAdmin>)
		CHAIN_MSG_MAP(IVsiPropertyPageImpl<CVsiOperatorAdmin>)

		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wp, LPARAM lp, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnRefreshOperatorList(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnBnClickedSecure(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedOperatorAdd(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedOperatorDelete(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedOperatorModify(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedOperatorExport(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedOperatorImport(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedUsage(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedUsageExport(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedUsageClear(WORD, WORD, HWND, BOOL&);

	LRESULT OnListCustomDraw(int, LPNMHDR, BOOL&);
	LRESULT OnListItemActivate(int, LPNMHDR, BOOL&);
	LRESULT OnListDeleteItem(int, LPNMHDR, BOOL&);
	LRESULT OnListItemChanged(int, LPNMHDR, BOOL&);
	LRESULT OnListColumnClick(int, LPNMHDR, BOOL&);

private:
	void InitOperatorListColumns();
	void RefreshOperators(LPCWSTR pszSelectId = NULL);
	void SelectOperator(LPCWSTR pszId);
	VSI_OPERATOR_TYPE GetLoginOperatorType();
	bool IsSelectedOperatorLogin();
	BOOL GetSelectedOperator(CString &strId);
	int AddOperator(IVsiOperator *pOperator);
	void UpdateOperator(int iIndex, IVsiOperator *pOperator);
	static int CALLBACK CompareOperatorProc(LPARAM lParam1, LPARAM lParam2, LPARAM lparamSort);
	void SortOperatorList();
	void SetOperatorSortIndicator(int iCol, int iImage);
	HRESULT UpdatePackageName(LPCWSTR strFile, LPCWSTR OpName);
	HRESULT ProcessCustomPackagesPresets();
	HRESULT SaveOperatorsSettings();
	void RefreshUsageLog();
};

OBJECT_ENTRY_AUTO(__uuidof(VsiOperatorAdmin), CVsiOperatorAdmin)
