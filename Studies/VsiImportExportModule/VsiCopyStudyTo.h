/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiCopyStudyTo.h
**
**	Description:
**		Declaration of the CVsiCopyStudyTo
**
*******************************************************************************/

#pragma once

#include <shldisp.h>
#include <VsiRes.h>
#include <VsiAppModule.h>
#include <VsiMessageFilter.h>
#include <VsiLogoViewBase.h>
#include <VsiAppViewModule.h>
#include <VsiImportExportModule.h>
#include <VsiStudyModule.h>
#include <VsiCdDvdWriterModule.h>
#include <VsiPathNameEditWnd.h>
#include <VsiImportExportUIHelper.h>
#include "resource.h"       // main symbols


#define WM_VSI_REFRESH_TREE		(WM_USER + 100)
#define WM_VSI_SET_SELECTION	(WM_USER + 101)


class ATL_NO_VTABLE CVsiCopyStudyTo :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiCopyStudyTo, &CLSID_VsiCopyStudyTo>,
	public CDialogImpl<CVsiCopyStudyTo>,
	public CVsiLogoViewBase<CVsiCopyStudyTo>,
	public CVsiImportExportUIHelper<CVsiCopyStudyTo>,
	public IVsiView,
	public IVsiPreTranslateMsgFilter,
	public IVsiCopyStudyTo
{
	friend CVsiImportExportUIHelper<CVsiCopyStudyTo>;

private:
	CComQIPtr<IVsiApp> m_pApp;
	CComPtr<IVsiCommandManager> m_pCmdMgr;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiStudyManager> m_pStudyMgr;
	CComPtr<IVsiCursorManager> m_pCursorMgr;
	CComQIPtr<IXMLDOMDocument> m_pXmlDoc;

	CComPtr<IAutoComplete> m_pAutoComplete;
	CComPtr<IVsiAutoCompleteSource> m_pAutoCompleteSource;

	int m_iStudyCount;

	DWORD m_dwFlags;
	CVsiPathNameEditWnd m_wndStudyName;

	// CD/DVD device testing
	CComPtr<IVsiCdDvdWriter> m_pWriter;
	BOOL m_bWriterSelected;
	CString m_strDriveLetter;

public:
	enum { IDD = IDD_COPY_STUDY_TO_VIEW };

	CVsiCopyStudyTo();
	~CVsiCopyStudyTo();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSICOPYSTUDYTO)

	BEGIN_COM_MAP(CVsiCopyStudyTo)
		COM_INTERFACE_ENTRY(IVsiCopyStudyTo)
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

	// IVsiCopyStudyTo
	STDMETHOD(Initialize)(const VARIANT *pvParam);

protected:
	BEGIN_MSG_MAP(CVsiCopyStudyTo)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_VSI_REFRESH_TREE, OnRefreshTree)
		MESSAGE_HANDLER(WM_VSI_SET_SELECTION, OnSetSelection)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)

		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)

		COMMAND_HANDLER(IDC_IE_SAVEAS, EN_CHANGE, OnStudyNameChanged)

		NOTIFY_HANDLER(IDC_IE_FOLDER_TREE, TVN_SELCHANGED, OnTreeSelChanged);
		NOTIFY_HANDLER(IDC_IE_FOLDER_TREE, NM_VSI_CHECK_ALLOW_CHILDREN, OnCheckAllowChildren);

		CHAIN_MSG_MAP(CVsiImportExportUIHelper<CVsiCopyStudyTo>)
		CHAIN_MSG_MAP(CVsiLogoViewBase<CVsiCopyStudyTo>)

		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnRefreshTree(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnSetSelection(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnWindowPosChanged(UINT, WPARAM, LPARAM lp, BOOL &bHandled);

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);

	LRESULT OnStudyNameChanged(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

	LRESULT OnTreeSelChanged(int, LPNMHDR, BOOL&);
	LRESULT OnCheckAllowChildren(int, LPNMHDR, BOOL&);

protected:
	// CVsiImportExportUIHelper
	BOOL GetAvailablePath(LPWSTR pszBuffer, int iBufferLength);
	int CheckForOverwrite();
	void UpdateUI();
	void UpdateAvailableSpace();
	BOOL UpdateSizeUI();

private:
	int GetNumberOfSelectedItems();
	BOOL WouldOverwrite(IXMLDOMElement *pXmlElm, LPCWSTR pszTargetFolderName);
	HRESULT UpdateJobXml(BOOL bOverwrite);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiCopyStudyTo), CVsiCopyStudyTo)
