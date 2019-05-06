/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiCopyStudyFrom.h
**
**	Description:
**		Declaration of the CVsiCopyStudyFrom
**
*******************************************************************************/

#pragma once


#include <VsiRes.h>
#include <VsiAppModule.h>
#include <VsiLogoViewBase.h>
#include <VsiAppViewModule.h>
#include <VsiMessageFilter.h>
#include <VsiImportExportModule.h>
#include <VsiImportExportUIHelper.h>
#include <VsiOperatorModule.h>
#include <VsiSelectOperator.h>
#include <VsiImagePreviewImpl.h>
#include <VsiStudyOverviewCtl.h>
#include "resource.h"       // main symbols


class ATL_NO_VTABLE CVsiCopyStudyFrom :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiCopyStudyFrom, &CLSID_VsiCopyStudyFrom>,
	public CDialogImpl<CVsiCopyStudyFrom>,
	public CVsiLogoViewBase<CVsiCopyStudyFrom>,
	public CVsiImagePreviewImpl<CVsiCopyStudyFrom>,
	public CVsiImportExportUIHelper<CVsiCopyStudyFrom>,
	public IVsiParamUpdateEventSink,
	public IVsiView,
	public IVsiPreTranslateMsgFilter,
	public IVsiCopyStudyFrom
{
	friend CVsiImagePreviewImpl<CVsiCopyStudyFrom>;
	friend CVsiImportExportUIHelper<CVsiCopyStudyFrom>;

private:
	enum
	{
		WM_VSI_REFRESH_TREE		= WM_USER + 100,
		WM_VSI_SET_SELECTION	= WM_USER + 101,
		WM_VSI_UPDATE_PREVIEW	= WM_USER + 102,
	};

	enum
	{
		VSI_CHECK_STATE_UNCHECKED = 0,
		VSI_CHECK_STATE_CHECKED = 1,
		VSI_CHECK_STATE_PARTIAL = 2,
	};

	CComPtr<IVsiApp> m_pApp;
	CComPtr<IVsiCommandManager> m_pCmdMgr;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;
	CComPtr<IVsiStudyManager> m_pStudyMgr;
	CComPtr<IVsiCursorManager> m_pCursorMgr;

	DWORD m_dwSink;  // PDM event sink

	// Operator
	CVsiSelectOperator m_wndOperator;

	// Operator List update event id
	DWORD m_dwEventOperatorListUpdate;
	CString m_strOwnerOperator;
	
	CString m_strLastAddPath;

	// Vector for study list management
	typedef std::set<CString> CVsiPathSet;
	CVsiPathSet m_setSelectedStudy;  // paths to import
	CVsiPathSet m_setRemovedSeries;  // series paths to exclude

	typedef std::set<CString> CVsiStudyIdSet;
	CVsiStudyIdSet m_setSelectedStudyId;

	CComPtr<IVsiDataListWnd> m_pDataList;
	CComPtr<IUnknown> m_pItemPreview;
	VSI_DATA_LIST_TYPE m_typeSel;

	// Study overview
	CVsiStudyOverviewCtl m_studyOverview;

	HBRUSH m_hbrushMandatory;

public:
	enum { IDD = IDD_COPY_STUDY_FROM_VIEW };

	CVsiCopyStudyFrom();
	~CVsiCopyStudyFrom();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSICOPYSTUDYFROM)

	BEGIN_COM_MAP(CVsiCopyStudyFrom)
		COM_INTERFACE_ENTRY(IVsiCopyStudyFrom)
		COM_INTERFACE_ENTRY(IVsiView)
		COM_INTERFACE_ENTRY(IVsiPreTranslateMsgFilter)
		COM_INTERFACE_ENTRY(IVsiParamUpdateEventSink)
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

	// IVsiParamUpdateEventSink
	STDMETHOD(OnParamUpdate)(DWORD *pdwParamIds, DWORD dwCount);

protected:
	BEGIN_MSG_MAP(CVsiCopyStudyFrom)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
		MESSAGE_HANDLER(WM_CTLCOLOREDIT, OnCtlColorEdit)
		MESSAGE_HANDLER(WM_VSI_REFRESH_TREE, OnRefreshTree)
		MESSAGE_HANDLER(WM_VSI_SET_SELECTION, OnSetSelection)
		MESSAGE_HANDLER(WM_VSI_UPDATE_PREVIEW, OnUpdatePreview)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)

		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)
		COMMAND_HANDLER(IDC_CSF_REMOVE, BN_CLICKED, OnBnClickedRemove)

		COMMAND_HANDLER(IDC_CF_OPERATOR_LIST, CBN_SELENDOK, OnOperatorSelenOk)

		NOTIFY_HANDLER(IDC_IE_FOLDER_TREE, NM_CLICK, OnTreeClick);
		NOTIFY_HANDLER(IDC_IE_FOLDER_TREE, TVN_SELCHANGED, OnTreeSelChanged);
		NOTIFY_HANDLER(IDC_IE_FOLDER_TREE, TVN_GETINFOTIP, OnTreeGetInfoTip);
		NOTIFY_HANDLER(IDC_IE_FOLDER_TREE, NM_VSI_FOLDER_GET_ITEM_INFO, OnCheckForStudySeries);

		NOTIFY_HANDLER(IDC_CSF_STUDY_LIST, LVN_ITEMCHANGED, OnItemChanged)
		NOTIFY_HANDLER(IDC_CSF_STUDY_LIST, NM_VSI_DL_GET_ITEM_STATUS, OnCheckItemStatus);

		NOTIFY_CODE_HANDLER(NM_VSI_SELECTION_CHANGED, OnSelChanged);

		CHAIN_MSG_MAP_MEMBER(m_wndOperator)

		CHAIN_MSG_MAP(CVsiImportExportUIHelper<CVsiCopyStudyFrom>)
		CHAIN_MSG_MAP(CVsiImagePreviewImpl<CVsiCopyStudyFrom>)
		CHAIN_MSG_MAP(CVsiLogoViewBase<CVsiCopyStudyFrom>)

		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM wParam, LPARAM, BOOL&);
	LRESULT OnSettingChange(UINT, WPARAM wParam, LPARAM, BOOL&);
	LRESULT OnCtlColorEdit(UINT uMsg, WPARAM, LPARAM, BOOL&);
	LRESULT OnRefreshTree(UINT uMsg, WPARAM, LPARAM, BOOL&);
	LRESULT OnSetSelection(UINT uMsg, WPARAM, LPARAM, BOOL&);
	LRESULT OnUpdatePreview(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnWindowPosChanged(UINT, WPARAM, LPARAM lp, BOOL &bHandled);

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedRemove(WORD, WORD, HWND, BOOL&);

	LRESULT OnOperatorSelenOk(WORD, WORD, HWND, BOOL&);

	LRESULT OnTreeClick(int, LPNMHDR, BOOL &bHandled);
	LRESULT OnTreeSelChanged(int, LPNMHDR, BOOL &bHandled);
	LRESULT OnTreeGetInfoTip(int, LPNMHDR, BOOL &bHandled);

	LRESULT OnItemChanged(int, LPNMHDR, BOOL &bHandled);
	LRESULT OnCheckForStudySeries(int, LPNMHDR, BOOL &bHandled);
	LRESULT OnCheckItemStatus(int, LPNMHDR, BOOL &bHandled);
	LRESULT OnDeleteItem(int, LPNMHDR, BOOL &bHandled);

	LRESULT OnSelChanged(int, LPNMHDR, BOOL &bHandled);

protected:
	// CVsiImportExportUIHelper
	int GetNumberOfSelectedItems();
	BOOL GetAvailablePath(LPWSTR pszBuffer, int iBufferLength);
	int CheckForOverwrite();
	void UpdateUI();

private:
	HRESULT AddSelectedItems();
	HRESULT RemoveSelectedItems();
	BOOL IsFolderHasStudy();
	BOOL IsStudyAdded(LPCWSTR pszFilePath);
	BOOL IsStudyFullyAdded(LPCWSTR pszFilePath);
	BOOL IsSeriesAdded(LPCWSTR pszFilePath);
	void SetStudyInfo(IVsiStudy *pStudy, bool bUpdateUI = true);
	VSI_PROP_SERIES GetSeriesPropIdFromColumnId(int iCol);
	void UpdateFEStudyDetails();
	void PreviewImages();
	void CalculateRequiredSize();
};

OBJECT_ENTRY_AUTO(__uuidof(VsiCopyStudyFrom), CVsiCopyStudyFrom)
