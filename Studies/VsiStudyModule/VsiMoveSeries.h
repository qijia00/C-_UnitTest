/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiMoveSeries.h
**
**	Description:
**		Declaration of CVsiMoveSeries
**
********************************************************************************/

#pragma once
#include <VsiRes.h>
#include <VsiAppModule.h>
#include <VsiLogoViewBase.h>
#include <VsiAppViewModule.h>
#include <VsiMessageFilter.h>
#include <VsiPdmModule.h>
#include <VsiStudyModule.h>
#include <VsiStudyOverviewCtl.h>
#include "resource.h"       // main symbols


#define WM_VSI_SB_FILL_LIST		(WM_USER + 100)
#define WM_VSI_SET_SELECTION	(WM_USER + 101)
#define WM_VSI_UPDATE_PREVIEW	(WM_USER + 102)


class ATL_NO_VTABLE CVsiMoveSeries :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiMoveSeries, &CLSID_VsiMoveSeries>,
	public CDialogImpl<CVsiMoveSeries>,
	public CVsiLogoViewBase<CVsiMoveSeries>,
	public IVsiView,
	public IVsiMoveSeries
{
private:
	CComPtr<IVsiApp> m_pApp;
	CComPtr<IVsiCommandManager> m_pCmdMgr;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiStudyManager> m_pStudyMgr;
	CComPtr<IVsiSession> m_pSession;

	CComPtr<IVsiSeries> m_pSeries;

	CComPtr<IVsiDataListWnd> m_pDataList;
	CComPtr<IUnknown> m_pItemPreview;

	// Study overview
	CVsiStudyOverviewCtl m_studyOverview;

	int   m_iSelectedStudies;

public:
	enum { IDD = IDD_MOVE_SERIES_VIEW };

	CVsiMoveSeries();
	~CVsiMoveSeries();


	DECLARE_REGISTRY_RESOURCEID(IDR_VSIMOVESERIES)

	BEGIN_COM_MAP(CVsiMoveSeries)
		COM_INTERFACE_ENTRY(IVsiMoveSeries)
		COM_INTERFACE_ENTRY(IVsiView)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

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

	// IVsiMoveSeries
	STDMETHOD(SetSeries)(IVsiSeries *pSeries);

protected:
	BEGIN_MSG_MAP(CVsiMoveSeries)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_VSI_SB_FILL_LIST, OnFillList)
		MESSAGE_HANDLER(WM_VSI_UPDATE_PREVIEW, OnUpdatePreview)

		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)

		NOTIFY_HANDLER(IDC_MS_STUDY_LIST, LVN_ITEMCHANGED, OnItemChanged)
		NOTIFY_HANDLER(IDC_MS_STUDY_LIST, NM_VSI_DL_GET_ITEM_STATUS, OnListGetItemStatus);

		CHAIN_MSG_MAP(CVsiLogoViewBase<CVsiMoveSeries>)

		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM wParam, LPARAM, BOOL&);
	LRESULT OnSetSelection(UINT uMsg, WPARAM, LPARAM, BOOL&);
	LRESULT OnUpdatePreview(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnFillList(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);

	LRESULT OnItemChanged(int, LPNMHDR, BOOL &bHandled);
	LRESULT OnListGetItemStatus(int, LPNMHDR, BOOL&);

private:
	void UpdateUI();
	void SetStudyInfo(IVsiStudy *pStudy);
	HRESULT FillDataList();
};

OBJECT_ENTRY_AUTO(__uuidof(VsiMoveSeries), CVsiMoveSeries)
