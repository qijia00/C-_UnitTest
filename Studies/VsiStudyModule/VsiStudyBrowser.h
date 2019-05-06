/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudyBrowser.h
**
**	Description:
**		Declaration of the CVsiStudyBrowser
**
*******************************************************************************/

#pragma once

#include <atltime.h>
#include <ATLComTime.h>
#include <VsiEvents.h>
#include <VsiRes.h>
#include <VsiStateHandler.h>
#include <VsiLogoViewBase.h>
#include <VsiToolsMore.h>
#include <VsiCommUtlLib.h>
#include <VsiAppModule.h>
#include <VsiAppControllerModule.h>
#include <VsiAppViewModule.h>
#include <VsiMessageFilter.h>
#include <VsiOperatorModule.h>
#include <VsiStudyModule.h>
#include <VsiSelectOperator.h>
#include <VsiImagePreviewImpl.h>
#include <VsiStudyOverviewCtl.h>
#include "VsiSearchWnd.h"
#include "resource.h"


#define WM_VSI_SB_FILL_LIST				(WM_USER + 101)
#define WM_VSI_SB_FILL_LIST_DONE		(WM_USER + 102)
#define WM_VSI_SB_UPDATE_UI				(WM_USER + 103)
#define WM_VSI_SB_UPDATE_ITEMS			(WM_USER + 104)


class ATL_NO_VTABLE CVsiStudyBrowser :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiStudyBrowser, &CLSID_VsiStudyBrowser>,
	public CVsiLogoViewBase<CVsiStudyBrowser>,
	public CVsiThread<CVsiStudyBrowser>,
	public CVsiImagePreviewImpl<CVsiStudyBrowser>,
	public CDialogImpl<CVsiStudyBrowser>,
	public CVsiToolsMore<CVsiStudyBrowser>,
	public IVsiCommandSink,
	public IVsiViewEventSink,
	public IVsiEventSink,
	public IVsiParamUpdateEventSink,
	public IVsiAppStateSink,
	public IVsiView,
	public IVsiPreTranslateMsgFilter,
	public IVsiStudyBrowser
{
	friend CVsiThread<CVsiStudyBrowser>;
	friend CVsiImagePreviewImpl<CVsiStudyBrowser>;

private:
	CComPtr<IVsiApp> m_pApp;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiAppController> m_pAppController;
	CComPtr<IVsiAppView> m_pAppView;
	CComPtr<IVsiCommandManager> m_pCmdMgr;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;
	CComPtr<IVsiStudyManager> m_pStudyMgr;
	CComPtr<IVsiModeManager> m_pModeMgr;
	CComPtr<IVsiDataListWnd> m_pDataList;
	CComPtr<IVsiSession> m_pSession;
	CComPtr<IVsiCursorManager> m_pCursorMgr;

	CComPtr<IVsiWindowLayout> m_pLayout;

	class CVsiLockData
	{
	private:
		CComPtr<IVsiDataListWnd> m_pDataList;
		bool m_bLocked;

	public:
		CVsiLockData(IVsiDataListWnd *pDataListWnd) :
			m_pDataList(pDataListWnd),
			m_bLocked(false)
		{
		}
		~CVsiLockData()
		{
			UnlockData();
		}
		bool LockData()
		{
			if (!m_bLocked)
			{
				m_bLocked = S_OK == m_pDataList->LockData();
			}

			return m_bLocked;
		}
		void UnlockData()
		{
			if (m_bLocked)
			{
				m_pDataList->UnlockData();
				m_bLocked = false;
			}
		}
	};

	enum
	{
		VSI_LIST_EXIT = 0,
		VSI_LIST_RELOAD,
		VSI_LIST_SEARCH,
		VSI_LIST_MAX,
	};
	CHandle m_hExit;
	CHandle m_hReload;
	CHandle m_hSearch;

	CVsiSearchWnd m_wndSearch;
	CString m_strSearch;
	CCriticalSection m_csSearch;
	BOOL m_bStopSearch;

	int m_iTracker;
	int m_iPreview;
	int m_iBorder;
	bool m_bVertical;

	typedef enum
	{
		VSI_SB_STATE_START = 0,
		VSI_SB_STATE_CANCEL,
	} VSI_SB_STATE;

	VSI_SB_STATE m_State;

	// Thumbnail
	CComPtr<IUnknown> m_pItemPreview;
	VSI_DATA_LIST_TYPE m_typeSel;
	CComVariant m_vIdSelected;

	// Study overview
	CVsiStudyOverviewCtl m_studyOverview;
	
	CString m_strReactivateSeriesNs;  // Series to reactivate
	
	CString m_strSelection;		// Selection
	
	DWORD m_dwSink;				// PDM event sink

	// Operator List update event id
	DWORD m_dwEventImagePropUpdate;
	// Study path update event id
	DWORD m_dwStudiesPathUpdate;
	// Study access type update event id
	DWORD m_dwStudiesAccessFilter;
	// Operator list update event id
	DWORD m_dwOperatorListUpdate;
	// Secure mode
	DWORD m_dwSecureModeChanged;
	// Application mode
	VSI_APP_MODE_FLAGS m_dwAppMode;

	// List of items needed to update
	typedef std::list<CString> CVsiListUpdateItems;
	typedef CVsiListUpdateItems::iterator CVsiListUpdateItemsIter;
	CVsiListUpdateItems m_listUpdateItems;

	CCriticalSection m_csUpdateItems;

	CComPtr<IVsiChangeImageLabel> m_pImageLabel;

	CComPtr<IVsiAnalysisSet> m_pAnalysisSet;

	VSI_VIEW_STATE m_ViewState;

	CImageList m_ilStudyAccess;
	CImageList m_ilHelp;
	CImageList m_ilPreferences;
	CImageList m_ilShutdown;

public:
	enum { IDD = IDD_STUDY_BROWSER_VIEW };

	CVsiStudyBrowser();
	~CVsiStudyBrowser();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSISTUDYBROWSER)

	BEGIN_COM_MAP(CVsiStudyBrowser)
		COM_INTERFACE_ENTRY(IVsiView)
		COM_INTERFACE_ENTRY(IVsiStudyBrowser)
		COM_INTERFACE_ENTRY(IVsiPreTranslateMsgFilter)
		COM_INTERFACE_ENTRY(IVsiCommandSink)
		COM_INTERFACE_ENTRY(IVsiAppStateSink)
		COM_INTERFACE_ENTRY(IVsiEventSink)
		COM_INTERFACE_ENTRY(IVsiViewEventSink)
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

	// CDialogImplBaseT
	virtual void OnFinalMessage(HWND hWnd)
	{
		UNREFERENCED_PARAMETER(hWnd);
	}

	VSI_DECLARE_USE_MESSAGE_LOG()

public:
	// IVsiAppStateSink
	STDMETHOD(OnPreStateUpdate)(
		VSI_APP_STATE dwStateCurrent,
		VSI_APP_STATE dwState,
		BOOL bTryRelax,
		BOOL *pbBusy,
		const VARIANT *pvParam);
	STDMETHOD(OnStateUpdate)(VSI_APP_STATE dwState, const VARIANT *pvParam);

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

	// IVsiViewEventSink
	STDMETHOD(OnViewStateChange)(VSI_VIEW_STATE state);

	// IVsiPreTranslateMsgFilter
	STDMETHOD(PreTranslateMessage)(MSG *pMsg, BOOL *pbHandled);

	// IVsiStudyBrowser
	STDMETHOD(SetSelection)(LPCOLESTR pszSelection);
	STDMETHOD(GetLastSelectedItem)(VSI_DATA_LIST_TYPE *pType, IUnknown **ppUnk);
	STDMETHOD(DeleteImage)(IVsiImage *pImage);

	// IVsiCommandSink
	STDMETHOD(StateChanged)(DWORD)
	{
		return E_NOTIMPL;
	}
	STDMETHOD(Execute)(
		DWORD dwCmd,
		const VARIANT *pvParam,
		BOOL *pbContinueExecute);

	// IVsiEventSink
	STDMETHOD(OnEvent)(DWORD dwEvent, const VARIANT *pv);

	// IVsiParamUpdateEventSink
	STDMETHOD(OnParamUpdate)(DWORD *pdwParamIds, DWORD dwCount);

protected:
	BEGIN_MSG_MAP(CVsiStudyBrowser)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
		MESSAGE_HANDLER(WM_SIZE, OnSize)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
		MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)

		MESSAGE_HANDLER(WM_VSI_SB_FILL_LIST, OnFillList)
		MESSAGE_HANDLER(WM_VSI_SB_FILL_LIST_DONE, OnFillListDone)
		MESSAGE_HANDLER(WM_VSI_SB_UPDATE_UI, OnUpdateUI)
		MESSAGE_HANDLER(WM_VSI_SB_UPDATE_ITEMS, OnUpdateItems)

		COMMAND_HANDLER(IDC_SB_TOOLS_MORE, BN_CLICKED, OnBnClickedMore)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)
		COMMAND_HANDLER(IDC_SB_HELP, BN_CLICKED, OnBnClickedHelp)
		COMMAND_HANDLER(IDC_SB_PREFS, BN_CLICKED, OnBnClickedPrefs)
		COMMAND_HANDLER(IDC_SB_SHUTDOWN, BN_CLICKED, OnBnClickedShutdown)
		COMMAND_HANDLER(IDC_SB_CLOSE, BN_CLICKED, OnBnClickedClose)
		COMMAND_HANDLER(IDC_SB_STUDY_SERIES_INFO, BN_CLICKED, OnBnClickedInfo)
		COMMAND_HANDLER(IDC_SB_DELETE, BN_CLICKED, OnBnClickedDelete)
		COMMAND_HANDLER(IDC_SB_EXPORT, BN_CLICKED, OnBnClickedExport)
		COMMAND_HANDLER(IDC_SB_COPY_TO, BN_CLICKED, OnBnClickedCopyTo)
		COMMAND_HANDLER(IDC_SB_COPY_FROM, BN_CLICKED, OnBnClickedCopyFrom)
		COMMAND_HANDLER(IDC_SB_NEW, BN_CLICKED, OnBnClickedNew)
		COMMAND_HANDLER(IDC_SB_ANALYSIS, BN_CLICKED, OnBnClickedAnalysis)
		COMMAND_HANDLER(IDC_SB_STRAIN, BN_CLICKED, OnBnClickedStrain)
		COMMAND_HANDLER(IDC_SB_SONO, BN_CLICKED, OnBnClickedVevoCQ)
		COMMAND_HANDLER(IDC_SB_LOGOUT, BN_CLICKED, OnBnClickedLogout)
		COMMAND_HANDLER(IDC_SB_ACCESS_FILTER, BN_CLICKED, OnBnClickedStudyAccessFilter)
		COMMAND_HANDLER(IDC_SB_SEARCH, EN_CHANGE, OnEnChangedSearch)

		NOTIFY_HANDLER(IDC_SB_STUDY_LIST, LVN_ITEMCHANGED, OnListItemChanged)
		NOTIFY_HANDLER(IDC_SB_STUDY_LIST, LVN_ITEMACTIVATE, OnListItemActivate)
		NOTIFY_HANDLER(IDC_SB_STUDY_LIST, NM_CLICK, OnListItemClick)
		NOTIFY_HANDLER(IDC_SB_STUDY_LIST, NM_RCLICK, OnListItemRClick)
		NOTIFY_HANDLER(IDC_SB_STUDY_LIST, NM_VSI_DL_SORT_CHANGED, OnListSortChanged)
		NOTIFY_HANDLER(IDC_SB_STUDY_LIST, NM_VSI_DL_GET_ITEM_STATUS, OnListGetItemStatus);

		NOTIFY_HANDLER(IDC_SB_PREVIEW, NM_DBLCLK, OnPreviewListDbClick)
		NOTIFY_HANDLER(IDC_SB_PREVIEW, LVN_ITEMCHANGED, OnPreviewListItemChanged)
		NOTIFY_HANDLER(IDC_SB_PREVIEW, LVN_KEYDOWN, OnPreviewListKeyDown)
		NOTIFY_HANDLER(IDC_SB_PREVIEW, LVN_DELETEITEM, OnPreviewListDeleteItem)

		NOTIFY_HANDLER(IDC_SB_TRACK_BAR, NM_VSI_TRACKBAR_MOVE, OnTrackBarMove)

		NOTIFY_HANDLER(IDC_SB_ACCESS_FILTER, BCN_DROPDOWN, OnStudyAccessFilterDropDown)
		NOTIFY_HANDLER(IDC_SB_LOGOUT, BCN_DROPDOWN, OnLogoutDropDown)

		NOTIFY_CODE_HANDLER(HDN_ENDDRAG, OnHdnEndDrag)
		NOTIFY_CODE_HANDLER(HDN_ENDTRACK, OnHdnEndTrack)

		REFLECT_NOTIFICATIONS()

		CHAIN_MSG_MAP(CVsiImagePreviewImpl<CVsiStudyBrowser>)
		CHAIN_MSG_MAP(CVsiLogoViewBase<CVsiStudyBrowser>)
	END_MSG_MAP()

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wp, LPARAM lp, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnSettingChange(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnSize(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnWindowPosChanged(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnContextMenu(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnFillList(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnFillListDone(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnUpdateUI(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnUpdateItems(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnBnClickedMore(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedHelp(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedPrefs(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedShutdown(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedClose(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedInfo(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedDelete(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedExport(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCopyTo(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCopyFrom(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedNew(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedAnalysis(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedStrain(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedVevoCQ(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedLogout(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedStudyAccessFilter(WORD, WORD, HWND, BOOL&);
	LRESULT OnEnChangedSearch(WORD, WORD, HWND, BOOL&);

	LRESULT OnListItemChanged(int, LPNMHDR, BOOL&);
	LRESULT OnListItemActivate(int, LPNMHDR, BOOL&);
	LRESULT OnListItemClick(int, LPNMHDR, BOOL&);
	LRESULT OnListItemRClick(int, LPNMHDR, BOOL&);
	LRESULT OnListSortChanged(int, LPNMHDR, BOOL&);
	LRESULT OnListGetItemStatus(int, LPNMHDR, BOOL&);

	LRESULT OnPreviewListDbClick(int, LPNMHDR, BOOL&);
	LRESULT OnPreviewListItemChanged(int, LPNMHDR, BOOL&);
	LRESULT OnPreviewListKeyDown(int, LPNMHDR, BOOL&);
	LRESULT OnPreviewListDeleteItem(int, LPNMHDR, BOOL&);

	LRESULT OnTrackBarMove(int, LPNMHDR, BOOL&);

	LRESULT OnStudyAccessFilterDropDown(int, LPNMHDR, BOOL&);
	LRESULT OnLogoutDropDown(int, LPNMHDR, BOOL&);

	LRESULT OnHdnEndDrag(int, LPNMHDR, BOOL&);
	LRESULT OnHdnEndTrack(int, LPNMHDR, BOOL&);

	#define VSI_STATES() \
		VSI_STATE(VSI_APP_STATE_MODE_VIEW_CLOSE) \
		VSI_STATE(VSI_APP_STATE_MODE_CLOSE) \
		VSI_STATE(VSI_APP_STATE_MODE) \
		VSI_STATE(VSI_APP_STATE_SPLIT_SCREEN) \
		VSI_STATE(VSI_APP_STATE_IMAGE_LABEL) \
		VSI_STATE(VSI_APP_STATE_IMAGE_LABEL_CLOSE) \
		VSI_STATE(VSI_APP_STATE_STUDY_BROWSER) \
		VSI_STATE(VSI_APP_STATE_ANALYSIS_BROWSER_CLOSE) \
		// End of VSI_STATES

	BEGIN_VSI_STATE_MAP()
		#define VSI_STATE VSI_STATE_HANDLER
			VSI_STATES()
		#undef VSI_STATE
	END_VSI_STATE_MAP()

	#define VSI_STATE VSI_STATE_HANDLER_DEC
		VSI_STATES()
	#undef VSI_STATE
	#undef VSI_STATES

private:
	#define VSI_CMDS() \
		VSI_CMD(ID_CMD_UPDATE_ANALYSIS_BROWSER) \
		VSI_CMD(ID_CMD_STUDY_SERIES_INFO) \
		VSI_CMD(ID_CMD_DELETE) \
		VSI_CMD(ID_CMD_EXPORT) \
		VSI_CMD(ID_CMD_REPORT) \
		VSI_CMD(ID_CMD_COPY_STUDY_TO) \
	// End of VSI_CMDS

	BEGIN_VSI_CMD_MAP()
		#define VSI_CMD VSI_CMD_HANDLER
			VSI_CMDS()
		#undef VSI_CMD
	END_VSI_CMD_MAP()

	#define VSI_CMD VSI_CMD_HANDLER_DEC
		VSI_CMDS()
	#undef VSI_CMD

private:
	// CVsiThread
	DWORD ThreadProc();

private:
	// CVsiImagePreviewImpl
	void PreviewImages();

private:
	void CancelSearchIfRunning();
	void UpdateSelection() throw(...);
	void UpdateUI();
	void UpdateLogoutState() throw(...);
	void OpenItem();
	void Open(IVsiStudy *pStudy);
	void Open(IVsiSeries *pSeries);
	void Open(IVsiImage *pImage);
	HRESULT Delete(
		IVsiStudy *pStudy,
		IXMLDOMDocument *pXmlDoc,
		IXMLDOMElement *pElmRemove);
	HRESULT Delete(
		IVsiSeries *pSeries,
		IXMLDOMDocument *pXmlDoc,
		IXMLDOMElement *pElmRemove);
	HRESULT Delete(
		IVsiImage *pImage,
		IXMLDOMDocument *pXmlDoc,
		IXMLDOMElement *pElmRemove);
	void SetStudyInfo(IVsiStudy *pStudy, bool bUpdateUI = true);
	VSI_PROP_IMAGE GetImagePropIdFromColumnId(int iCol);
	VSI_PROP_SERIES GetSeriesPropIdFromColumnId(int iCol);
	BOOL CheckForOpenItems(
		const VSI_DATA_LIST_COLLECTION &listSel,
		int *piNumStudy,
		int *piNumSeries,
		int *piNumImage);
	void UpdateNewCloseButtons();

	void UpdateDataList(IXMLDOMDocument *pXmlDoc);
		
	void SetPreviewSize(int iSize);

	HRESULT GenerateFullDisclosureMsmntReport(BOOL bActivateReport = TRUE);
	HRESULT CountSonoVevoImages(int *piImages);
	HRESULT UpdateSonoVevoButton();
	HRESULT	UpdateStudyAccessFilterImage();
	HRESULT FillStudyList();
	HRESULT UpdateColumnsRecord();
	HRESULT	FixSaxCalc(IVsiEnumImage *pEnumImages);
	HRESULT DeleteEkvRaw(IVsiEnumImage *pEnumImageAll);
	void DisableSelectionUI();
};

OBJECT_ENTRY_AUTO(__uuidof(VsiStudyBrowser), CVsiStudyBrowser)
