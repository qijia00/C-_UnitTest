/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudyInfo.h
**
**	Description:
**		Declaration of the CVsiStudyInfo
**
*******************************************************************************/

#pragma once

#include <VsiRes.h>
#include <VsiCommCtlLib.h>
#include <VsiScrollImpl.h>
#include <VsiAppModule.h>
#include <VsiParameterModule.h>
#include <VsiLogoViewBase.h>
#include <VsiMessageFilter.h>
#include <VsiAppViewModule.h>
#include <VsiOperatorModule.h>
#include <VsiCommCtl.h>
#include <VsiScanFormatModule.h>
#include <VsiMeasurementModule.h>
#include <VsiSelectOperator.h>
#include <VsiStudyModule.h>
#include <VsiDateEditWnd.h>
#include <VsiStudyAccessCombo.h>
#include "resource.h"


class ATL_NO_VTABLE CVsiStudyInfo :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiStudyInfo, &CLSID_VsiStudyInfo>,
	public CDialogImpl<CVsiStudyInfo>,
	public CVsiScrollImpl<CVsiStudyInfo>,
	public CVsiLogoViewBase<CVsiStudyInfo, CVsiScrollImpl<CVsiStudyInfo>>,
	public IVsiViewEventSink,
	public IVsiParamUpdateEventSink,
	public IVsiView,
	public IVsiPreTranslateMsgFilter,
	public IVsiCommandSink,
	public IVsiStudyInfo
{
private:
	CComPtr<IVsiApp> m_pApp;
	CComPtr<IVsiCommandManager> m_pCmdMgr;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;
	CComPtr<IVsiStudyManager> m_pStudyMgr;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiStudy> m_pStudy;
	CComPtr<IVsiSeries> m_pSeries;
	CComPtr<IVsiSession> m_pSession;
	CComPtr<IVsiScanFormatManager> m_pScanFormatMgr;
	CComPtr<IVsiScanFormat> m_pScanFormat;
	CComPtr<IVsiMsmntManager> m_pMsmntMgr;
	CComPtr<IVsiCursorManager> m_pCursorMgr;

	DWORD m_dwSink;  // PDM event sink

	VSI_VIEW_STATE m_ViewState;

	VSI_SI_STATE m_state;
	BOOL m_bClosed;
	BOOL m_bSecureMode;
	BOOL m_bAdmin;
	bool m_bStudyChanged;
	bool m_bSeriesChanged;

	CVsiStudyAccessCombo m_studyAccess;

	// Operator
	CVsiSelectOperator m_wndOperatorOwner;
	BOOL m_bAllowChangingOperatorOwner;
	CVsiSelectOperator m_wndOperatorAcq;
	BOOL m_bAllowChangingOperatorAcq;

	// Operator List update event id
	DWORD m_dwEventOperatorListUpdate;
	DWORD m_dwEventSysState;

	SYSTEMTIME m_stDateOfBirth;
	SYSTEMTIME m_stDateMated;
	SYSTEMTIME m_stDatePlugged;

	CComPtr<IVsiDatePickerWnd> m_pDatePicker;
	HIMAGELIST m_hilDate;

	CString m_strOwnerOperator;
	CString m_strAcquireByOperator;
	
	HBRUSH m_hbrushMandatory;

	CVsiDateEditWnd m_wndDateOfBirth;
	CVsiDateEditWnd m_wndDateMated;
	CVsiDateEditWnd m_wndDatePlugged;

public:
	enum { IDD = IDD_STUDY_INFORMATION };

	CVsiStudyInfo();
	~CVsiStudyInfo();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSISTUDYINFO)

	BEGIN_COM_MAP(CVsiStudyInfo)
		COM_INTERFACE_ENTRY(IVsiView)
		COM_INTERFACE_ENTRY(IVsiStudyInfo)
		COM_INTERFACE_ENTRY(IVsiPreTranslateMsgFilter)
		COM_INTERFACE_ENTRY(IVsiParamUpdateEventSink)
		COM_INTERFACE_ENTRY(IVsiCommandSink)
		COM_INTERFACE_ENTRY(IVsiViewEventSink)
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

	// IVsiStudyInfo
	STDMETHOD(SetState)(VSI_SI_STATE state);
	STDMETHOD(SetStudy)(IVsiStudy *pStudy);
	STDMETHOD(SetSeries)(IVsiSeries *pSeries);

	// IVsiViewEventSink
	STDMETHOD(OnViewStateChange)(VSI_VIEW_STATE state);

	// IVsiPreTranslateMsgFilter
	STDMETHOD(PreTranslateMessage)(MSG *pMsg, BOOL *pbHandled);

	// IVsiParamUpdateEventSink
	STDMETHOD(OnParamUpdate)(DWORD *pdwParamIds, DWORD dwCount);

	// IVsiCommandSink
	STDMETHOD(StateChanged)(DWORD)
	{
		return E_NOTIMPL;
	}
	STDMETHOD(Execute)(
		DWORD dwCmd,
		const VARIANT *pvParam,
		BOOL *pbContinueExecute);

protected:
	BEGIN_MSG_MAP(CVsiStudyInfo)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_CTLCOLOREDIT, OnCtlColorEdit)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)

		COMMAND_HANDLER(IDC_SI_COPY_PREV, BN_CLICKED, OnBnClickedCopyPrev)
		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)

		CHAIN_MSG_MAP_MEMBER(m_wndOperatorOwner)
		CHAIN_MSG_MAP_MEMBER(m_wndOperatorAcq)

		COMMAND_HANDLER(IDC_SI_OPERATOR_NAME, LBN_SELCHANGE, OnOperatorOwnerChanged)
		COMMAND_HANDLER(IDC_SI_ACQ_OPERATOR_NAME, LBN_SELCHANGE, OnOperatorAcqChanged)

		COMMAND_HANDLER(IDC_SI_STUDY_NAME, EN_CHANGE, OnEnStudyChange)
		COMMAND_HANDLER(IDC_SI_GRANTING_INSTITUTION, EN_CHANGE, OnEnStudyChange)
		COMMAND_HANDLER(IDC_SI_STUDY_ACCESS_PRIVATE, EN_CHANGE, OnEnStudyChange)
		COMMAND_HANDLER(IDC_SI_STUDY_ACCESS_PUBLIC, EN_CHANGE, OnEnStudyChange)
		COMMAND_HANDLER(IDC_SI_STUDY_NOTES, EN_CHANGE, OnEnStudyChange)
		COMMAND_HANDLER(IDC_SI_SERIES_NAME, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_APPLICATION, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_PACKAGE_SELECTION, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_ANIMAL_ID, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_TYPE, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_COLOR, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_HERITAGE, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_SOURCE, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_WEIGHT, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_HEART_RATE, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_TEMPERATURE, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_SERIES_NOTES, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C1, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C2, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C3, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C4, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C5, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C6, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C7, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C8, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C9, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C10, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C11, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C12, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C13, EN_CHANGE, OnEnSeriesChange)
		COMMAND_HANDLER(IDC_SI_C14, EN_CHANGE, OnEnSeriesChange)

		COMMAND_ID_HANDLER(IDC_SI_MALE, OnCmdMale)
		COMMAND_ID_HANDLER(IDC_SI_FEMALE, OnCmdFemale)
		COMMAND_ID_HANDLER(IDC_SI_PREGNANT, OnCmdPregnant)
		COMMAND_ID_HANDLER(IDC_SI_DATE_OF_BIRTH_PICKER, OnCmdDateOfBirthPicker)
		COMMAND_ID_HANDLER(IDC_SI_DATE_MATED_PICKER, OnCmdDateMatedPicker)
		COMMAND_ID_HANDLER(IDC_SI_DATE_PLUGGED_PICKER, OnCmdDatePluggedPicker)

		CHAIN_MSG_MAP(CVsiScrollImpl)
		CHAIN_MSG_MAP(CVsiLogoViewBase)

	END_MSG_MAP()

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnCtlColorEdit(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnWindowPosChanged(UINT, WPARAM, LPARAM lp, BOOL &bHandled);
	LRESULT OnBnClickedCopyPrev(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);
	LRESULT OnOperatorOwnerChanged(WORD, WORD, HWND, BOOL&);
	LRESULT OnOperatorAcqChanged(WORD, WORD, HWND, BOOL&);
	LRESULT OnEnStudyChange(WORD, WORD, HWND, BOOL&);
	LRESULT OnEnSeriesChange(WORD, WORD, HWND, BOOL&);
	LRESULT OnCmdMale(WORD, WORD, HWND, BOOL&);
	LRESULT OnCmdFemale(WORD, WORD, HWND, BOOL&);
	LRESULT OnCmdPregnant(WORD, WORD, HWND, BOOL&);
	LRESULT OnCmdDateOfBirthPicker(WORD, WORD, HWND, BOOL&);
	LRESULT OnCmdDateMatedPicker(WORD, WORD, HWND, BOOL&);
	LRESULT OnCmdDatePluggedPicker(WORD, WORD, HWND, BOOL&);

private:
	#define VSI_CMDS() \
		VSI_CMD(ID_CMD_STUDY_SERIES_INFO)
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
	void Initialize() throw(...);
	void InitializeSeriesInfo(IVsiSeries *pSeries, bool bSkipName = false) throw(...);
	void UpdateUI();
	void UpdateSexUI(BOOL bFemale);
	void AddUserFields();
	void FillOperatorList();
	void GetOwnerOperator(CString &strOperator) throw(...);
	void GetAcquireByOperator(CString &strOperator) throw(...);
	void GetMsmntPackage(CString &strPackage) throw(...);
	void GetApplication(CString &strApplication) throw(...);
	void SaveChanges(IVsiStudy *pStudy, IVsiSeries *pSeries) throw(...);
	void ShowSeriesControls(BOOL bShow);
	void CreateNewSeries() throw(...);
	void CreateNewStudy() throw(...);
	void InitializeDateFormat();
	HRESULT ScanformatChange();
	HRESULT PopulatePackageSelection(void);
	HRESULT CheckDates(WORD *pwCtl, BOOL *pbValid);
	void UpdateLastModified();
	bool SelectedStudyActive();
};

OBJECT_ENTRY_AUTO(__uuidof(VsiStudyInfo), CVsiStudyInfo)
