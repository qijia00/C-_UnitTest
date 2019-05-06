/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiCdDvdWriterView.h
**
**	Description:
**		Declaration of the CVsiCdDvdWriterView
**
*******************************************************************************/

#pragma once


#include <VsiRes.h>
#include <VsiEvents.h>
#include <VsiLogoViewBase.h>
#include <VsiMessageFilter.h>
#include <VsiAppViewModule.h>
#include <VsiImportExportModule.h>
#include <VsiCdDvdWriterModule.h>
#include <VsiShellHandlerImpl.h>
#include <VsiPathNameEditWnd.h>
#include "VsiImpExpUtility.h"
#include "resource.h"       // main symbols


#define WM_VSI_PROGRESS_WND_CLOSED	(WM_USER + 8)
#define WM_VSI_UPDATE_UI			(WM_USER + 9)
#define WM_VSI_WRITE_CANCELLED		(WM_USER + 10)
#define WM_VSI_START_WRITE			(WM_USER + 11)
#define WM_VSI_WRITE_ABORTED		(WM_USER + 12)


// CVsiCdDvdWriterView

class ATL_NO_VTABLE CVsiCdDvdWriterView :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiCdDvdWriterView, &CLSID_VsiCdDvdWriterView>,
	public CDialogImpl<CVsiCdDvdWriterView>,
	public CVsiLogoViewBase<CVsiCdDvdWriterView>,
	public CVsiShellHandlerImpl<CVsiCdDvdWriterView>,
	public IVsiView,
	public IVsiPreTranslateMsgFilter,
	public IVsiCdDvdWriterView,
	public IVsiEventSink
{
private:

	CComQIPtr<IVsiApp> m_pApp;
	CComPtr<IVsiCommandManager> m_pCmdMgr;
	CComQIPtr<IVsiServiceProvider> m_pServiceProvider;

	CComQIPtr<IXMLDOMDocument> m_pXmlDoc;
	CString m_strDriveLetter;

	HWND m_hwndProgress;
	BOOL m_bCancelWrite;  // Set this TRUE to abort the writing

	//timer for writing progress
	UINT_PTR m_uiTimerProgress;
	DWORD m_dwBlocksToWrite;
	DWORD m_dwBlocksWritten;

	double m_dbRequiredSizeMB;

	DWORD m_dwSink;  // CD/DVD writer event sink

	CComQIPtr<IVsiCdDvdWriter> m_pWriter;
	BOOL m_bIdle;
	BOOL m_bErasable;
//	BOOL m_bAddDataLater;
	BOOL m_bVerifyData;
	BOOL m_bMediaRemoved;
	BOOL m_bDeviceRemoved;
	BOOL m_bSpaceAvailable;
	BOOL m_bCanFinalizeMedia;
	BOOL m_bHasData;
	BOOL m_bErased;

	HWND m_hCancelPromptWnd;

	BOOL m_bErasing;
	UINT m_uiCountdown;

	CVsiPathNameEditWnd m_wndVolumeName;

public:

	enum { IDD = IDD_WRITE_CDDVD_VIEW };

	CVsiCdDvdWriterView();
	~CVsiCdDvdWriterView();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSICDDVDWRITERVIEW)

	BEGIN_COM_MAP(CVsiCdDvdWriterView)
		COM_INTERFACE_ENTRY(IVsiCdDvdWriterView)
		COM_INTERFACE_ENTRY(IVsiView)
		COM_INTERFACE_ENTRY(IVsiPreTranslateMsgFilter)
		COM_INTERFACE_ENTRY(IVsiEventSink)
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

	// IVsiCdDvdWriterView
	STDMETHOD(Initialize)(LPCOLESTR szDriveLetter, IUnknown *pUnkCdDvdWriter, const VARIANT* pvParam);

	// IVsiEventSink
	STDMETHOD(OnEvent)(DWORD dwEvent, const VARIANT *pv);

protected:
	BEGIN_MSG_MAP(CVsiCdDvdWriterView)
		CHAIN_MSG_MAP(CVsiShellHandlerImpl<CVsiCdDvdWriterView>)

		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_TIMER, OnTimer)

		MESSAGE_HANDLER(WM_VSI_PROGRESS_WND_CLOSED, OnProgressWindowClosed)
		MESSAGE_HANDLER(WM_VSI_UPDATE_UI, OnUpdateUI);

		MESSAGE_HANDLER(WM_VSI_WRITE_CANCELLED, OnCancelWrite);
		MESSAGE_HANDLER(WM_VSI_WRITE_ABORTED, OnAbortedWrite);
		MESSAGE_HANDLER(WM_VSI_START_WRITE, OnStartWrite);

		COMMAND_HANDLER(IDC_ERASE_DISC_BUTTON, BN_CLICKED, OnBnClickedErase)
		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)

//		COMMAND_HANDLER(IDC_FINALIZE_CHECK, BN_CLICKED, OnBnClickedFinalize)
		COMMAND_HANDLER(IDC_VERIFY_CHECK, BN_CLICKED, OnBnClickedVerify)

		CHAIN_MSG_MAP(CVsiLogoViewBase<CVsiCdDvdWriterView>)
	END_MSG_MAP()


	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnTimer(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnProgressWindowClosed(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnUpdateUI(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnCancelWrite(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnAbortedWrite(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnStartWrite(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnBnClickedErase(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);

//	LRESULT OnBnClickedFinalize(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedVerify(WORD, WORD, HWND, BOOL&);

	virtual BOOL ProcessNotification(WPARAM wParam, LPARAM lParam);


private:
	void UpdateUI();
	static DWORD CALLBACK WriteCallback(HWND hWnd, UINT uiState, LPVOID pData);
	static DWORD CALLBACK EraseCallback(HWND hWnd, UINT uiState, LPVOID pData);

	static DWORD CALLBACK CancelProgressCallback(int iID, LPVOID pData);

	BOOL IsDeviceIdle();
	void UpdateState();

};

OBJECT_ENTRY_AUTO(__uuidof(VsiCdDvdWriterView), CVsiCdDvdWriterView)
