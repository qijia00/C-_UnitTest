/*******************************************************************************
**
**  Copyright (c) 1999-2009 VisualSonics Inc. All Rights Reserved.
**
**  VsiLoginView.h
**
**	Description:
**		Declaration of the CVsirLogon
**
*******************************************************************************/

#pragma once

#include <VsiMathematics.h>
#include <VsiRes.h>
#include <VsiWtl.h>
#include <VsiMessageFilter.h>
#include <VsiOperatorModule.h>
#include <VsiAppViewModule.h>
#include "resource.h"

class ATL_NO_VTABLE CVsiLoginView :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiLoginView, &CLSID_VsiLoginView>,
	public CDialogImpl<CVsiLoginView>,
	public IVsiView,
	public IVsiPreTranslateMsgFilter
{
private:
	typedef enum
	{
		VSI_LOGIN_MESSAGE_NONE = 0,
		VSI_LOGIN_MESSAGE_CAPSLOCK_ON,
		VSI_LOGIN_MESSAGE_ACCOUNT_DISABLED,
		VSI_LOGIN_MESSAGE_PASSWORD_INCORRECT,
	} VSI_LOGIN_MESSAGE_TYPE;

	CComQIPtr<IVsiApp> m_pApp;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;
	CComPtr<IVsiCommandManager> m_pCmdMgr;

	std::unique_ptr<Gdiplus::Bitmap> m_pLogo;
	std::unique_ptr<Gdiplus::Bitmap> m_pEnter;
	std::unique_ptr<Gdiplus::Bitmap> m_pIcons;
	std::unique_ptr<Gdiplus::CachedBitmap> m_pLoginBkg;
	CSize m_sizeLoginBkg;
	CImageList m_imageListEnter;
	CImageList m_imageListShutdown;

	HBRUSH m_hBrushBkg;

	VSI_LOGIN_MESSAGE_TYPE m_dwMsgType;
	CString m_strMsg;
	int m_iLoginSessionCounter;
#ifdef _DEBUG
	CString m_strPassword;  
#endif  // _DEBUG
	CWindow m_wndName;

	class CVsiPasswordEdit :
		public CWindowImpl<CVsiPasswordEdit>
	{
	private:
		CVsiLoginView *m_pParent;

	public:
		CVsiPasswordEdit(CVsiLoginView *pParent) :
			m_pParent(pParent)
		{
		}
	protected:
		BEGIN_MSG_MAP(CVsiLoginView)
			MESSAGE_HANDLER(EM_SHOWBALLOONTIP, OnEmShowBalloonTip)
		END_MSG_MAP()

		LRESULT OnEmShowBalloonTip(UINT, WPARAM, LPARAM, BOOL&)
		{
			return FALSE;
		}
	} m_wndPassword;

	class CVsiBtnState
	{
	public:
		typedef enum
		{
			VSI_STATE_HIDDEN = 0,
			VSI_STATE_SHOWN,
			VSI_STATE_HIDING,
			VSI_STATE_SHOWING,
		} VSI_STATE;

		double m_dbAlpha;
		double m_dbPosition;

	private:
		VSI_STATE m_state;
		DWORD m_dwTick;

	public:
		CVsiBtnState() :
			m_dbAlpha(0.0),
			m_dbPosition(0.0),
			m_state(VSI_STATE_HIDDEN),
			m_dwTick(GetTickCount())
		{
		}

		VSI_STATE GetState() { return m_state; }

		void Show()
		{
			switch (m_state)
			{
			case VSI_STATE_HIDDEN:
				m_dwTick = GetTickCount();
				m_state = VSI_STATE_SHOWING;
				break;
			case VSI_STATE_HIDING:
				m_state = VSI_STATE_SHOWING;
				break;
			case VSI_STATE_SHOWN:
			case VSI_STATE_SHOWING:
				// Nothing
				break;
			}
		}

		void Hide()
		{
			switch (m_state)
			{
			case VSI_STATE_SHOWN:
				m_dwTick = GetTickCount();
				m_state = VSI_STATE_HIDING;
				break;
			case VSI_STATE_SHOWING:
				m_state = VSI_STATE_HIDING;
				break;
			case VSI_STATE_HIDDEN:
			case VSI_STATE_HIDING:
				// Nothing
				break;
			}
		}

		VSI_STATE Tick()
		{
			switch (m_state)
			{
			case VSI_STATE_HIDING:
				{
					const double dbAlphaTime(500.0);
					const double dbPositionTime(500.0);

					DWORD dwTick = GetTickCount();
					DWORD dwTimeDiff = dwTick - m_dwTick;

					m_dbAlpha -= 1.0 * dwTimeDiff / dbAlphaTime;
					if (m_dbAlpha < 0.0)
					{
						m_dbAlpha = 0.0;
					}

					m_dbPosition -= 1.0 * dwTimeDiff / dbPositionTime;
					if (m_dbPosition < 0.0)
					{
						m_dbPosition = 0.0;
					}

					m_dwTick = dwTick;

					if (VsiMaths::IsEqualUlps(m_dbAlpha, 0.0) &&
						VsiMaths::IsEqualUlps(m_dbPosition, 0.0))
					{
						m_state = VSI_STATE_HIDDEN;
					}
				}
				break;
			case VSI_STATE_SHOWING:
				{
					const double dbAlphaTime(500.0);
					const double dbPositionTime(500.0);

					DWORD dwTick = GetTickCount();
					DWORD dwTimeDiff = dwTick - m_dwTick;

					m_dbPosition += 1.0 * dwTimeDiff / dbPositionTime;
					if (m_dbPosition > 1.0)
					{
						m_dbPosition = 1.0;
					}

					m_dbAlpha += 1.0 * dwTimeDiff / dbAlphaTime;
					if (m_dbAlpha > 1.0)
					{
						m_dbAlpha = 1.0;
					}

					m_dwTick = dwTick;

					if (VsiMaths::IsEqualUlps(m_dbAlpha, 1.0) &&
						VsiMaths::IsEqualUlps(m_dbPosition, 1.0))
					{
						m_state = VSI_STATE_SHOWN;
					}
				}
				break;
			case VSI_STATE_HIDDEN:
			case VSI_STATE_SHOWN:
				// Nothing
				break;
			}

			return m_state;
		}
	} m_stateEnter;

public:
	enum { IDD = IDD_LOGIN_VIEW };

	CVsiLoginView();
	~CVsiLoginView();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSILOGINVIEW)

	BEGIN_COM_MAP(CVsiLoginView)
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

	// IVsiPreTranslateMsgFilter
	STDMETHOD(PreTranslateMessage)(MSG *pMsg, BOOL *pbHandled);

protected:
	BEGIN_MSG_MAP(CVsiLoginView)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)

		MESSAGE_HANDLER(WM_CTLCOLORDLG, OnCtlColorDlg)
		MESSAGE_HANDLER(WM_ERASEBKGND, OnEraseBkgnd)
		MESSAGE_HANDLER(WM_PRINTCLIENT, OnPrintClient)
		MESSAGE_HANDLER(WM_PAINT, OnPaint)
		MESSAGE_HANDLER(WM_TIMER, OnTimer)

		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDC_LOGIN_SHUTDOWN, BN_CLICKED, OnBnClickedShutdown)
		COMMAND_HANDLER(IDC_LOGIN_NAME_COMBO, CBN_CLOSEUP, OnCbSelChange)
		COMMAND_HANDLER(IDC_LOGIN_PASSWORD_EDIT, EN_UPDATE, OnEnChange);
	END_MSG_MAP()

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wp, LPARAM lp, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnCtlColorDlg(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnEraseBkgnd(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnPrintClient(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnPaint(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnTimer(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedShutdown(WORD, WORD, HWND, BOOL&);

	LRESULT OnCbSelChange(WORD, WORD, HWND, BOOL&);
	LRESULT OnEnChange(WORD, WORD, HWND, BOOL&);

private:
	void SetMessage(VSI_LOGIN_MESSAGE_TYPE type, LPCWSTR pszMsg);
	void UpdateUI();
	void FillNameCombo(bool bServiceMode = false);
	void Cleanup();
	void UpdateCapsLockState();
	void CheckAnimation();
};

OBJECT_ENTRY_AUTO(__uuidof(VsiLoginView), CVsiLoginView)
