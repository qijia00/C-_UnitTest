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

#define _USE_MATH_DEFINES
#include <math.h>
#include <VsiCommCtlLib.h>
#include <VsiRes.h>


class CVsiSearchWnd :
	public CWindowImpl<CVsiSearchWnd>
{
private:
	CWindow m_wndEdit;

	std::unique_ptr<Gdiplus::Bitmap> m_pbmpControl;

	typedef enum
	{
		VSI_IMAGE_SEARCH = 0,
		VSI_IMAGE_SEARCH_HOT,
		VSI_IMAGE_CANCEL,
		VSI_IMAGE_CANCEL_HOT,
		VSI_IMAGE_MAX,
		VSI_IMAGE_NONE,
	} VSI_IMAGE;

	VSI_IMAGE m_stateSearch;
	VSI_IMAGE m_stateCancel;

	bool m_bTrackLeave;
	bool m_bDown;

	UINT_PTR m_iTimer;
	float m_pfAlpha[25];
	int m_iPaintCycle;

public:
	CVsiSearchWnd() :
		m_stateSearch(VSI_IMAGE_SEARCH),
		m_stateCancel(VSI_IMAGE_NONE),
		m_bTrackLeave(false),
		m_bDown(false),
		m_iTimer(0),
		m_iPaintCycle(0)
	{
		float fStep = (float)(M_PI) / _countof(m_pfAlpha);
		float fPos((float)(M_PI) / 2.0f);
		for (int i = 0; i < _countof(m_pfAlpha); ++i)
		{
			m_pfAlpha[i] = abs(sin(fPos)) * 0.5f + 0.5f;
			fPos += fStep;
		}
	}
	~CVsiSearchWnd()
	{
	}

	void GetWindowTextW(CString& strText)
	{
		m_wndEdit.GetWindowText(strText);
	}
	void SetWindowText(LPCWSTR pszText)
	{
		m_wndEdit.SetWindowTextW(pszText);
	}

	void SetSearchAnimation(bool bAnim)
	{
		if (bAnim)
		{
			if (0 == m_iTimer)
			{
				m_iPaintCycle = 0;
				m_iTimer = SetTimer(1001, 50, NULL);
			}
		}
		else
		{
			if (0 != m_iTimer)
			{
				KillTimer(m_iTimer);
				m_iTimer = 0;
			}

			m_iPaintCycle = 0;
			Invalidate();
		}
	}

protected:
	BEGIN_MSG_MAP(CVsiSearchWnd)
		MESSAGE_HANDLER(WM_CREATE, OnCreate)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_GETTEXT, OnGetText)
		MESSAGE_HANDLER(WM_SETTEXT, OnSetText)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
		MESSAGE_HANDLER(WM_SIZE, OnSize)
		MESSAGE_HANDLER(WM_TIMER, OnTimer)
		MESSAGE_HANDLER(WM_ERASEBKGND, OnEraseBkgnd)
		MESSAGE_HANDLER(WM_NCPAINT, OnNcPaint)
		MESSAGE_HANDLER(WM_PRINTCLIENT, OnPrintClient)
		MESSAGE_HANDLER(WM_PAINT, OnPaint)
		MESSAGE_HANDLER(WM_CTLCOLOREDIT, OnCtlColorEdit)
		MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
		MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
		MESSAGE_HANDLER(WM_LBUTTONDOWN, OnMouseDown)
		MESSAGE_HANDLER(WM_LBUTTONUP, OnMouseUp)
		MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
		MESSAGE_HANDLER(WM_PARENTNOTIFY, OnParentNotify)
		MESSAGE_HANDLER(WM_CAPTURECHANGED, OnCaptureChanged)

		COMMAND_CODE_HANDLER(EN_CHANGE, OnEditChanged)
		COMMAND_CODE_HANDLER(EN_SETFOCUS, OnEditFocus)
		COMMAND_CODE_HANDLER(EN_KILLFOCUS, OnEditFocus)
	END_MSG_MAP()

	// Handler prototypes:
	//  LRESULT MessageHandler(UINT uMsg, WPARAM wp, LPARAM lp, BOOL& bHandled);
	//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnCreate(UINT, WPARAM, LPARAM, BOOL&)
	{
		HRESULT hr = S_OK;

		try
		{
			CRect rc;
			GetClientRect(&rc);

			m_pbmpControl.reset(
				VsiLoadImageFromResource(
					MAKEINTRESOURCE(IDR_SEARCH), 
					L"PNG", 
					_AtlBaseModule.GetResourceInstance()));
			VSI_CHECK_MEM(m_pbmpControl, VSI_LOG_ERROR, NULL);

			rc.left = 18;
			rc.top = 2;

			m_wndEdit.Create(WC_EDIT, *this, rc, L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL);
		}
		VSI_CATCH_(err)
		{
			hr = err;
			if (FAILED(hr))
				VSI_ERROR_LOG(err);

			DestroyWindow();
		}

		return 0;
	}
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
	{
		if (0 != m_iTimer)
		{
			KillTimer(m_iTimer);
			m_iTimer = 0;
		}

		m_wndEdit.DestroyWindow();

		m_pbmpControl.reset();

		return 0;
	}
	LRESULT OnGetText(UINT uiMsg, WPARAM wp, LPARAM lp, BOOL&)
	{
		return m_wndEdit.SendMessage(uiMsg, wp, lp);
	}
	LRESULT OnSetText(UINT uiMsg, WPARAM wp, LPARAM lp, BOOL&)
	{
		return m_wndEdit.SendMessage(uiMsg, wp, lp);
	}
	LRESULT OnSetFocus(UINT, WPARAM, LPARAM, BOOL&)
	{
		m_wndEdit.SetFocus();

		return 0;
	}
	LRESULT OnSize(UINT, WPARAM, LPARAM lp, BOOL&)
	{
		CRect rc(18, 2, LOWORD(lp) - 18, HIWORD(lp));

		m_wndEdit.SetWindowPos(NULL, &rc, SWP_NOACTIVATE | SWP_NOZORDER);

		return 0;
	}
	LRESULT OnTimer(UINT, WPARAM wp, LPARAM, BOOL &bHandled)
	{
		bHandled = FALSE;

		if (m_iTimer == wp)
		{
			m_iPaintCycle = (m_iPaintCycle + 1) % _countof(m_pfAlpha);

			Invalidate();
		}
	
		return 0;
	}
	LRESULT OnEraseBkgnd(UINT, WPARAM, LPARAM, BOOL&)
	{
		return 1;  // don't erase
	}
	LRESULT OnNcPaint(UINT, WPARAM wp, LPARAM, BOOL&)
	{
		CWindow wndFocus(GetFocus());

		BOOL bActive = IsWindowEnabled() && (wndFocus == *this || wndFocus == m_wndEdit);

		VsiNcPaintHandlerActive(*this, wp, bActive);

		return 0;
	}
	LRESULT OnPrintClient(UINT, WPARAM wp, LPARAM, BOOL&)
	{
		// Draw
		HDC hdcTarget = (HDC)wp;

		CRect rc;
		GetClientRect(&rc);

		HDC hdc(NULL);
		HPAINTBUFFER hpb = VsiBeginBufferedPaint(hdcTarget, &rc, BPBF_COMPATIBLEBITMAP, NULL, &hdc); 
		{
			Gdiplus::Graphics graphics(hdc);

			// Fill background
			Gdiplus::Color colorBkg;
			colorBkg.SetFromCOLORREF(VSI_COLOR_EDIT_BKG);
			Gdiplus::SolidBrush brushBkg(colorBkg);

			graphics.FillRectangle(&brushBkg, rc.left, rc.top, rc.right, rc.bottom);

			int iImageCx = m_pbmpControl->GetWidth() / VSI_IMAGE_MAX;
			int iY = (rc.Height() - m_pbmpControl->GetHeight() + 1) / 2;

			Gdiplus::Rect rcIcon(2, iY, iImageCx, m_pbmpControl->GetHeight());

			// In Vista, the bitmap doesn't appear when using ColorMatrix (is fine in Windows 7)
			// Work fine when making a copy first
			Gdiplus::Bitmap bmp(rcIcon.Width, rcIcon.Height);
			{
				Gdiplus::Graphics graphics(&bmp);
				graphics.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);
				graphics.ScaleTransform(1.0f, 1.0f);
				graphics.DrawImage(
					m_pbmpControl.get(),
					Gdiplus::Rect(0, 0, rcIcon.Width, rcIcon.Height),
					m_stateSearch * iImageCx, 0, rcIcon.Width, rcIcon.Height,
					Gdiplus::UnitPixel);
			}

			// Set the opacity
			Gdiplus::ColorMatrix cm = {
				1.0f, 0.0f, 0.0f, 0.0f, 0.0f,
				0.0f, 1.0f, 0.0f, 0.0f, 0.0f,
				0.0f, 0.0f, 1.0f, 0.0f, 0.0f,
				0.0f, 0.0f, 0.0f, m_pfAlpha[m_iPaintCycle], 0.0f,
				0.0f, 0.0f, 0.0f, 0.0f, 1.0f,
			};

			Gdiplus::ImageAttributes ias;
			ias.SetColorMatrix(
				&cm,
				Gdiplus::ColorMatrixFlagsDefault,
				Gdiplus::ColorAdjustTypeBitmap);

			graphics.DrawImage(
				&bmp, rcIcon,
				0, 0, rcIcon.Width, rcIcon.Height,
				Gdiplus::UnitPixel,
				&ias);

			if (VSI_IMAGE_NONE != m_stateCancel)
			{
				rcIcon.X = rc.right - iImageCx - 2;
				graphics.DrawImage(
					m_pbmpControl.get(), rcIcon,
					m_stateCancel * iImageCx, 0, rcIcon.Width, rcIcon.Height,
					Gdiplus::UnitPixel);
			}
		}
		VsiEndBufferedPaint(hpb, TRUE);

		return 0;
	}
	LRESULT OnPaint(UINT, WPARAM, LPARAM, BOOL& bHandled)
	{
		PAINTSTRUCT ps;

		BeginPaint(&ps);

		OnPrintClient(WM_PRINTCLIENT, (WPARAM)ps.hdc, PRF_CLIENT, bHandled);

		EndPaint(&ps);

		return 0;
	}
	LRESULT OnCtlColorEdit(UINT uiMsg, WPARAM wp, LPARAM lp, BOOL&)
	{
		return GetParent().SendMessage(uiMsg, wp, lp);
	}
	LRESULT OnSetCursor(UINT, WPARAM, LPARAM, BOOL& bHandled)
	{
		UpdateState();

		bHandled = FALSE;

		return 0;
	}
	LRESULT OnMouseMove(UINT, WPARAM wp, LPARAM, BOOL&)
	{
		if (!m_bTrackLeave)
		{
			m_bTrackLeave = true;
	
			TRACKMOUSEEVENT tme = { sizeof(TRACKMOUSEEVENT), TME_LEAVE, m_hWnd, 0 };
			TrackMouseEvent(&tme);
		}

		UpdateState(0 < (MK_LBUTTON & wp));

		return 0;
	}
	LRESULT OnMouseDown(UINT, WPARAM, LPARAM lp, BOOL&)
	{
		int iLen = m_wndEdit.GetWindowTextLength();
		if (0 < iLen)
		{
			POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };

			CRect rc;
			GetButtonRect(&rc);

			if (rc.PtInRect(pt))
			{
				m_bDown = true;
				SetCapture();
			}

			UpdateState(true);
		}

		return 0;
	}
	LRESULT OnMouseUp(UINT, WPARAM, LPARAM lp, BOOL&)
	{
		if (m_bDown)
		{
			POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };

			CRect rc;
			GetButtonRect(&rc);

			if (rc.PtInRect(pt))
			{
				m_wndEdit.SetWindowText(L"");

				GetParent().SendMessage(
					WM_COMMAND,
					MAKEWPARAM(GetDlgCtrlID(), EN_CHANGE),
					reinterpret_cast<LPARAM>(m_hWnd));
			}
		}

		if (*this == GetCapture())
		{
			ReleaseCapture();
		}

		return 0;
	}
	LRESULT OnMouseLeave(UINT, WPARAM, LPARAM, BOOL&)
	{
		m_bTrackLeave = false;

		UpdateState();

		return 0;
	}
	LRESULT OnParentNotify(UINT, WPARAM wp, LPARAM, BOOL &bHandled)
	{
		if (WM_MOUSELEAVE == LOWORD(wp))
		{
			UpdateState();
		}

		bHandled = FALSE;

		return 0;
	}
	LRESULT OnCaptureChanged(UINT, WPARAM, LPARAM, BOOL&)
	{
		m_bDown = false;

		UpdateState();

		return 0;
	}
	LRESULT OnEditChanged(WORD wNotifyCode, WORD, HWND, BOOL&)
	{
		UpdateState();

		return GetParent().SendMessage(
			WM_COMMAND,
			MAKEWPARAM(GetDlgCtrlID(), wNotifyCode),
			reinterpret_cast<LPARAM>(m_hWnd));
	}
	LRESULT OnEditFocus(WORD, WORD, HWND, BOOL &bHandled)
	{
		bHandled = FALSE;

		UpdateState();

		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE);

		return 0;
	}

private:
	void UpdateState(bool bDown = false)
	{
		int iLen = m_wndEdit.GetWindowTextLength();

		VSI_IMAGE stateSearch(VSI_IMAGE_SEARCH);
		VSI_IMAGE stateCancel(VSI_IMAGE_NONE);
		
		DWORD dwPos = GetMessagePos();
		POINT pt = { GET_X_LPARAM(dwPos), GET_Y_LPARAM(dwPos) };
		ScreenToClient(&pt);

		if (0 < iLen)
		{
			stateSearch = VSI_IMAGE_SEARCH_HOT;

			CRect rc;
			GetButtonRect(&rc);

			if (rc.PtInRect(pt))
			{
				stateCancel = bDown ? VSI_IMAGE_CANCEL : VSI_IMAGE_CANCEL_HOT;
			}
			else
			{
				stateCancel = m_bDown ? VSI_IMAGE_CANCEL_HOT : VSI_IMAGE_CANCEL;
			}
		}
		else
		{
			stateCancel = VSI_IMAGE_NONE;

			CRect rc;
			GetClientRect(&rc);

			if (rc.PtInRect(pt) || (GetFocus() == m_wndEdit))
			{
				stateSearch = VSI_IMAGE_SEARCH_HOT;
			}
		}

		if (stateSearch != m_stateSearch || stateCancel != m_stateCancel)
		{
			m_stateSearch = stateSearch;
			m_stateCancel = stateCancel;

			RedrawWindow();
		}
	}

	void GetButtonRect(LPRECT prc)
	{
		GetClientRect(prc);
		prc->left = prc->right - m_pbmpControl->GetWidth() / VSI_IMAGE_MAX - 2;
	}
};


