/*******************************************************************************
**
**  Copyright (c) 1999-2009 VisualSonics Inc. All Rights Reserved.
**
**  VsiLoginView.cpp
**
**	Description:
**		Implementation of CVsiLoginView
**
*******************************************************************************/

#include "stdafx.h"
#ifdef _DEBUG
#include <VsiCommUtlLib.h>   // VsiCreateSha1
#endif  // _DEBUG
#include <VsiThemeColor.h>
#include <VsiServiceProvider.h>
#include <VsiCommCtlLib.h>
#include <VsiCommCtl.h>
#include <VsiBuildNumber.h>
#include <VsiServiceKey.h>
#include <VsiAppControllerModule.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiPdmModule.h>
#include "VsiLoginView.h"

#define VSI_LOGIN_ATTEMPTS_MAX 10
#define VSI_LOGIN_CLOCK_TIMER_ID 1

#define VSI_LOGIN_FADE_CX 16
#define VSI_EFFECTS_SOFT

CVsiLoginView::CVsiLoginView() :
	m_hBrushBkg(NULL),
	m_dwMsgType(VSI_LOGIN_MESSAGE_NONE),
	m_iLoginSessionCounter(VSI_LOGIN_ATTEMPTS_MAX),
	m_wndPassword(this)
{
}

CVsiLoginView::~CVsiLoginView()
{
	_ASSERT(NULL == m_hBrushBkg);
	_ASSERT(NULL == m_pLogo);
	_ASSERT(NULL == m_pLoginBkg);
	_ASSERT(NULL == m_pEnter);
}

STDMETHODIMP CVsiLoginView::Activate(
	IVsiApp *pApp,
	HWND hwndParent)
{
	HRESULT hr = S_OK;

	try
	{
		m_pApp = pApp;

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Get PDM
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Operator Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_OPERATOR_MGR,
			__uuidof(IVsiOperatorManager),
			(IUnknown**)&m_pOperatorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Command Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CMD_MGR,
			__uuidof(IVsiCommandManager),
			(IUnknown**)&m_pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HWND hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
			VSI_ERROR_LOG(err);

		Deactivate();
	}

	return hr;
}

STDMETHODIMP CVsiLoginView::Deactivate()
{
	HRESULT hr = S_FALSE;

	if (IsWindow())
	{
		BOOL bRet = DestroyWindow();
		_ASSERT(bRet);

		hr = S_OK;
	}

	m_pCmdMgr.Release();
	m_pOperatorMgr.Release();
	m_pPdm.Release();
	m_pApp.Release();

	return hr;
}

STDMETHODIMP CVsiLoginView::GetWindow(HWND *phWnd)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(phWnd, VSI_LOG_ERROR, L"phWnd");

		*phWnd = m_hWnd;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiLoginView::GetAccelerator(HACCEL *phAccel)
{
	UNREFERENCED_PARAMETER(phAccel);
	return E_NOTIMPL;
}

STDMETHODIMP CVsiLoginView::GetIsBusy(
	DWORD dwStateCurrent,
	DWORD dwState,
	BOOL bTryRelax,
	BOOL *pbBusy)
{
	UNREFERENCED_PARAMETER(dwStateCurrent);
	UNREFERENCED_PARAMETER(dwState);
	UNREFERENCED_PARAMETER(bTryRelax);

	*pbBusy = !GetParent().IsWindowEnabled();

	return S_OK;
}

STDMETHODIMP CVsiLoginView::PreTranslateMessage(MSG *pMsg, BOOL *pbHandled)
{
	UpdateCapsLockState();

	// Ctrl-Alt-S
	if (GetKeyState(VK_CONTROL) < 0  && GetKeyState(VK_MENU) < 0 && GetKeyState(L'S') < 0)
	{
		FillNameCombo(true);
		UpdateUI();
	}

	*pbHandled = IsDialogMessage(pMsg);

	return S_OK;
}

LRESULT CVsiLoginView::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	LRESULT lRet = 1;
	HBITMAP hBmp(NULL);

	try
	{
		_ASSERT(NULL == m_hBrushBkg);
		m_hBrushBkg = CreateSolidBrush(VSI_COLOR_LOGIN_BKG);

		// Set font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_L, &hFont);
		VsiThemeRecurSetFont(m_hWnd, hFont);

		// Images
		m_pLogo.reset(
			VsiLoadImageFromResource(
				MAKEINTRESOURCE(IDR_LOGIN_LOGO),
				L"PNG",
				_AtlBaseModule.GetResourceInstance()));
		VSI_VERIFY(NULL != m_pLogo, VSI_LOG_ERROR, NULL);

		m_pEnter.reset(
			VsiLoadImageFromResource(
				MAKEINTRESOURCE(IDR_LOGIN_ENTER),
				L"PNG",
				_AtlBaseModule.GetResourceInstance()));
		VSI_VERIFY(NULL != m_pEnter, VSI_LOG_ERROR, NULL);

		m_pIcons.reset(
			VsiLoadImageFromResource(
				MAKEINTRESOURCE(IDR_LOGIN_ICONS),
				L"PNG",
				_AtlBaseModule.GetResourceInstance()));
		VSI_VERIFY(NULL != m_pIcons, VSI_LOG_ERROR, NULL);

		// Buttons
		{
			CRect rcEnter;
			CWindow wndBtnOk(GetDlgItem(IDOK));
			wndBtnOk.GetWindowRect(&rcEnter);

			std::unique_ptr<Gdiplus::Bitmap> pBmpEnter(
				VsiLoadImageFromResource(
					MAKEINTRESOURCE(IDR_LOGIN_ENTER),
					L"PNG",
					_AtlBaseModule.GetResourceInstance()));
			VSI_VERIFY(NULL != pBmpEnter, VSI_LOG_ERROR, NULL);

			rcEnter.right = rcEnter.left + rcEnter.Height() * (pBmpEnter->GetWidth() / 4) / pBmpEnter->GetHeight();
			wndBtnOk.SetWindowPos(NULL, &rcEnter, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOZORDER);

			m_imageListEnter.Create(rcEnter.Width(), rcEnter.Height(), ILC_COLOR32, 4, 1);

			Gdiplus::Bitmap bmp(rcEnter.Width() * 4, rcEnter.Height());
			Gdiplus::Graphics graphic(&bmp);
			graphic.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);

			graphic.DrawImage(
				pBmpEnter.get(),
				Gdiplus::RectF(0.0f, 0.0f, (Gdiplus::REAL)bmp.GetWidth(), (Gdiplus::REAL)bmp.GetHeight()),
				0.0f,
				0.0f,
				(Gdiplus::REAL)pBmpEnter->GetWidth(),
				(Gdiplus::REAL)pBmpEnter->GetHeight(),
				Gdiplus::UnitPixel);

			bmp.GetHBITMAP(Gdiplus::Color(), &hBmp);

			m_imageListEnter.Add(hBmp, RGB(0, 0, 0));

			DeleteObject(hBmp);
			hBmp = NULL;

			VSI_SETIMAGE si = { VSI_BTN_FLAG_SINGLE_IMAGELIST };
			si.hImageLists[0] = m_imageListEnter;

			UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);
			wndBtnOk.SendMessage(uiThemeCmd, VSI_BT_CMD_SETIMAGE, (LPARAM)&si);
			wndBtnOk.SendMessage(uiThemeCmd, VSI_BT_CMD_SETTIP, (LPARAM)L"Enter");
		}

		// Shut down button - system only
		VSI_APP_MODE_FLAGS dwAppMode(VSI_APP_MODE_NONE);
		m_pApp->GetAppMode(&dwAppMode);

		if (0 != (VSI_APP_MODE_SYSTEM & dwAppMode))
		{
			CWindow wndBtnShutdown(GetDlgItem(IDC_LOGIN_SHUTDOWN));

			CRect rcEnter;
			wndBtnShutdown.GetWindowRect(&rcEnter);

			std::unique_ptr<Gdiplus::Bitmap> pBmpShutdown(
				VsiLoadImageFromResource(
					MAKEINTRESOURCE(IDR_LOGIN_SHUTDOWN),
					L"PNG",
					_AtlBaseModule.GetResourceInstance()));
			VSI_VERIFY(NULL != pBmpShutdown, VSI_LOG_ERROR, NULL);

			rcEnter.right = rcEnter.left + rcEnter.Height() * (pBmpShutdown->GetWidth() / 4) / pBmpShutdown->GetHeight();
			wndBtnShutdown.SetWindowPos(NULL, &rcEnter, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOZORDER);

			m_imageListShutdown.Create(rcEnter.Width(), rcEnter.Height(), ILC_COLOR32, 4, 1);

			Gdiplus::Bitmap bmp(rcEnter.Width() * 4, rcEnter.Height());
			Gdiplus::Graphics graphic(&bmp);
			graphic.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);

			graphic.DrawImage(
				pBmpShutdown.get(),
				Gdiplus::RectF(0.0f, 0.0f, (Gdiplus::REAL)bmp.GetWidth(), (Gdiplus::REAL)bmp.GetHeight()),
				0.0f,
				0.0f,
				(Gdiplus::REAL)pBmpShutdown->GetWidth(),
				(Gdiplus::REAL)pBmpShutdown->GetHeight(),
				Gdiplus::UnitPixel);

			bmp.GetHBITMAP(Gdiplus::Color(), &hBmp);

			m_imageListShutdown.Add(hBmp, RGB(0, 0, 0));

			DeleteObject(hBmp);
			hBmp = NULL;

			VSI_SETIMAGE si = { VSI_BTN_FLAG_SINGLE_IMAGELIST };
			si.hImageLists[0] = m_imageListShutdown;

			UINT uiThemeCmd = RegisterWindowMessage(VSI_THEME_COMMAND);
			wndBtnShutdown.SendMessage(uiThemeCmd, VSI_BT_CMD_SETIMAGE, (LPARAM)&si);
			wndBtnShutdown.SendMessage(uiThemeCmd, VSI_BT_CMD_SETTIP, (LPARAM)L"Shut Down");

			wndBtnShutdown.EnableWindow();
			wndBtnShutdown.ShowWindow(SW_SHOWNA);
		}

		// Layout
		CComPtr<IVsiWindowLayout> pLayout;
		HRESULT hr = pLayout.CoCreateInstance(__uuidof(VsiWindowLayout));
		if (SUCCEEDED(hr))
		{
			hr = pLayout->Initialize(m_hWnd, VSI_WL_FLAG_NONE);
			if (SUCCEEDED(hr))
			{
				UINT nIdGroup;
				pLayout->AddControl(0, IDC_LOGIN_SHUTDOWN, VSI_WL_MOVE_XY);
				pLayout->AddGroup(0, VSI_WL_CENTER_XY, &nIdGroup);
				pLayout->AddControl(nIdGroup, IDC_LOGIN_TOP_DUMMY, VSI_WL_NONE);
				pLayout->AddControl(nIdGroup, IDC_LOGIN_NAME_COMBO, VSI_WL_NONE);
				pLayout->AddControl(nIdGroup, IDC_LOGIN_PASSWORD_EDIT, VSI_WL_NONE);
				pLayout->AddControl(nIdGroup, IDC_LOGIN_MSG, VSI_WL_NONE);
				pLayout->AddControl(nIdGroup, IDOK, VSI_WL_NONE);
				pLayout->AddControl(0, IDC_LOGIN_VER, VSI_WL_MOVE_Y);
				pLayout->AddControl(0, IDC_LOGIN_TIME, VSI_WL_MOVE_Y);
				pLayout->AddControl(0, IDC_LOGIN_DATE, VSI_WL_MOVE_Y);

				pLayout.Detach();
			}
		}

		m_dwMsgType = VSI_LOGIN_MESSAGE_NONE;
		m_strMsg = L"";
		m_iLoginSessionCounter = VSI_LOGIN_ATTEMPTS_MAX;

		// Populate list
		m_wndName = GetDlgItem(IDC_LOGIN_NAME_COMBO);
		m_wndPassword.SubclassWindow(GetDlgItem(IDC_LOGIN_PASSWORD_EDIT));

		m_wndPassword.SendMessage(EM_SETCUEBANNER, TRUE, (LPARAM)L"Password");

		// Sets limits
		m_wndPassword.SendMessage(EM_SETLIMITTEXT, VSI_OPERATOR_PASSWORD_MAX);

		FillNameCombo();

		RedrawWindow();

		UpdateUI();

		if (m_wndPassword.IsWindowEnabled())
		{
			PostMessage(WM_NEXTDLGCTL, (WPARAM)m_wndPassword.operator HWND(), TRUE);
			lRet = 0;
		}

		SetTimer(VSI_LOGIN_CLOCK_TIMER_ID, 250, NULL);

// Shows service mode password in debug build
#ifdef _DEBUG
		SYSTEMTIME stLocal;
		GetLocalTime(&stLocal);

		CString strDate;
		strDate.Format(L"%4d%2d%2d", stLocal.wYear, stLocal.wMonth, stLocal.wDay);

		CString strHashed;
		hr = VsiCreateSha1(
			(const BYTE*)strDate.GetString(),
			strDate.GetLength() * sizeof(WCHAR),
			strHashed.GetBufferSetLength(41),
			41);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		strHashed.ReleaseBuffer();

		m_strPassword = strHashed.Left(10);
#endif  // _DEBUG
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
			VSI_ERROR_LOG(err);

		if (NULL != hBmp)
		{
			DeleteObject(hBmp);
		}

		DestroyWindow();
	}

	return lRet;
}

LRESULT CVsiLoginView::OnDestroy(UINT, WPARAM, LPARAM, BOOL& bHandled)
{
	KillTimer(VSI_LOGIN_CLOCK_TIMER_ID);

	m_wndPassword.UnsubclassWindow();

	Cleanup();

	m_pLoginBkg.reset();
	m_pIcons.reset();
	m_pLogo.reset();
	m_pEnter.reset();

	if (NULL != m_hBrushBkg)
	{
		DeleteObject(m_hBrushBkg);
		m_hBrushBkg = NULL;
	}

	if (!m_imageListEnter.IsNull())
	{
		m_imageListEnter.Destroy();
	}

	if (!m_imageListShutdown.IsNull())
	{
		m_imageListShutdown.Destroy();
	}

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiLoginView::OnCtlColorDlg(UINT, WPARAM, LPARAM, BOOL&)
{
	return reinterpret_cast<LRESULT>(m_hBrushBkg);
}

LRESULT CVsiLoginView::OnEraseBkgnd(UINT, WPARAM, LPARAM, BOOL &)
{
	return 1;  // don't erase
}

LRESULT CVsiLoginView::OnPrintClient(UINT, WPARAM wp, LPARAM, BOOL &)
{
	HRESULT hr = S_OK;
	HPAINTBUFFER hpb(NULL);

	try
	{
		HDC hdcTarget = (HDC)wp;

		CRect rcClip;
		GetClipBox(hdcTarget, &rcClip);

		HDC hdc(NULL);
		hpb = VsiBeginBufferedPaint(hdcTarget, &rcClip, BPBF_COMPATIBLEBITMAP, NULL, &hdc);
		VSI_VERIFY(NULL != hpb, VSI_LOG_ERROR, NULL);

		int iLogoW = m_pLogo->GetWidth();
		int iLogoH = m_pLogo->GetHeight();

		CRect rcClient;
		GetClientRect(&rcClient);

		int iW = rcClient.Width();
		int iH = rcClient.Height();
		const int iFooterHeight = 150;

		Gdiplus::Graphics graphics(hdc);
		graphics.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);

		// Background
		if (NULL == m_pLoginBkg)
		{
			std::unique_ptr<Gdiplus::Bitmap> pBmp(
				VsiLoadImageFromResource(
					MAKEINTRESOURCE(IDR_LOGIN_BKG),
					L"PNG",
					_AtlBaseModule.GetResourceInstance()));

			if (NULL != pBmp)
			{
				m_sizeLoginBkg.cx = pBmp->GetWidth();
				m_sizeLoginBkg.cy = pBmp->GetHeight();

				m_pLoginBkg.reset(new Gdiplus::CachedBitmap(pBmp.get(), &graphics));
				VSI_VERIFY(NULL != m_pLoginBkg, VSI_LOG_ERROR, NULL);
			}
		}

		if (NULL != m_pLoginBkg)
		{
			int iBgW = 0;
			int iByY = iH - iFooterHeight - m_sizeLoginBkg.cy;

			do
			{
				if (Gdiplus::Ok != graphics.DrawCachedBitmap(m_pLoginBkg.get(), iBgW, iByY))
				{
					// User changes resolutions or color depths, need to recreate the cached bitmap
					std::unique_ptr<Gdiplus::Bitmap> pBmp(
						VsiLoadImageFromResource(
							MAKEINTRESOURCE(IDR_LOGIN_BKG),
							L"PNG",
							_AtlBaseModule.GetResourceInstance()));

					if (NULL != pBmp)
					{
						m_sizeLoginBkg.cx = pBmp->GetWidth();
						m_sizeLoginBkg.cy = pBmp->GetHeight();

						m_pLoginBkg.reset(new Gdiplus::CachedBitmap(pBmp.get(), &graphics));
						VSI_VERIFY(NULL != m_pLoginBkg, VSI_LOG_ERROR, NULL);
					}
				}
				else
				{
					iBgW += m_sizeLoginBkg.cx;
				}
			}
			while (iBgW < iW);

			// Header
			if (iByY > 0)
			{
				Gdiplus::Rect rc(0, 0, iW, iByY);

				Gdiplus::Color colorBkg;
				colorBkg.SetFromCOLORREF(VSI_COLOR_LOGIN_BKG_TOP);
				Gdiplus::SolidBrush brush(colorBkg);
				graphics.FillRectangle(&brush, rc);
			}

			// Footer
			{
				Gdiplus::Rect rc(0, iH - iFooterHeight, iW, iFooterHeight);

				Gdiplus::Color colorBkg;
				colorBkg.SetFromCOLORREF(VSI_COLOR_LOGIN_BKG);
				Gdiplus::SolidBrush brush(colorBkg);
				graphics.FillRectangle(&brush, rc);
			}
		}

		graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);

		// Logo
		{
			CRect rcName;
			GetDlgItem(IDC_LOGIN_NAME_COMBO).GetWindowRect(&rcName);
			ScreenToClient(&rcName);

			int iImageW = min(iLogoW, MulDiv(iW, 8, 10));
			int iImageH = min(iLogoH, MulDiv(max(0, rcName.top), 8, 10));

			double dbWdivH1 = (double)iLogoW / (double)iLogoH;
			double dbWdivH2 = (double)iImageW / (double)iImageH;

			if (dbWdivH1 > dbWdivH2)
			{
				iImageH =  (int)(iImageW / dbWdivH1);
			}
			else
			{
				iImageW = (int)(dbWdivH1 * iImageH);
			}

			graphics.DrawImage(m_pLogo.get(), (rcClient.Width() - iImageW) / 2, (rcName.top - iImageH) / 2, iImageW, iImageH);
		}

		// Button
		switch (m_stateEnter.GetState())
		{
		case CVsiBtnState::VSI_STATE_HIDING:
		case CVsiBtnState::VSI_STATE_SHOWING:
			{
				CRect rcEnter;
				CWindow wndBtnOk(GetDlgItem(IDOK));
				wndBtnOk.GetWindowRect(&rcEnter);
				ScreenToClient(&rcEnter);

				Gdiplus::Rect rcDraw(rcEnter.left, rcEnter.top, rcEnter.Width(), rcEnter.Height());

				int iCxEnter = m_pEnter->GetWidth() / 4;
				int iCyEnter = m_pEnter->GetHeight();

				int iOffset = (int)(rcEnter.Width() * (1.0 - m_stateEnter.m_dbPosition) * (1.0 - m_stateEnter.m_dbPosition));

				#ifdef VSI_EFFECTS_SOFT
					rcDraw.X -= VSI_LOGIN_FADE_CX;
					rcDraw.Width += VSI_LOGIN_FADE_CX;
					Gdiplus::Bitmap bmp(rcEnter.Width() + VSI_LOGIN_FADE_CX, rcEnter.Height());
					{
						Gdiplus::Graphics graphicBmp(&bmp);
						graphicBmp.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
						graphicBmp.ScaleTransform(1.0f, 1.0f);

						// Set the opacity
						Gdiplus::ColorMatrix cm = {
							1.0f, 0.0f, 0.0f, 0.0f, 0.0f,
							0.0f, 1.0f, 0.0f, 0.0f, 0.0f,
							0.0f, 0.0f, 1.0f, 0.0f, 0.0f,
							0.0f, 0.0f, 0.0f, (float)(m_stateEnter.m_dbAlpha), 0.0f,
							0.0f, 0.0f, 0.0f, 0.0f, 1.0f,
						};

						Gdiplus::ImageAttributes ias;
						ias.SetColorMatrix(
							&cm,
							Gdiplus::ColorMatrixFlagsDefault,
							Gdiplus::ColorAdjustTypeBitmap);

						graphicBmp.DrawImage(
							m_pEnter.get(),
							Gdiplus::Rect(VSI_LOGIN_FADE_CX - iOffset, 0, rcEnter.Width(), rcEnter.Height()),
							0, 0, iCxEnter, iCyEnter,
							Gdiplus::UnitPixel,
							&ias);

						Gdiplus::Rect rcFade(0, 0, VSI_LOGIN_FADE_CX, rcEnter.Height());

						BYTE pAlpha[VSI_LOGIN_FADE_CX];
						for (int x = 0; x < VSI_LOGIN_FADE_CX; ++x)
						{
							double dbCurve = ((double)x / VSI_LOGIN_FADE_CX);
							dbCurve *= dbCurve;

							pAlpha[x] = (BYTE)(255.0 * dbCurve);
						}

						Gdiplus::BitmapData lockedBitmapData = { 0 };
						bmp.LockBits(
							&rcFade,
							Gdiplus::ImageLockModeRead,
							PixelFormat32bppARGB,
							&lockedBitmapData);
						if (NULL ==lockedBitmapData.Scan0)
						{
							VSI_VERIFY(lockedBitmapData.Scan0 != NULL, VSI_LOG_ERROR, NULL);
						}

						BYTE *pbtData = (BYTE *)lockedBitmapData.Scan0;
						if (lockedBitmapData.Stride < 0)
						{
							pbtData = pbtData + ((int)lockedBitmapData.Height - 1) * lockedBitmapData.Stride;
						}

						for (int y = 0; y < rcFade.Height; ++y)
						{
							BYTE *pbtWrite = pbtData + 3;

							for (int x = 0; x < rcFade.Width; ++x)
							{
								*pbtWrite = min(*pbtWrite, pAlpha[x]);
								pbtWrite += 4;
							}

							pbtData += lockedBitmapData.Stride;
						}

						bmp.UnlockBits(&lockedBitmapData);
					}

					graphics.DrawImage(
						&bmp, rcDraw,
						0, 0, rcEnter.Width() + VSI_LOGIN_FADE_CX, rcEnter.Height(),
						Gdiplus::UnitPixel);

				#else  // VSI_EFFECTS_SOFT

					Gdiplus::Region rgnClip;
					graphics.GetClip(&rgnClip);
					graphics.SetClip(rcDraw);

					rcDraw.Offset(-iOffset, 0);

					// Set the opacity
					Gdiplus::ColorMatrix cm = {
						1.0f, 0.0f, 0.0f, 0.0f, 0.0f,
						0.0f, 1.0f, 0.0f, 0.0f, 0.0f,
						0.0f, 0.0f, 1.0f, 0.0f, 0.0f,
						0.0f, 0.0f, 0.0f, (float)(m_stateEnter.m_dbAlpha * m_stateEnter.m_dbAlpha), 0.0f,
						0.0f, 0.0f, 0.0f, 0.0f, 1.0f,
					};

					Gdiplus::ImageAttributes ias;
					ias.SetColorMatrix(
						&cm,
						Gdiplus::ColorMatrixFlagsDefault,
						Gdiplus::ColorAdjustTypeBitmap);

					int iW = m_pEnter->GetWidth() / 4;
					rcDraw.Offset(-(int)(iW * (1.0 - m_stateEnter.m_dbPosition)), 0);

					graphics.DrawImage(
						m_pEnter.get(), rcDraw,
						0, 0, iW, m_pEnter->GetHeight(),
						Gdiplus::UnitPixel,
						&ias);
					graphics.SetClip(&rgnClip);

				#endif  // VSI_EFFECTS_SOFT
			}
			break;
		}

		// Time and Date
		int iRet = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLONGDATE, NULL, 0);
		if (0 < iRet)
		{
			CString strDateFormat;
			LPWSTR pszDateFormat = strDateFormat.GetBufferSetLength(iRet);

			iRet = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLONGDATE, pszDateFormat, iRet);
			if (0 < iRet)
			{
				strDateFormat.ReleaseBuffer();

				strDateFormat.Remove(L'y');
				strDateFormat.Trim();
				// Remove symbols on left
				WCHAR test = strDateFormat.GetAt(0);
				while (test != L'M' && test != L'd')
				{
					strDateFormat.TrimLeft(test);
					if (strDateFormat.IsEmpty())
					{
						break;
					}
					test = strDateFormat.GetAt(0);
				}
				if (!strDateFormat.IsEmpty())
				{
					// Remove symbols on right
					test = strDateFormat.GetAt(strDateFormat.GetLength() - 1);
					while (test != L'M' && test != L'd')
					{
						strDateFormat.TrimRight(test);
						if (strDateFormat.IsEmpty())
						{
							break;
						}
						test = strDateFormat.GetAt(strDateFormat.GetLength() - 1);
					}
				}

				SYSTEMTIME stLocal;
				GetLocalTime(&stLocal);

				WCHAR szDate[50], szTime[50];
				GetDateFormat(LOCALE_USER_DEFAULT, 0, &stLocal, strDateFormat, szDate, _countof(szDate));
				GetTimeFormat(LOCALE_USER_DEFAULT, TIME_NOSECONDS, &stLocal, NULL, szTime, _countof(szTime));

				Gdiplus::FontFamily fontFamily(VSI_FONT_LOGIN);
				Gdiplus::Font fontM(&fontFamily, VSI_FONT_SIZE_LOGIN_DATE, Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);
				Gdiplus::Font fontL(&fontFamily, VSI_FONT_SIZE_LOGIN_TIME, Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);

				CRect rcTime;
				GetDlgItem(IDC_LOGIN_TIME).GetWindowRect(&rcTime);
				ScreenToClient(&rcTime);

				CRect rcDate;
				GetDlgItem(IDC_LOGIN_DATE).GetWindowRect(&rcDate);
				ScreenToClient(&rcDate);

				Gdiplus::RectF rcTimeF(
					(Gdiplus::REAL)rcTime.left,
					(Gdiplus::REAL)rcTime.top,
					(Gdiplus::REAL)rcTime.Width(),
					(Gdiplus::REAL)rcTime.Height());

				Gdiplus::StringFormat sf(Gdiplus::StringFormatFlagsNoWrap);

				if (stLocal.wSecond & 1)
				{
					LPCWSTR pszSept = wcschr(szTime, L':');
					if (NULL != pszSept)
					{
						int iPos = pszSept - szTime;

						Gdiplus::CharacterRange charRange(iPos, 1);
						sf.SetMeasurableCharacterRanges(1, &charRange);

						Gdiplus::Region region;

						if (Gdiplus::Ok == graphics.MeasureCharacterRanges(
								szTime, -1, &fontL, rcTimeF, &sf, 1, &region))
						{
							graphics.IntersectClip(&region);

							Gdiplus::Color colorDots(
								150,
								GetRValue(VSI_COLOR_LOGIN_TIME),
								GetGValue(VSI_COLOR_LOGIN_TIME),
								GetBValue(VSI_COLOR_LOGIN_TIME));
							Gdiplus::SolidBrush brushDots(colorDots);
							graphics.DrawString(
								szTime,
								-1,
								&fontL,
								rcTimeF,
								&sf,
								&brushDots);

							graphics.ResetClip();
							graphics.ExcludeClip(&region);
						}
					}
				}

				Gdiplus::SolidBrush brushTime(Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_LOGIN_TIME)));

				graphics.DrawString(
					szTime,
					-1,
					&fontL,
					rcTimeF,
					&sf,
					&brushTime);

				Gdiplus::SolidBrush brushDate(Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_LOGIN_DATE)));

				graphics.DrawString(
					szDate,
					-1,
					&fontM,
					Gdiplus::RectF(
						(Gdiplus::REAL)rcDate.left,
						(Gdiplus::REAL)rcDate.top,
						(Gdiplus::REAL)rcDate.Width(),
						(Gdiplus::REAL)rcDate.Height()),
					NULL,
					&brushDate);
			}
		}

		// Version
		{
			Gdiplus::FontFamily fontFamily(VSI_FONT_LOGIN);
			Gdiplus::Font font(&fontFamily, VSI_FONT_SIZE_LOGIN_VERSION, Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);

			CRect rcVer;
			GetDlgItem(IDC_LOGIN_VER).GetWindowRect(&rcVer);
			ScreenToClient(&rcVer);

			CString strVersion;
			strVersion.Format(L"Version %d.%d.%d.%d", VSI_SOFTWARE_MAJOR, VSI_SOFTWARE_MIDDLE, VSI_SOFTWARE_MINOR, VSI_BUILD_NUMBER);

// Shows service mode password in debug build
#ifdef _DEBUG
			if (!m_strPassword.IsEmpty())
			{
				strVersion += L"          ";
				strVersion += m_strPassword;
			}
#endif  // _DEBUG

			Gdiplus::SolidBrush solidBrush(Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_LOGIN_VERSION)));

			graphics.DrawString(strVersion,
				-1,
				&font,
				Gdiplus::RectF(
					(Gdiplus::REAL)rcVer.left,
					(Gdiplus::REAL)rcVer.top,
					(Gdiplus::REAL)rcVer.Width(),
					(Gdiplus::REAL)rcVer.Height()),
				NULL,
				&solidBrush);
		}

		// Warn / error message
		if (!m_strMsg.IsEmpty())
		{
			HFONT hFont;
			VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
			Gdiplus::Font font(hdc, hFont);

			CRect rcMsg;
			GetDlgItem(IDC_LOGIN_MSG).GetWindowRect(&rcMsg);
			ScreenToClient(&rcMsg);

			if (VSI_LOGIN_MESSAGE_NONE != m_dwMsgType)
			{
				CRect rcPassword;
				GetDlgItem(IDC_LOGIN_PASSWORD_EDIT).GetWindowRect(&rcPassword);
				ScreenToClient(&rcPassword);

				enum
				{
					VSI_LOGIN_IMAGE_CAPS_LOCK = 0,
					VSI_LOGIN_IMAGE_WARNING,
					VSI_LOGIN_IMAGE_ERROR,
					VSI_LOGIN_IMAGE_MAX
				};

				int iMsgIconSrcW = m_pIcons->GetWidth() / VSI_LOGIN_IMAGE_MAX;
				int iMsgIconSrcH = m_pIcons->GetHeight();
				int iMsgIconW = (int)(rcPassword.Height() * 0.8);
				int iMsgIconH = iMsgIconW;

				int iX = 0;

				switch (m_dwMsgType)
				{
				case VSI_LOGIN_MESSAGE_CAPSLOCK_ON:
					iX = VSI_LOGIN_IMAGE_CAPS_LOCK;
					break;
				case VSI_LOGIN_MESSAGE_PASSWORD_INCORRECT:
					iX = VSI_LOGIN_IMAGE_WARNING;
					break;
				case VSI_LOGIN_MESSAGE_ACCOUNT_DISABLED:
					iX = VSI_LOGIN_IMAGE_ERROR;
					break;
				}

				iX *= iMsgIconSrcW;

				graphics.DrawImage(
					m_pIcons.get(),
					Gdiplus::RectF(
						(Gdiplus::REAL)rcMsg.left,
						(Gdiplus::REAL)rcMsg.top,
						(Gdiplus::REAL)iMsgIconW,
						(Gdiplus::REAL)iMsgIconH),
					(Gdiplus::REAL)iX, 0.0f, (Gdiplus::REAL)iMsgIconSrcW, (Gdiplus::REAL)iMsgIconSrcH,
					Gdiplus::UnitPixel);

				rcMsg.left += (int)(iMsgIconW * 1.2);

				rcMsg.top += (int)((iMsgIconH - font.GetHeight(graphics.GetDpiY())) / 2.0f);
			}

			Gdiplus::Color colorMsg(255, 255, 210);
			Gdiplus::SolidBrush solidBrush(colorMsg);

			graphics.DrawString(
				m_strMsg,
				-1,
				&font,
				Gdiplus::RectF(
					(Gdiplus::REAL)rcMsg.left,
					(Gdiplus::REAL)rcMsg.top,
					(Gdiplus::REAL)rcMsg.Width(),
					(Gdiplus::REAL)rcMsg.Height()),
				NULL,
				&solidBrush);
		}
	}
	VSI_CATCH(hr);

	if (NULL != hpb)
	{
		VsiEndBufferedPaint(hpb, TRUE);
	}

	return 0;
}

LRESULT CVsiLoginView::OnPaint(UINT, WPARAM, LPARAM, BOOL &bHandled)
{
	PAINTSTRUCT ps;

	BeginPaint(&ps);

	OnPrintClient(WM_PRINTCLIENT, (WPARAM)ps.hdc, PRF_CLIENT, bHandled);

	EndPaint(&ps);

	CheckAnimation();

	return 0;
}

LRESULT CVsiLoginView::OnTimer(UINT, WPARAM, LPARAM, BOOL&)
{
	CRect rcTime;
	GetDlgItem(IDC_LOGIN_TIME).GetWindowRect(&rcTime);
	ScreenToClient(&rcTime);

	CRect rcDate;
	GetDlgItem(IDC_LOGIN_DATE).GetWindowRect(&rcDate);
	ScreenToClient(&rcDate);

	InvalidateRect(&rcTime);
	InvalidateRect(&rcDate);

	CheckAnimation();

	return 0;
}

LRESULT CVsiLoginView::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		int iIndex = (int)m_wndName.SendMessage(CB_GETCURSEL);
		if (CB_ERR != iIndex)
		{
			CString strPassword;
			m_wndPassword.GetWindowText(strPassword);

			if (!strPassword.IsEmpty())
			{
				CComPtr<IVsiOperator> pOperator((IVsiOperator*)m_wndName.SendMessage(CB_GETITEMDATA, iIndex));

				if (S_OK == pOperator->TestPassword(strPassword))
				{
					CComHeapPtr<OLECHAR> pszOperId;
					hr = pOperator->GetId(&pszOperId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Set new current operator
					if (NULL != pszOperId)
					{
						hr = m_pOperatorMgr->SetCurrentOperator(pszOperId);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						// Add operator to the set of logged in operators
						hr = m_pOperatorMgr->SetIsAuthenticated(pszOperId, TRUE);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}

					m_pCmdMgr->InvokeCommand(ID_CMD_LOG_IN, NULL);
				}
				else
				{
					if (--m_iLoginSessionCounter > 0)
					{
						m_wndPassword.SetFocus();
						m_wndPassword.SendMessage(EM_SETSEL, 0, -1);

						SetMessage(VSI_LOGIN_MESSAGE_PASSWORD_INCORRECT, L"Password is incorrect.");
					}
					else
					{
						hr = pOperator->SetState(VSI_OPERATOR_STATE_DISABLE_SESSION);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						UpdateUI();
					}
				}
			}
		}
		else
		{
			m_wndName.SetFocus();
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiLoginView::OnBnClickedShutdown(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		// Needed the SE_SHUTDOWN_NAME privilege to shutdown
		TOKEN_PRIVILEGES privileges;
		HANDLE hToken(NULL);
		BOOL bCanShutdown(FALSE);
		if (OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken))
		{
			privileges.PrivilegeCount = 1;
			privileges.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;
			if (LookupPrivilegeValue(NULL, SE_SHUTDOWN_NAME, &privileges.Privileges[0].Luid))
			{
				bCanShutdown = AdjustTokenPrivileges(
					hToken, FALSE, &privileges, 0, NULL, NULL);
			}

			CloseHandle(hToken);
		}

		if (bCanShutdown)
		{
			CComVariant vShutdown(true);
			hr = m_pCmdMgr->InvokeCommand(ID_CMD_EXIT_BEGIN, &vShutdown);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiLoginView::OnCbSelChange(WORD, WORD, HWND, BOOL& bHandled)
{
	// Reset login retries counter
	m_iLoginSessionCounter = VSI_LOGIN_ATTEMPTS_MAX;

	// Clear password field
	m_wndPassword.SetWindowText(L"");

	SetMessage(VSI_LOGIN_MESSAGE_NONE, L"");

	UpdateCapsLockState();

	UpdateUI();

	if (m_wndPassword.IsWindowEnabled())
	{
		// Set focus to password
		m_wndPassword.SetFocus();
	}

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiLoginView::OnEnChange(WORD, WORD, HWND, BOOL& bHandled)
{
	bHandled = FALSE;

	CString strPassword;
	m_wndPassword.GetWindowText(strPassword);

	BOOL bShowOk = !strPassword.IsEmpty();
	GetDlgItem(IDOK).EnableWindow(bShowOk);
	if (bShowOk)
	{
		m_stateEnter.Show();
	}
	else
	{
		m_stateEnter.Hide();
	}

	switch (m_dwMsgType)
	{
	case VSI_LOGIN_MESSAGE_PASSWORD_INCORRECT:
		SetMessage(VSI_LOGIN_MESSAGE_NONE, L"");
		break;
	}

	UpdateCapsLockState();

	CheckAnimation();

	return 0;
}

void CVsiLoginView::FillNameCombo(bool bServiceMode)
{
	HRESULT hr = S_OK;

	try
	{
		// UI
		Cleanup();

		// Get last selected operator id
		CString strIdSelected(
			VsiGetRangeValue<CString>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_LAST_LOGGED_IN,
				m_pPdm));

		CComPtr<IVsiParameter> pParamOperator;
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, VSI_PARAMETER_SYS_LAST_LOGGED_IN,
			&pParamOperator);
		CComQIPtr<IVsiParameterRange> pOperator(pParamOperator);
		VSI_CHECK_INTERFACE(pOperator, VSI_LOG_ERROR, NULL);

		VSI_OPERATOR_FILTER_TYPE types = VSI_OPERATOR_FILTER_TYPE_STANDARD | VSI_OPERATOR_FILTER_TYPE_ADMIN | VSI_OPERATOR_FILTER_TYPE_PASSWORD;
		if (bServiceMode)
		{
			types |= VSI_OPERATOR_FILTER_TYPE_SERVICE_MODE;
		}

		int iSelect(-1);
		int iCount(0);
		do
		{
			CComPtr<IVsiEnumOperator> pOperators;
			hr = m_pOperatorMgr->GetOperators(types, &pOperators);
			if (SUCCEEDED(hr))
			{
				CComPtr<IVsiOperator> pOperator;
				while (pOperators->Next(1, &pOperator, NULL) == S_OK)
				{
					CComHeapPtr<OLECHAR> pszName;
					hr = pOperator->GetName(&pszName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComHeapPtr<OLECHAR> pszId;
					hr = pOperator->GetId(&pszId);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					int iIndex = (int)m_wndName.SendMessage(CB_ADDSTRING, 0, (LPARAM)pszName.m_pData);
					if (CB_ERR != iIndex)
					{
						// Set the operator as part of the item data
						int iRet = (int)m_wndName.SendMessage(CB_SETITEMDATA, iIndex, (LPARAM)pOperator.p);
						_ASSERT(CB_ERR != iRet);
						pOperator.Detach();

						++iCount;
					}

					if (strIdSelected == pszId)
					{
						// Select operator
						iSelect = iIndex;
						m_wndName.SendMessage(CB_SETCURSEL, iSelect);
					}

					pOperator.Release();
				}

				// Select 1st operator
				if ((0 > iSelect)  || bServiceMode)
				{
					m_wndName.SendMessage(CB_SETCURSEL, 0);
				}
			}

			if (0 == iCount)
			{
				// No operators, switch to service mode
				types |= VSI_OPERATOR_FILTER_TYPE_SERVICE_MODE;
			}
		}
		while (0 == iCount);
	}
	VSI_CATCH(hr);
}

void CVsiLoginView::SetMessage(VSI_LOGIN_MESSAGE_TYPE type, LPCWSTR pszMsg)
{
	if (m_dwMsgType != type && m_strMsg != pszMsg)
	{
		m_dwMsgType = type;
		m_strMsg = pszMsg;

		Invalidate();
	}
}

void CVsiLoginView::UpdateUI()
{
	HRESULT hr(S_OK);

	try
	{
		BOOL bPasswordEnabled(FALSE);
		BOOL bOkBtnEnabled(TRUE);

		int index = (int)m_wndName.SendMessage(CB_GETCURSEL);
		if (CB_ERR != index)
		{
			IVsiOperator *pOperator = (IVsiOperator*)m_wndName.SendMessage(CB_GETITEMDATA, index);

			VSI_OPERATOR_STATE dwState(VSI_OPERATOR_STATE_DISABLE);
			hr = pOperator->GetState(&dwState);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if ((VSI_OPERATOR_STATE_ENABLE == dwState) && (S_OK == pOperator->HasPassword()) )
			{
				bPasswordEnabled = TRUE;
			}

			CString strPassword;
			m_wndPassword.GetWindowText(strPassword);

			if (strPassword.IsEmpty())
			{
				bOkBtnEnabled = FALSE;
			}

			m_wndPassword.EnableWindow(bPasswordEnabled);

			if (bPasswordEnabled)
			{
				m_wndPassword.ModifyStyle(0, ES_PASSWORD, 0);
			}
			else
			{
				m_wndPassword.SetWindowText(L"");
				m_wndPassword.ModifyStyle(ES_PASSWORD, 0, 0);
				m_wndPassword.SetWindowText(L"Account Disabled");

				m_wndName.SetFocus();

				SetMessage(VSI_LOGIN_MESSAGE_ACCOUNT_DISABLED,
					L"Account disabled due to multiple failed login attempts.\nPlease contact administrator for assistance.");
			}
		}

		GetDlgItem(IDOK).EnableWindow(bPasswordEnabled && bOkBtnEnabled);
		if (bPasswordEnabled && bOkBtnEnabled)
		{
			m_stateEnter.Show();
		}
		else
		{
			m_stateEnter.Hide();
		}

		CheckAnimation();
	}
	VSI_CATCH(hr);
}

void CVsiLoginView::Cleanup()
{
	int iCount = (int)m_wndName.SendMessage(CB_GETCOUNT, 0, 0);
	for (int i = 0; i < iCount; ++i)
	{
		IVsiOperator *pOperator = (IVsiOperator*)m_wndName.SendMessage(CB_GETITEMDATA, i);
		if (NULL != pOperator)
		{
			pOperator->Release();
		}
	}

	m_wndName.SendMessage(CB_RESETCONTENT);
}

void CVsiLoginView::UpdateCapsLockState()
{
	if (GetKeyState(VK_CAPITAL) & 1)
	{
		if (VSI_LOGIN_MESSAGE_ACCOUNT_DISABLED != m_dwMsgType && VSI_LOGIN_MESSAGE_PASSWORD_INCORRECT != m_dwMsgType)
		{
			SetMessage(VSI_LOGIN_MESSAGE_CAPSLOCK_ON, L"Caps Lock is on.");
		}
	}
	else
	{
		if (VSI_LOGIN_MESSAGE_CAPSLOCK_ON == m_dwMsgType)
		{
			SetMessage(VSI_LOGIN_MESSAGE_NONE, L"");
		}
	}
}

void CVsiLoginView::CheckAnimation()
{
	CWindow wndEnter(GetDlgItem(IDOK));

	bool bRedraw(false);

	CVsiBtnState::VSI_STATE stateOld = m_stateEnter.GetState();
	CVsiBtnState::VSI_STATE state = m_stateEnter.Tick();
	switch (state)
	{
	case CVsiBtnState::VSI_STATE_HIDDEN:
		{
			bRedraw = (state != stateOld);
		}
		break;
	case CVsiBtnState::VSI_STATE_SHOWN:
		{
			if (!wndEnter.IsWindowVisible())
			{
				wndEnter.ShowWindow(SW_SHOWNA);
			}
			bRedraw = (state != stateOld);
		}
		break;
	case CVsiBtnState::VSI_STATE_HIDING:
		{
			if (wndEnter.IsWindowVisible())
			{
				wndEnter.ShowWindow(SW_HIDE);
			}
			bRedraw = true;
		}
		break;
	case CVsiBtnState::VSI_STATE_SHOWING:
		{
			bRedraw = true;
		}
		break;
	}

	if (bRedraw)
	{
		CRect rc;
		wndEnter.GetWindowRect(&rc);
		ScreenToClient(&rc);

		#ifdef VSI_EFFECTS_SOFT
			rc.left -= VSI_LOGIN_FADE_CX;
		#endif // VSI_EFFECTS_SOFT
		InvalidateRect(&rc);
	}
}