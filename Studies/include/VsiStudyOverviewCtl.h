/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudyOverviewCtl.h
**
**	Description:
**		Study summary control
**
********************************************************************************/

#pragma once


#include <VsiGdiPlus.h>
#include <VsiWindow.h>
#include <VsiGroupList.h>
#include <VsiScrollImpl.h>
#include <VsiThemeColor.h>
#include <VsiDrawBuffer.h>
#include <VsiStudyModule.h>


class CVsiStudyOverviewCtl :
	public CWindowImpl<CVsiStudyOverviewCtl, CVsiWindow>,
	public CVsiScrollImpl<CVsiStudyOverviewCtl>,
	public CVsiDrawBuffer<CVsiStudyOverviewCtl>,
	public CVsiThread<CVsiStudyOverviewCtl>
{
	friend CVsiThread<CVsiStudyOverviewCtl>;
	typedef CWindowImpl<CVsiStudyOverviewCtl, CVsiWindow> baseClass;

public:
	// Constants
	typedef enum
	{
		VSI_STUDY_OVERVIEW_DISPLAY_MINIMAL = 0,
		VSI_STUDY_OVERVIEW_DISPLAY_DETAILS,
	} VSI_STUDY_OVERVIEW_DISPLAY;

private:
	// Constants
	enum
	{
		// Update timer message
		VSI_WM_STUDY_OVERVIEW_UPDATE = (WM_USER + 100),
	};

	CCriticalSection m_cs;

	VSI_STUDY_OVERVIEW_DISPLAY m_display;

	bool m_bVertical;
	CSize m_size;
	CString m_strTextHighlight;

	BOOL m_bCancel;

	VSI_PROP_SERIES m_sortBy;
	BOOL m_bDescending;

	CVsiGroupListData m_GroupListData;
	CVsiGroupListLayout m_GroupListLayout;
	CVsiGroupListDraw m_GroupListDraw;

	class CVsiStudyInfo
	{
	public:
		CComPtr<IVsiStudy> m_pStudy;
		CVsiGroupListData *m_pGroupListData;

		enum
		{
			VSI_STUDY_HEADER = 0,
			VSI_STUDY_OWNER,
			VSI_STUDY_DATE,
			VSI_STUDY_NOTES,
			VSI_STUDY_SERIES,
			VSI_STUDY_SIZE,
			VSI_STUDY_MAX,
		};

		int m_piRowIndex[VSI_STUDY_MAX];

		CVsiStudyInfo(CVsiGroupListData *pGroupListData) :
			m_pGroupListData(pGroupListData)
		{
			ZeroMemory(m_piRowIndex, sizeof(m_piRowIndex));
		}

		HRESULT Update(BOOL &bCancel, VSI_STUDY_OVERVIEW_DISPLAY display)
		{
			HRESULT hr(S_OK);

			try
			{
				m_piRowIndex[VSI_STUDY_HEADER] = m_pGroupListData->InsertHeader(0, NULL, NULL, false);
				m_piRowIndex[VSI_STUDY_OWNER] = m_pGroupListData->InsertRow(L"Owner", NULL, true);
				m_piRowIndex[VSI_STUDY_DATE] = m_pGroupListData->InsertRow(L"Date", NULL, true);
				m_piRowIndex[VSI_STUDY_NOTES] = m_pGroupListData->InsertRow(L"Notes", NULL, true, true);
				m_piRowIndex[VSI_STUDY_SERIES] = m_pGroupListData->InsertRow(L"Series", NULL, false);
				m_piRowIndex[VSI_STUDY_SIZE] = m_pGroupListData->InsertRow(L"Size", NULL, false);

				// Study Date
				if (!bCancel)
				{
					if (VSI_STUDY_OVERVIEW_DISPLAY_DETAILS == display)
					{
						CComVariant vDate;
						hr =  m_pStudy->GetProperty(VSI_PROP_STUDY_CREATED_DATE, &vDate);
						if (S_OK == hr)
						{
							COleDateTime date(vDate);

							SYSTEMTIME stCreate;
							date.GetAsSystemTime(stCreate);

							// The time is stored as UTC. Convert it to local time for display.
							BOOL bRet = SystemTimeToTzSpecificLocalTime(NULL, &stCreate, &stCreate);
							if (bRet)
							{
								WCHAR szDate[100];
								*szDate = 0;

								GetDateFormat(
									LOCALE_USER_DEFAULT,
									DATE_SHORTDATE,
									&stCreate,
									NULL,
									szDate,
									_countof(szDate));

								m_pGroupListData->SetItemText(
									m_piRowIndex[CVsiStudyInfo::VSI_STUDY_DATE],
									szDate);
							}
						}
					}
				}

				// Study Notes
				if (!bCancel)
				{
					CString strText;

					CComVariant vNotes;
					hr = m_pStudy->GetProperty(VSI_PROP_STUDY_NOTES, &vNotes);
					if (S_OK == hr)
					{
						strText = V_BSTR(&vNotes);
					}

					m_pGroupListData->SetItemText(
						m_piRowIndex[CVsiStudyInfo::VSI_STUDY_NOTES],
						strText);
				}

				// Series
				if (!bCancel)
				{
					LONG lSeries(0);
					hr = m_pStudy->GetSeriesCount(&lSeries, &bCancel);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (!bCancel)
					{
						CString strText;
						strText.Format(L"%d", lSeries);

						m_pGroupListData->SetItemText(m_piRowIndex[VSI_STUDY_SERIES], strText);
					}
				}

				// Study size
				if (!bCancel)
				{
					CComHeapPtr<OLECHAR> pszPath;
					hr = m_pStudy->GetDataPath(&pszPath);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					PathRemoveFileSpec(pszPath);

					__int64 iSize = GetDirectorySize(pszPath, &bCancel);

					if (!bCancel)
					{
						CString strSize;
						FormatSize(strSize, iSize);
						m_pGroupListData->SetItemText(m_piRowIndex[VSI_STUDY_SIZE], strSize);
					}
				}
			}
			VSI_CATCH(hr);

			return hr;
		}
	} m_studyInfo;

	class CVsiSeriesInfo
	{
	public:
		CComPtr<IVsiSeries> m_pSeries;
		CVsiGroupListData *m_pGroupListData;

		enum
		{
			VSI_SERIES_HEADER = 0,
			VSI_SERIES_ACQ_BY,
			VSI_SERIES_DATE,
			VSI_SERIES_ANIMAL_ID,
			VSI_SERIES_PROTOCOL_ID,
			VSI_SERIES_NOTES,
			VSI_SERIES_IMAGES,
			VSI_SERIES_SIZE,
			VSI_SERIES_MAX,
		};

		int m_piRowIndex[VSI_SERIES_MAX];

		CVsiSeriesInfo(CVsiGroupListData *pGroupListData) :
			m_pGroupListData(pGroupListData)
		{
		}

		HRESULT Update(IVsiSeries *pSeries, BOOL &bCancel)
		{
			HRESULT hr(S_OK);

			try
			{
				m_pSeries = pSeries;

				m_piRowIndex[VSI_SERIES_HEADER] = m_pGroupListData->InsertHeader(1, NULL, NULL, false);

				// Acquired By
				if (!bCancel)
				{
					CString strAcquiredBy;

					CComVariant vAcquiredBy;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_ACQUIRED_BY, &vAcquiredBy);
					if (S_OK == hr && 0 != *V_BSTR(&vAcquiredBy))
					{
						strAcquiredBy = V_BSTR(&vAcquiredBy);
					}

					m_piRowIndex[VSI_SERIES_ACQ_BY] = m_pGroupListData->InsertRow(L"Acquired By", strAcquiredBy, true);
				}

				// Date
				if (!bCancel)
				{
					CString strDate;

					CComVariant vDate;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_CREATED_DATE, &vDate);
					if (S_OK == hr)
					{
						COleDateTime date(vDate);

						SYSTEMTIME stCreate;
						date.GetAsSystemTime(stCreate);

						// The time is stored as UTC. Convert it to local time for display.
						BOOL bRet = SystemTimeToTzSpecificLocalTime(NULL, &stCreate, &stCreate);
						if (bRet)
						{
							WCHAR szDate[100];
							*szDate = 0;

							GetDateFormat(
								LOCALE_USER_DEFAULT,
								DATE_SHORTDATE,
								&stCreate,
								NULL,
								szDate,
								_countof(szDate));

							strDate = szDate;
						}
					}

					m_piRowIndex[VSI_SERIES_DATE] = m_pGroupListData->InsertRow(L"Date", strDate, true);
				}

				// Animal ID
				if (!bCancel)
				{
					LPWSTR pszAnimalId(NULL);

					CComVariant vAnimalId;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_ANIMAL_ID, &vAnimalId);
					if (S_OK == hr && 0 != V_BSTR(&vAnimalId))
					{
						pszAnimalId = V_BSTR(&vAnimalId);
					}

					m_piRowIndex[VSI_SERIES_ANIMAL_ID] = m_pGroupListData->InsertRow(L"Animal ID", pszAnimalId, true);
				}

				// Protocol ID
				if (!bCancel)
				{
					LPWSTR pszProtocolId(NULL);

					CComVariant vProtocolId;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_CUSTOM4, &vProtocolId);
					if (S_OK == hr && 0 != V_BSTR(&vProtocolId))
					{
						pszProtocolId = V_BSTR(&vProtocolId);
					}

					m_piRowIndex[VSI_SERIES_PROTOCOL_ID] = m_pGroupListData->InsertRow(L"Protocol ID", pszProtocolId, true);
				}

				// Notes
				if (!bCancel)
				{
					LPWSTR pszNote(NULL);

					CComVariant vNotes;
					hr = pSeries->GetProperty(VSI_PROP_SERIES_NOTES, &vNotes);
					if (S_OK == hr && 0 != V_BSTR(&vNotes))
					{
						pszNote = V_BSTR(&vNotes);
					}

					m_piRowIndex[VSI_SERIES_NOTES] = m_pGroupListData->InsertRow(L"Notes", pszNote, true, true);
				}

				// Images
				if (!bCancel)
				{
					CString strImages;

					LONG lImages(0);
					hr = pSeries->GetImageCount(&lImages, &bCancel);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (!bCancel)
					{
						strImages.Format(L"%d", lImages);
					}

					m_piRowIndex[VSI_SERIES_IMAGES] = m_pGroupListData->InsertRow(L"Images", strImages, false);
				}

				// Series size
				if (!bCancel)
				{
					CString strSize;

					CComHeapPtr<OLECHAR> pszPath;
					hr = pSeries->GetDataPath(&pszPath);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					PathRemoveFileSpec(pszPath);

					__int64 iSize = GetDirectorySize(pszPath, &bCancel);
					if (!bCancel)
					{
						FormatSize(strSize, iSize);
					}

					m_piRowIndex[VSI_SERIES_SIZE] = m_pGroupListData->InsertRow(L"Size", strSize, false);
				}
			}
			VSI_CATCH(hr);

			return hr;
		}
	};

	std::vector<CVsiSeriesInfo> m_vecSeriesInfo;
	CCriticalSection m_csSeriesInfo;

public:
	CVsiStudyOverviewCtl() :
		m_bVertical(true),
		m_size(0, 0),
		m_bCancel(false),
		m_sortBy(VSI_PROP_SERIES_CREATED_DATE),
		m_bDescending(FALSE),
		m_studyInfo(&m_GroupListData)
	{
		SetScrollExtendedStyle(VSI_SCRL_REDRAW_ALL, VSI_SCRL_REDRAW_ALL);

		CVsiGroupListDraw::CVsiLevelStyle style0(
			Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_STUDY_LEFT)),
			Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_STUDY_RIGHT)));
		m_GroupListDraw.m_levelStyle.push_back(style0);

		CVsiGroupListDraw::CVsiLevelStyle style1(
			Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_SERIES_LEFT)),
			Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_SERIES_RIGHT)));
		m_GroupListDraw.m_levelStyle.push_back(style1);
	}

	~CVsiStudyOverviewCtl()
	{
	}

	void Initialize(VSI_STUDY_OVERVIEW_DISPLAY display = VSI_STUDY_OVERVIEW_DISPLAY_MINIMAL)
	{
		m_display = display;
	}

	void Uninitialize()
	{
	}

	void SetStudy(
		IVsiStudy *pStudy,
		VSI_PROP_SERIES sortBy,
		BOOL bDescending,
		bool bForceUpdate = false)
	{
		HRESULT hr(S_OK);

		try
		{
			bool bSame = false;

			if (!bForceUpdate && NULL != m_studyInfo.m_pStudy && NULL != pStudy)
			{
				int iCmp(0);
				hr = m_studyInfo.m_pStudy->Compare(pStudy, VSI_PROP_STUDY_NS, &iCmp);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				bSame = (0 == iCmp);
			}

			if (!bSame)
			{
				bool bStopped = Stop();

				{
					CCritSecLock cs(m_cs);

					m_GroupListLayout.Clear();
					m_GroupListData.Clear();
				}

				{
					CCritSecLock cs(m_csSeriesInfo);

					m_vecSeriesInfo.clear();
				}

				m_studyInfo.m_pStudy.Release();

				if (NULL != pStudy)
				{
					hr = pStudy->Clone(&m_studyInfo.m_pStudy);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				m_sortBy = sortBy;
				m_bDescending = bDescending;

				PostMessage(VSI_WM_STUDY_OVERVIEW_UPDATE);

				if (m_studyInfo.m_pStudy && bStopped)
				{
					m_bCancel = FALSE;
					CVsiThread::RunThread(L"Study Overview");
				}
			}
		}
		VSI_CATCH(hr);
	}

	bool Stop()
	{
		if (S_OK == CVsiThread::IsThreadRunning())
		{
			m_bCancel = TRUE;
			CVsiThread::StopThread(50000);
		}

		return (S_OK != CVsiThread::IsThreadRunning());
	}

	void Refresh()
	{
		UpdateLayout();
	}

	void SetLayout(bool bVertical, bool bRedraw)
	{
		if (bVertical != m_bVertical)
		{
			m_bVertical = bVertical;

			if (bRedraw)
			{
				UpdateLayout();
			}
		}
	}

	void SetTextHighlight(LPCWSTR pszText)
	{
		m_strTextHighlight = pszText;

		Redraw();

		Invalidate();
	}

// CVsiScrollImpl
	void ScrollRedraw(UINT flags)
	{
		Redraw();

		if (GetCurrentThreadId() != GetWindowThreadID())
		{
			flags &= ~(RDW_ERASENOW | RDW_UPDATENOW);
		}

		RedrawWindow(NULL, NULL, flags);
	}

protected:
	// Message map and handlers
	BEGIN_MSG_MAP(CVsiStudyOverviewCtl)
		CHAIN_MSG_MAP(CVsiDrawBuffer<CVsiStudyOverviewCtl>)

		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_NCHITTEST, OnNcHitTest)
		MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtondown)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetKillFocus)
		MESSAGE_HANDLER(WM_KILLFOCUS, OnSetKillFocus)
		MESSAGE_HANDLER(WM_SIZE, OnSize)
		MESSAGE_HANDLER(WM_NCPAINT, OnNcPaint)
		MESSAGE_HANDLER(WM_PRINTCLIENT, OnPrintClient)
		MESSAGE_HANDLER(WM_PAINT, OnPaint)
		MESSAGE_HANDLER(VSI_WM_STUDY_OVERVIEW_UPDATE, OnUpdate)

		CHAIN_MSG_MAP(CVsiScrollImpl)
	END_MSG_MAP()

	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
	{
		Stop();

		return 0;
	}

	LRESULT OnNcHitTest(UINT uiMsg, WPARAM wp, LPARAM lp, BOOL &)
	{
		return ::DefWindowProc(*this, uiMsg, wp, lp);
	}

	LRESULT OnLButtondown(UINT, WPARAM, LPARAM, BOOL &bHandled)
	{
		if (IsWindowEnabled())
		{
			SetFocus();
		}

		bHandled = FALSE;

		return 0;
	}

	LRESULT OnSetKillFocus(UINT, WPARAM, LPARAM, BOOL&)
	{
		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE);

		return 0;
	}

	LRESULT OnSize(UINT, WPARAM, LPARAM lp, BOOL &bHandled)
	{
		CSize size(LOWORD(lp), HIWORD(lp));

		if (m_bVertical ? (size.cx != m_size.cx) : (size.cy != m_size.cy))
		{
			m_sizePage.cx = m_sizeClient.cx = size.cx;
			m_sizePage.cy = m_sizeClient.cy = size.cy;

			UpdateLayout();
		}
		else
		{
			Redraw();
			RedrawWindow(NULL, NULL, RDW_INVALIDATE);
		}

		m_size = size;

		bHandled = FALSE;

		return 0;
	}

	LRESULT OnNcPaint(UINT, WPARAM wp, LPARAM, BOOL &)
	{
		// Paint scrollbars
		DWORD dwStyle = GetStyle();
		if ((dwStyle & WS_VSCROLL) || (dwStyle & WS_HSCROLL))
		{
			DefWindowProc();
		}

		VsiNcPaintHandler(*this, wp);

		return 0;
	}

	LRESULT OnPrintClient(UINT, WPARAM wp, LPARAM, BOOL&)
	{
		DrawBuffer(reinterpret_cast<HDC>(wp));

		return 0;
	}

	LRESULT OnPaint(UINT, WPARAM, LPARAM, BOOL&)
	{
		DrawBuffer();

		return 0;
	}

	LRESULT OnUpdate(UINT, WPARAM, LPARAM, BOOL &)
	{
		UpdateLayout();

		return 0;
	}

private:
	void UpdateLayout()
	{
		HRESULT hr(S_OK);
		int iXYmax(0);

		CRect rc;
		GetClientRect(&rc);

		m_size.cx = rc.right;
		m_size.cy = rc.bottom;

		if (0 < m_GroupListData.m_vecRows.size())
		{
			HDC hdc = GetDC();
			{
				Gdiplus::Graphics graphic(hdc);

				HFONT hFont = (HFONT)SendMessage(WM_GETFONT);
				if (NULL == hFont)
				{
					VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
				}
				Gdiplus::Font font(hdc, hFont);
				Gdiplus::StringFormat format;

				CRect rcLayout(rc);

				{
					CCritSecLock cs(m_cs);
					{
						CCritSecLock cs(m_csSeriesInfo);

						// Study Name
						{
							CComVariant vName;
							hr = m_studyInfo.m_pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vName);
							if (S_OK == hr)
							{
								m_GroupListData.SetItemText(
									m_studyInfo.m_piRowIndex[CVsiStudyInfo::VSI_STUDY_HEADER],
									V_BSTR(&vName));
							}
						}

						// Study Owner
						if (VSI_STUDY_OVERVIEW_DISPLAY_DETAILS == m_display)
						{
							CComVariant vOwner;
							hr = m_studyInfo.m_pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
							if (S_OK == hr && 0 != *V_BSTR(&vOwner))
							{
								m_GroupListData.SetItemText(
									m_studyInfo.m_piRowIndex[CVsiStudyInfo::VSI_STUDY_OWNER],
									V_BSTR(&vOwner));
							}
						}

						// Series
						{
							for (auto iterSeries = m_vecSeriesInfo.begin(); iterSeries != m_vecSeriesInfo.end(); ++iterSeries)
							{
								auto &si = *iterSeries;

								// Name
								{
									CString strName;
									CComVariant vName;
									hr = si.m_pSeries->GetProperty(VSI_PROP_SERIES_NAME, &vName);
									if (S_OK == hr)
									{
										strName = V_BSTR(&vName);
									}

									m_GroupListData.SetItemText(
										si.m_piRowIndex[CVsiSeriesInfo::VSI_SERIES_HEADER],
										strName);
								}
							}
						}

						iXYmax = m_GroupListLayout.UpdateLayout(
							m_GroupListData, graphic, font, rcLayout, 24.0f, 16.0f, m_bVertical);
					}
				}
			}
			ReleaseDC(hdc);

			UpdateScrollbar(iXYmax + 4);
		}

		Redraw();
		RedrawWindow(NULL, NULL, RDW_INVALIDATE);
	}

	void UpdateScrollbar(int iXY)
	{
		CRect rc;
		GetClientRect(&rc);

		if (m_bVertical)
		{
			rc.bottom = iXY;
		}
		else
		{
			rc.right = iXY;
		}

		if (rc.right > 0 && rc.bottom > 0)
		{
			CVsiScrollImpl::SetScrollSize(rc.right, rc.bottom, FALSE, FALSE);
			CVsiScrollImpl::GetSystemSettings();
		}
	}

private:
	// CVsiThread
	DWORD ThreadProc()
	{
		HRESULT hr(S_OK);

		try
		{
			{
				CCritSecLock cs(m_cs);
				hr = m_studyInfo.Update(m_bCancel, m_display);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			if (!m_bCancel)
			{
				CComPtr<IVsiEnumSeries> pEnum;
				hr = m_studyInfo.m_pStudy->GetSeriesEnumerator(m_sortBy, m_bDescending, &pEnum, &m_bCancel);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					CComPtr<IVsiSeries> pSeries;
					while (!m_bCancel && pEnum->Next(1, &pSeries, NULL) == S_OK)
					{
						CComVariant vNs;
						hr = pSeries->GetProperty(VSI_PROP_SERIES_NS, &vNs);
						if (S_OK == hr)
						{
							CCritSecLock cs(m_cs);

							CVsiSeriesInfo info(&m_GroupListData);
							hr = info.Update(pSeries, m_bCancel);
							if (S_OK == hr && !m_bCancel)
							{
								CCritSecLock csSeriesInfo(m_csSeriesInfo);
								m_vecSeriesInfo.push_back(info);
								PostMessage(VSI_WM_STUDY_OVERVIEW_UPDATE);
							}
						}

						pSeries.Release();
					}
				}
			}
		}
		VSI_CATCH(hr);

		return 0;
	}

	static void FormatSize(CString &strSize, __int64 iSize)
	{
		if (iSize < 1024)
		{
			strSize.Format(L"%d B", static_cast<int>(iSize));
		}
		else if (iSize < (1024 * 1024))
		{
			iSize /= 1024;
			strSize.Format(L"%d KB", static_cast<int>(iSize));
		}
		else if (iSize < (1024 * 1024 * 1024))
		{
			iSize /= 1024 * 1024;
			strSize.Format(L"%d MB", static_cast<int>(iSize));
		}
		else
		{
			iSize /= 1024 * 1024 * 1024;
			strSize.Format(L"%d GB", static_cast<int>(iSize));
		}
	}

	static __int64 GetDirectorySize(LPCWSTR pszDir, BOOL *pbCancel)
	{
		HRESULT hr = S_OK;
		__int64 iSize = 0;
		HANDLE hFile = INVALID_HANDLE_VALUE;

		try
		{
			// Make sure we have a valid directory to start in.
			if (lstrlen(pszDir) > 0)
			{
				CString strPath, strFolder;
				strFolder.Format(L"%s\\*.*", pszDir);

				WIN32_FIND_DATA ff;
				hFile = FindFirstFile(strFolder, &ff);
				BOOL bWorking = (INVALID_HANDLE_VALUE != hFile);

				while (bWorking)
				{
					// Skip the dots.
					if ((0 == lstrcmp(ff.cFileName, L"."))  || (0 == lstrcmp(ff.cFileName, L"..")))
					{
						NULL;
					}
					else if (ff.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
					{
						// This is a subdirectory. Get its size and add to the total
						strPath.Format(L"%s\\%s", pszDir, ff.cFileName);
						iSize += GetDirectorySize(strPath, pbCancel);
					}
					else
					{
						LARGE_INTEGER li;  // Make sure the alignment is correct
						li.HighPart = ff.nFileSizeHigh;
						li.LowPart = ff.nFileSizeLow;
						iSize += li.QuadPart;
					}

					bWorking = !(*pbCancel) && FindNextFile(hFile, &ff);
				}

				// Done with this directory.
				FindClose(hFile);
				hFile = INVALID_HANDLE_VALUE;
			}
		}
		VSI_CATCH(hr);

		if (INVALID_HANDLE_VALUE != hFile)
		{
			FindClose(hFile);
		}

		return iSize;
	}

	void Redraw()
	{
		CCritSecLock cs(m_cs);

		CRect rc;
		GetClientRect(&rc);

		HDC hdc = CreateCompatibleDC(NULL);
		HBITMAP hbmpOld = (HBITMAP)SelectObject(hdc, GetBitmap());

		Gdiplus::Graphics graphic(hdc);

		Gdiplus::SolidBrush brush(Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_L3GRAY)));
		graphic.FillRectangle(
			&brush,
			Gdiplus::Rect(rc.left, rc.top, rc.Width(), rc.Height()));

		int iRows = m_GroupListLayout.m_vecRows.size();
		if (0 < iRows)
		{
			bool bLabelColumn = true;

			Gdiplus::SolidBrush brushLabelColumn(Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_L2P5GRAY)));
			if (m_bVertical && bLabelColumn)
			{
				graphic.FillRectangle(
					&brushLabelColumn,
					Gdiplus::Rect(rc.left, rc.top, (int)m_GroupListLayout.m_fLabelColumnCx, rc.Height()));
			}

			POINT pt = { 0, 0 };
			CVsiScrollImpl::GetScrollOffset(pt);

			graphic.TranslateTransform((float)-pt.x, (float)-pt.y);

			graphic.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);
			graphic.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);
	
			HFONT hFont = (HFONT)SendMessage(WM_GETFONT);
			if (NULL == hFont)
			{
				VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
			}
			Gdiplus::Font font(hdc, hFont);
			Gdiplus::SolidBrush brushText(Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_TEXT)));
			Gdiplus::SolidBrush brushTextHighlight(Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_TEXT_FILTER)));
			Gdiplus::SolidBrush brushLabel(Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_D1GRAY)));
			Gdiplus::StringFormat formatHeader;
			Gdiplus::StringFormat formatLabel;
			Gdiplus::StringFormat format;

			formatHeader.SetLineAlignment(Gdiplus::StringAlignmentCenter);
			formatHeader.SetFormatFlags(Gdiplus::StringFormatFlagsNoWrap);
			formatHeader.SetTrimming(Gdiplus::StringTrimmingEllipsisCharacter);
			formatLabel.SetAlignment(Gdiplus::StringAlignmentFar);
			format.SetTrimming(Gdiplus::StringTrimmingEllipsisCharacter);
			format.SetFormatFlags(Gdiplus::StringFormatFlagsLineLimit);

			Gdiplus::SolidBrush brushHeaderText(Gdiplus::Color(VSI_RGBTOARGB(VSI_COLOR_TEXT_LIGHT)));

	
			int iHeader(-1);
			float fHeaderWidth(0);
			Gdiplus::RectF rcLast;

			for (int i = 0; i < iRows; ++i)
			{
				const CVsiGroupListLayout::CVsiRowLayout &rowLayout = m_GroupListLayout.m_vecRows[i];
				const CVsiGroupListData::CVsiRow &rowData = m_GroupListData.m_vecRows[rowLayout.m_iData];

				bool bDrawLastHeader(false);

				if (CVsiGroupListData::VSI_ROW_TYPE_HEADER == rowData.m_type)
				{
					fHeaderWidth = rowLayout.m_rc.Width;

					if (m_bVertical)
					{
						if (rowLayout.m_rc.GetBottom() > pt.y)
						{
							if (0 <= iHeader)
							{
								bDrawLastHeader = true;
							}
						}
					}
					else
					{
						iHeader = i;
						bDrawLastHeader = true;
					}
				}
				else  // CVsiGroupListData::VSI_ROW_TYPE_DATA
				{
					if (m_bVertical)
					{
						if (rowLayout.m_rc.GetBottom() < pt.y)
						{
							rcLast = rowLayout.m_rc;
							continue;
						}

						if (rowLayout.m_rc.Y > (pt.y + rc.Height()))
						{
							bDrawLastHeader = true;
						}
					}
				}

				if (bDrawLastHeader)
				{
					const CVsiGroupListLayout::CVsiRowLayout &headerLayout = m_GroupListLayout.m_vecRows[iHeader];
					const CVsiGroupListData::CVsiRow &headerData = m_GroupListData.m_vecRows[headerLayout.m_iData];

					if (!m_bVertical && bLabelColumn)
					{
						graphic.FillRectangle(
							&brushLabelColumn,
							Gdiplus::Rect((int)headerLayout.m_rc.X, rc.top, (int)m_GroupListLayout.m_fLabelColumnCx, rc.Height()));
					}

					Gdiplus::RectF rcDraw(headerLayout.m_rc);
					if (m_bVertical)
					{
						if (rcDraw.Y < pt.y)
						{
							rcDraw.Y = (float)pt.y - 0.5f;
							if (rcDraw.GetBottom() > (rcLast.GetBottom() + m_GroupListLayout.m_fRowGaps + 0.5f))
							{
								rcDraw.Offset(0.0f, rcLast.GetBottom() + m_GroupListLayout.m_fRowGaps + 0.5f - rcDraw.GetBottom());
							}
						}
					}

					m_GroupListDraw.DrawHeader(graphic, font, formatHeader, brushHeaderText, rcDraw, headerData);

					if (rowLayout.m_rc.Y > (pt.y + rc.Height()))
					{
						break;
					}

					iHeader = -1;
				}

				if (CVsiGroupListData::VSI_ROW_TYPE_HEADER == rowData.m_type)
				{
					iHeader = i;
				}
				else  // CVsiGroupListData::VSI_ROW_TYPE_DATA
				{
					if (rowLayout.m_rc.GetBottom() > pt.y && rowLayout.m_rc.Y < (pt.y + rc.Height()))
					{
						// Label
						if (bLabelColumn)
						{
							Gdiplus::RectF rcLabel(rowLayout.m_rc);
							rcLabel.Width = m_GroupListLayout.m_fLabelColumnCx - m_GroupListLayout.m_fTextGaps;
							graphic.DrawString(rowData.m_strLabel, -1, &font, rcLabel, &formatLabel, &brushLabel);
						}

						// Text
						Gdiplus::SolidBrush *pBrush(&brushText);
						if (rowData.m_bCanHighlight && !m_strTextHighlight.IsEmpty())
						{
							if (NULL != StrStrI(rowData.m_strText, m_strTextHighlight))
							{
								pBrush = &brushTextHighlight;
							}
						}

						Gdiplus::RectF rcText(rowLayout.m_rc);
						rcText.X += (float)(m_GroupListLayout.m_fLabelColumnCx + m_GroupListLayout.m_fTextGaps);
						rcText.Width = fHeaderWidth - m_GroupListLayout.m_fLabelColumnCx - m_GroupListLayout.m_fTextGaps;
						graphic.DrawString(rowData.m_strText, -1, &font, rcText, &format, pBrush);
					}

					rcLast = rowLayout.m_rc;
				}
			}

			if (0 <= iHeader)
			{
				const CVsiGroupListLayout::CVsiRowLayout &headerLayout = m_GroupListLayout.m_vecRows[iHeader];
				const CVsiGroupListData::CVsiRow &headerData = m_GroupListData.m_vecRows[headerLayout.m_iData];

				Gdiplus::RectF rcDraw(headerLayout.m_rc);
				if (m_bVertical)
				{
					if (rcDraw.Y < pt.y)
					{
						rcDraw.Y = (float)pt.y - 0.5f;
						if (rcDraw.GetBottom() > (rcLast.GetBottom() + m_GroupListLayout.m_fRowGaps + 0.5f))
						{
							rcDraw.Offset(0.0f, rcLast.GetBottom() + m_GroupListLayout.m_fRowGaps + 0.5f - rcDraw.GetBottom());
						}
					}
				}

				m_GroupListDraw.DrawHeader(graphic, font, formatHeader, brushHeaderText, rcDraw, headerData);
			}
		}

		SelectObject(hdc, hbmpOld);
		DeleteDC(hdc);
	}
};

