/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsImagePreviewImpl.h
**
**	Description:
**		Declaration of the ImagePreview template
**
*******************************************************************************/


#pragma once


// preview delay timer
#define VSI_TIMER_ID_PREVIEW			1000
#define VSI_TIMER_ID_PREVIEW_ELAPSE		200
// image update timer
#define VSI_TIMER_ID_IMAGE				1001
#define VSI_TIMER_ID_IMAGE_ELAPSE		30


#define VSI_PREVIEW_LIST_CLEAR	TRUE


template<class T>
class CVsiImagePreviewImpl
{
protected:
	UINT_PTR m_uiTimer;
	CComPtr<IVsiEnumImage> m_pEnumImage;
	CWindow m_wndImageList;

public:
	CVsiImagePreviewImpl() :
		m_uiTimer(0),
		m_pEnumImage(NULL)
	{
	}

	~CVsiImagePreviewImpl()
	{
		_ASSERT(0 == m_uiTimer);
	}

	void StartPreview(BOOL bNow = FALSE)
	{
		T *pT = static_cast<T*>(this);

		_ASSERT(0 == m_uiTimer);

		if (bNow)
			m_uiTimer = pT->SetTimer(VSI_TIMER_ID_IMAGE, VSI_TIMER_ID_IMAGE_ELAPSE);
		else
			m_uiTimer = pT->SetTimer(VSI_TIMER_ID_PREVIEW, VSI_TIMER_ID_PREVIEW_ELAPSE);

		_ASSERT(0 != m_uiTimer);
	}

	void StopPreview(BOOL bClearImageList = FALSE)
	{
		T *pT = static_cast<T*>(this);

		if (bClearImageList)
			ListView_DeleteAllItems(m_wndImageList);

		if (0 != m_uiTimer)
		{
			// Kill old timer
			pT->KillTimer(m_uiTimer);
			m_uiTimer = 0;

			MSG msg;
			while (PeekMessage(&msg, pT->m_hWnd, WM_TIMER, WM_TIMER, TRUE))
				NULL;
		}

		if (m_pEnumImage != NULL)
			m_pEnumImage.Release();
	}

	void PreviewImages()
	{
		// Implemented by client
	}

protected:

	BEGIN_MSG_MAP(CVsiImagePreviewImpl)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_TIMER, OnTimer)
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL& bHandled)
	{
		bHandled = FALSE;

		_ASSERT(m_wndImageList.IsWindow());

		// Init preview control
		ListView_SetBkColor(m_wndImageList, VSI_COLOR_L2GRAY);
		ListView_SetTextBkColor(m_wndImageList, VSI_COLOR_D5GRAY);
		ListView_SetTextColor(m_wndImageList, VSI_COLOR_L2GRAY);

		return 0;
	}

	LRESULT OnTimer(UINT, WPARAM wp, LPARAM, BOOL& bHandled)
	{
		bHandled = FALSE;

		T *pT = static_cast<T*>(this);

		if (m_uiTimer == wp)
		{
			switch (m_uiTimer)
			{
			case VSI_TIMER_ID_PREVIEW:
				{
					// Kill old timer
					pT->KillTimer(m_uiTimer);
					// Create VSI_TIMER_ID_IMAGE timer
					m_uiTimer = pT->SetTimer(VSI_TIMER_ID_IMAGE, VSI_TIMER_ID_IMAGE_ELAPSE);
				}
				break;

			case VSI_TIMER_ID_IMAGE:
				{
					pT->PreviewImages();
				}
				break;
			}
		}

		return 0;
	}
};