/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiCdDvdWriterView.cpp
**
**	Description:
**		Implementation of CVsiCdDvdWriterView
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiWTL.h>
#include <VsiWindow.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiCommCtlLib.h>
#include <VsiServiceProvider.h>
#include <VsiServiceKey.h>
#include <VsiStudyModule.h>
#include "VsiImportExportXml.h"
#include "VsiCdDvdWriterView.h"
#include "VsiPathNameEditWnd.h"
#include "Dbt.h"


// Write progress timer
#define VSI_TIMER_ID_WRITE_PROGRESS				1000
#define VSI_TIMER_ID_WRITE_PROGRESS_ELAPSE      1000

#define VSI_CD_DVD_WRITER_TITLE			L"Burn CD/DVD"


// CVsiCdDvdWriterView

CVsiCdDvdWriterView::CVsiCdDvdWriterView() :
	m_uiTimerProgress(0),
	m_dwBlocksWritten(0),
	m_dwSink(0),
	m_bIdle(TRUE),
	m_bErasable(FALSE),
//	m_bAddDataLater(TRUE),
	m_bVerifyData(FALSE),
	m_bMediaRemoved(FALSE),
	m_bDeviceRemoved(FALSE),
	m_bSpaceAvailable(FALSE),
	m_bCancelWrite(FALSE),
	m_bCanFinalizeMedia(FALSE),
	m_bHasData(FALSE),
	m_bErased(FALSE),
	m_hCancelPromptWnd(NULL),
	m_bErasing(FALSE),
	m_uiCountdown(0)
{
}

CVsiCdDvdWriterView::~CVsiCdDvdWriterView()
{
	_ASSERT(m_pWriter == NULL);
	_ASSERT(m_pCmdMgr == NULL);
}

STDMETHODIMP CVsiCdDvdWriterView::Activate(
	IVsiApp *pApp,
	HWND hwndParent)
{
	HRESULT hr = S_OK;

	try
	{
		m_pApp = pApp;
		VSI_CHECK_INTERFACE(m_pApp, VSI_LOG_ERROR, NULL);

		m_pServiceProvider = m_pApp;
		VSI_CHECK_INTERFACE(m_pServiceProvider, VSI_LOG_ERROR, NULL);

		// Get Command Manager
		hr = m_pServiceProvider->QueryService(
			VSI_SERVICE_CMD_MGR,
			__uuidof(IVsiCommandManager),
			(IUnknown**)&m_pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Events sink
		CComQIPtr<IVsiEventSource> pWriterEvents(m_pWriter);
		pWriterEvents->AdviseEventSink(static_cast<IVsiEventSink*>(this));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HWND hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiCdDvdWriterView::Deactivate()
{
	HRESULT hr = S_FALSE;

	if (IsWindow())
	{
		BOOL bRet = DestroyWindow();
		_ASSERT(bRet);

		hr = S_OK;
	}

	// Unadvise
	CComQIPtr<IVsiEventSource> pWriterEvents(m_pWriter);
	pWriterEvents->UnadviseEventSink(static_cast<IVsiEventSink*>(this));

	if (m_pWriter != NULL)
	{
		m_pWriter.Release();
	}

	if (m_pXmlDoc != NULL)
		m_pXmlDoc.Release();

	if (m_pCmdMgr != NULL)
		m_pCmdMgr.Release();

	if (m_pApp != NULL)
		m_pApp.Release();

	if (m_pServiceProvider != NULL)
		m_pServiceProvider.Release();

	return hr;
}

STDMETHODIMP CVsiCdDvdWriterView::GetWindow(HWND *phWnd)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(phWnd, VSI_LOG_ERROR, L"GetWindow failed");

		*phWnd = m_hWnd;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiCdDvdWriterView::GetAccelerator(HACCEL *phAccel)
{
	UNREFERENCED_PARAMETER(phAccel);
	return E_NOTIMPL;
}

STDMETHODIMP CVsiCdDvdWriterView::GetIsBusy(
	DWORD dwStateCurrent,
	DWORD dwState,
	BOOL bTryRelax,
	BOOL *pbBusy)
{
	UNREFERENCED_PARAMETER(dwStateCurrent);
	UNREFERENCED_PARAMETER(dwState);
	UNREFERENCED_PARAMETER(bTryRelax);

	*pbBusy = FALSE;

	return S_OK;
}

STDMETHODIMP CVsiCdDvdWriterView::PreTranslateMessage(
	MSG *pMsg, BOOL *pbHandled)
{
	*pbHandled = IsDialogMessage(pMsg);

	return S_OK;
}

// IVsiCdDvdWriterView

STDMETHODIMP CVsiCdDvdWriterView::Initialize(LPCOLESTR szDriveLetter, IUnknown *pUnkCdDvdWriter, const VARIANT* pvParam)
{
	HRESULT hr = S_OK;

	try
	{
		m_pWriter = pUnkCdDvdWriter;
		VSI_CHECK_INTERFACE(m_pWriter != NULL, VSI_LOG_ERROR, NULL);

		m_pXmlDoc = V_UNKNOWN(pvParam);
		VSI_CHECK_INTERFACE(m_pXmlDoc != NULL, VSI_LOG_ERROR, NULL);

		VSI_CHECK_ARG(szDriveLetter, VSI_LOG_ERROR, NULL);
		VSI_VERIFY((1 == lstrlen(szDriveLetter)), VSI_LOG_ERROR, NULL);
		m_strDriveLetter = szDriveLetter;

		//calculate required size
		//get required size
		INT64 iTotalSize = 0;
		CComVariant vSize;
		CComPtr<IXMLDOMNodeList> pListStudy;
		hr = m_pXmlDoc->selectNodes(
			CComBSTR(L"//" VSI_STUDY_XML_ELM_STUDIES L"/" VSI_STUDY_XML_ELM_STUDY),
			&pListStudy);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting studies");

		long lLength = 0;
		hr = pListStudy->get_length(&lLength);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");
		for (long l = 0; l < lLength; l++)
		{
			CComPtr<IXMLDOMNode> pNode;
			hr = pListStudy->get_item(l, &pNode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

			CComQIPtr<IXMLDOMElement> pElemParam(pNode);

			hr = pElemParam->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_SIZE), &vSize);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_SIZE));

			hr = VariantChangeTypeEx(
				&vSize, &vSize,
				MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US), SORT_DEFAULT),
				0, VT_UI8);

			iTotalSize += V_UI8(&vSize);
		}

		m_dbRequiredSizeMB = (double)(iTotalSize / VSI_ONE_MEGABYTE);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiCdDvdWriterView::OnEvent(DWORD dwEvent, const VARIANT *pv)
{
	UNREFERENCED_PARAMETER(pv);

	::PostMessage(m_hWnd, WM_VSI_UPDATE_UI, 0, (LPARAM)dwEvent);

	return S_OK;
}

LRESULT CVsiCdDvdWriterView::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	CWaitCursor wait;

	try
	{
		// Set font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
		VsiThemeRecurSetFont(m_hWnd, hFont);

		// Layout
		CComPtr<IVsiWindowLayout> pLayout;
		HRESULT hr = pLayout.CoCreateInstance(__uuidof(VsiWindowLayout));
		if (SUCCEEDED(hr))
		{
			hr = pLayout->Initialize(m_hWnd, VSI_WL_FLAG_AUTO_RELEASE);
			if (SUCCEEDED(hr))
			{
				pLayout->AddControl(0, IDOK, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDCANCEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_ERASE_DISC_BUTTON, VSI_WL_MOVE_X);

				UINT nID = 0;
				pLayout->AddGroup(0 , VSI_WL_CENTER_X, &nID);

				pLayout->AddControl(nID, IDC_DEVICE_NAME_LABEL, VSI_WL_GROUP);
				pLayout->AddControl(nID, IDC_DEVICE_NAME_TEXT, VSI_WL_GROUP);
				pLayout->AddControl(nID, IDC_MEDIA_TYPE_LABEL, VSI_WL_GROUP);
				pLayout->AddControl(nID, IDC_MEDIA_TYPE_TEXT, VSI_WL_GROUP);
				pLayout->AddControl(nID, IDC_DISC_SIZE_LABEL, VSI_WL_GROUP);
				pLayout->AddControl(nID, IDC_DISC_SIZE_TEXT, VSI_WL_GROUP);
				pLayout->AddControl(nID, IDC_DISC_NAME_LABEL, VSI_WL_GROUP);
				pLayout->AddControl(nID, IDC_DISC_NAME_EDIT, VSI_WL_GROUP);

				pLayout->AddControl(nID, IDC_WRITE_OPTIONS_GROUP, VSI_WL_GROUP);
				pLayout->AddControl(nID, IDC_VERIFY_CHECK, VSI_WL_GROUP);
				pLayout->AddControl(nID, IDC_WRITE_SPEED_LABEL, VSI_WL_GROUP);
				pLayout->AddControl(nID, IDC_WRITE_SPEED_COMBO, VSI_WL_GROUP);

				pLayout.Detach();
			}
		}
		
		m_wndVolumeName.SubclassWindow(GetDlgItem(IDC_DISC_NAME_EDIT));
		m_wndVolumeName.SendMessage(EM_SETLIMITTEXT, 11, 0);

//		CheckDlgButton(IDC_FINALIZE_CHECK, m_bAddDataLater ? BST_CHECKED : BST_UNCHECKED);
		CheckDlgButton(IDC_VERIFY_CHECK, m_bVerifyData ? BST_CHECKED : BST_UNCHECKED);

		SetWindowText(VSI_CD_DVD_WRITER_TITLE);

		UpdateUI();
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCdDvdWriterView::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	return 0;
}

LRESULT CVsiCdDvdWriterView::OnTimer(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	DWORD dwBlocksWritten = 0;

	try
	{
		if (m_bErasing)
		{
			if (0 == m_uiCountdown)
				::SendMessage(m_hwndProgress, WM_VSI_PROGRESS_SETMSG, 0,
					(LPARAM)L"Erasing data .");
			else if (1 == m_uiCountdown)
				::SendMessage(m_hwndProgress, WM_VSI_PROGRESS_SETMSG, 0,
					(LPARAM)L"Erasing data . .");
			else if (2 == m_uiCountdown)
				::SendMessage(m_hwndProgress, WM_VSI_PROGRESS_SETMSG, 0,
					(LPARAM)L"Erasing data . . .");

			m_uiCountdown++;
			m_uiCountdown %= 3;
		}
		else
		{
			hr = m_pWriter->GetSizeWritten(&dwBlocksWritten);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (dwBlocksWritten > m_dwBlocksWritten)
			{
				m_dwBlocksWritten = dwBlocksWritten;

#ifdef _DEBUG
				ATLTRACE(L"Update Progress m_dwBlocksWritten=[%i]\n", m_dwBlocksWritten);
#endif

				SendMessage(m_hwndProgress,
					WM_VSI_PROGRESS_SETPOS, (int)(((double)m_dwBlocksWritten/m_dwBlocksToWrite)*100) , 0);
				::UpdateWindow(m_hwndProgress);

				if (m_dwBlocksWritten == m_dwBlocksToWrite)
				{
					m_uiCountdown = 1;
				}
			}
			else if (m_uiCountdown > 0)
			{
				//disable cancel when closing disc
				CWindow wndProgress(m_hwndProgress);
				wndProgress.GetDlgItem(IDCANCEL).EnableWindow(FALSE);


				SendMessage(m_hwndProgress, WM_VSI_PROGRESS_SETMSG, 0,
							(LPARAM)L"Closing disc. Please wait...");
				m_uiCountdown = 0;
			}
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCdDvdWriterView::OnProgressWindowClosed(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		if (m_uiTimerProgress != 0)
		{
			KillTimer(m_uiTimerProgress);
			m_uiTimerProgress = 0;
		}

		hr = m_pWriter->Uninitialize();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCdDvdWriterView::OnUpdateUI(UINT, WPARAM, LPARAM lp, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		CWaitCursor wait;

		if (VSI_WRITE_RESULT_SUCCESS == lp ||
			VSI_WRITE_RESULT_FAILURE == lp)
		{
			m_bIdle = TRUE;

			if (NULL != m_hCancelPromptWnd)
			{
				::DestroyWindow(m_hCancelPromptWnd);
				m_hCancelPromptWnd = NULL;
			}

			::SendMessage(m_hwndProgress, WM_CLOSE, 0, 0);

			if (m_uiTimerProgress != 0)
			{
				KillTimer(m_uiTimerProgress);
				m_uiTimerProgress = 0;
			}

			//unlock medium
			hr = m_pWriter->UnlockDevice(m_strDriveLetter);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pWriter->Uninitialize();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pWriter->Eject(m_strDriveLetter);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (VSI_WRITE_RESULT_SUCCESS == lp)
			{
				VsiMessageBox(GetTopLevelParent(), L"Write operation completed successfully",
					VSI_CD_DVD_WRITER_TITLE, MB_OK | MB_ICONINFORMATION);
			}
			else if (!m_bCancelWrite)
			{
				VsiMessageBox(GetTopLevelParent(), L"Write operation failed",
					VSI_CD_DVD_WRITER_TITLE, MB_OK | MB_ICONERROR);
			}
			else if (m_bCancelWrite)
			{
				VsiMessageBox(GetTopLevelParent(), L"Write operation cancelled",
					VSI_CD_DVD_WRITER_TITLE, MB_OK | MB_ICONINFORMATION);
			}

			m_pCmdMgr->InvokeCommand(ID_CMD_BURN_DISC_DONE, NULL);
		}
		else if (VSI_ERASE_RESULT_SUCCESS == lp ||
				 VSI_ERASE_RESULT_FAILURE == lp)
		{
			m_bIdle = TRUE;

			if (m_uiTimerProgress != 0)
			{
				KillTimer(m_uiTimerProgress);
				m_uiTimerProgress = 0;
			}

			//unlock medium
			hr = m_pWriter->UnlockDevice(m_strDriveLetter);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			::SendMessage(m_hwndProgress,
				WM_CLOSE, 0, 0);

			if (VSI_ERASE_RESULT_SUCCESS == lp)
				VsiMessageBox(GetTopLevelParent(), L"Erase operation completed successfully", VSI_CD_DVD_WRITER_TITLE, MB_APPLMODAL | MB_OK);
			else
				VsiMessageBox(GetTopLevelParent(), L"Erase operation failed", VSI_CD_DVD_WRITER_TITLE, MB_APPLMODAL | MB_OK | MB_ICONERROR);

			UpdateUI();

		}

	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiCdDvdWriterView::OnCancelWrite(UINT, WPARAM, LPARAM, BOOL&)
{
	m_hCancelPromptWnd = VsiMessageBoxModeless(
		m_hwndProgress,
		L"This action will render the media unusable.\r\n" \
			L"Would you like to cancel anyway?",
		L"Cancel Confirmation",
		MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2,
		CancelProgressCallback,
		this);

	return 0;
}

LRESULT CVsiCdDvdWriterView::OnAbortedWrite(UINT, WPARAM, LPARAM, BOOL& b)
{
	m_bCancelWrite = FALSE;
	m_pWriter->AbortWrite();
	OnUpdateUI(0, 0, VSI_WRITE_RESULT_FAILURE, b);
	return 0;
}

LRESULT CVsiCdDvdWriterView::OnStartWrite(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	int iSelSpeedItemIndex = 0;
	int iSelSpeed = 0;

	CWaitCursor wait;

	try
	{
		CWaitCursor wait;

		m_bIdle = FALSE;

		//lock medium
		hr = m_pWriter->LockDevice(m_strDriveLetter);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		UpdateState();

		m_dwBlocksWritten = 0;

		//set volume name
		WCHAR szVolumeName[200];
		::GetWindowText(GetDlgItem(IDC_DISC_NAME_EDIT), szVolumeName, _countof(szVolumeName));

		if (lstrlen(szVolumeName) > 0)
		{
			hr = m_pWriter->SetVolumeName(m_strDriveLetter, szVolumeName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		//set write speed
		CVsiWindow wndWriteSpeed = GetDlgItem(IDC_WRITE_SPEED_COMBO);
		iSelSpeedItemIndex = (int)wndWriteSpeed.SendMessage(CB_GETCURSEL, 0, 0);
		_ASSERT(CB_ERR != iSelSpeedItemIndex);

		iSelSpeed = (int)wndWriteSpeed.SendMessage(CB_GETITEMDATA, (WPARAM)iSelSpeedItemIndex, 0);
		_ASSERT(CB_ERR != iSelSpeed);

		hr = m_pWriter->SetWriteSpeed(m_strDriveLetter, (DWORD)iSelSpeed);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		m_uiTimerProgress = SetTimer(VSI_TIMER_ID_WRITE_PROGRESS, VSI_TIMER_ID_WRITE_PROGRESS_ELAPSE);
		VSI_VERIFY((m_uiTimerProgress != 0), VSI_LOG_ERROR, NULL);

		VsiProgressDialogBox(
			NULL, (LPCWSTR)MAKEINTRESOURCE(VSI_PROGRESS_TEMPLATE_MPC),
			m_hWnd, WriteCallback, this);
	}
	VSI_CATCH_(_err)
	{
		hr = _err;

		m_bIdle = TRUE;
	}

	return 0;
}


LRESULT CVsiCdDvdWriterView::OnBnClickedErase(WORD, WORD, HWND, BOOL&)
{
	HRESULT hr = S_OK;

	CWaitCursor wait;

	try
	{
		m_bIdle = FALSE;
		int iRet = VsiMessageBox(
			GetTopLevelParent(),
			L"This action will erase all data on the disc." \
				L"\r\nWould You like to continue?",
			L"Erase Confirmation", MB_YESNO | MB_APPLMODAL | MB_ICONQUESTION);

		if (IDYES != iRet)
			return 0;

		//lock medium
		hr = m_pWriter->LockDevice(m_strDriveLetter);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		UpdateState();

		m_uiTimerProgress = SetTimer(VSI_TIMER_ID_WRITE_PROGRESS, VSI_TIMER_ID_WRITE_PROGRESS_ELAPSE);
		VSI_VERIFY((m_uiTimerProgress != 0), VSI_LOG_ERROR, NULL);

		VsiProgressDialogBox(
			NULL, (LPCWSTR)MAKEINTRESOURCE(VSI_PROGRESS_TEMPLATE_M),
			m_hWnd, EraseCallback, this);
	}
	VSI_CATCH_(_err)
	{
		hr = _err;

		m_bIdle = TRUE;
	}

	return 0;
}

LRESULT CVsiCdDvdWriterView::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	::PostMessage(m_hWnd, WM_VSI_START_WRITE, 0, 0);

	return 0;
}

LRESULT CVsiCdDvdWriterView::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	m_pCmdMgr->InvokeCommand(ID_CMD_CLOSE_BURN_DISC, NULL);

	return 0;
}
/*
LRESULT CVsiCdDvdWriterView::OnBnClickedFinalize(WORD, WORD, HWND, BOOL&)
{
	m_bAddDataLater = !m_bAddDataLater;

	return 0;
}
*/
LRESULT CVsiCdDvdWriterView::OnBnClickedVerify(WORD, WORD, HWND, BOOL&)
{
	m_bVerifyData = !m_bVerifyData;

	return 0;
}

BOOL CVsiCdDvdWriterView::ProcessNotification(WPARAM wParam, LPARAM lParam)
{
	BOOL bRet = TRUE;

	if (SHCNE_MEDIAREMOVED == lParam ||
		SHCNE_MEDIAINSERTED == lParam)
	{
		SHNOTIFYSTRUCT *pshn = (SHNOTIFYSTRUCT*)wParam;

		//check whether change affects selected device
		WCHAR szPath[MAX_PATH];
		SHGetPathFromIDList(pshn->lpidl1, szPath);

		if (NULL == wcsstr(szPath, m_strDriveLetter))
			return bRet;

		if (SHCNE_MEDIAREMOVED == lParam)
			m_bMediaRemoved = TRUE;
		else
			m_bMediaRemoved = FALSE;

		if (SHCNE_MEDIAINSERTED == lParam)
			m_bIdle = TRUE;

		UpdateUI();
	}
	//check whether selected drive is suddenly removed
	else if (SHCNE_DRIVEREMOVED == lParam ||
			 SHCNE_DRIVEADD == lParam)
	{
		SHNOTIFYSTRUCT *pshn = (SHNOTIFYSTRUCT*)wParam;

		//check whether change affects selected device
		WCHAR szPath[MAX_PATH];
		SHGetPathFromIDList(pshn->lpidl1, szPath);

		if (NULL == wcsstr(szPath, m_strDriveLetter))
			return bRet;

		if (SHCNE_DRIVEREMOVED == lParam)
			m_bDeviceRemoved = TRUE;
		else
			m_bDeviceRemoved = FALSE;

		UpdateUI();
	}

	return bRet;
}

void CVsiCdDvdWriterView::UpdateUI()
{
	HRESULT hr = S_OK;
	VSI_WRITE_DEVICE_INFO DeviceInfo;
	VSI_WRITE_MEDIA_INFO MediaInfo;

	ZeroMemory(&DeviceInfo, sizeof(VSI_WRITE_DEVICE_INFO));
	ZeroMemory(&MediaInfo, sizeof(VSI_WRITE_MEDIA_INFO));

	try
	{
		CVsiWindow wndDeviceName(GetDlgItem(IDC_DEVICE_NAME_TEXT));

		if (m_bDeviceRemoved)
		{
			wndDeviceName.SetWindowText(L"No Device");
		}
		else
		{
			hr = m_pWriter->GetDeviceInfo(m_strDriveLetter, &DeviceInfo);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CString strDeviceName(DeviceInfo.szDeviceName);
			strDeviceName.Trim();
			wndDeviceName.SetWindowText(strDeviceName);
		}

		m_bCanFinalizeMedia = FALSE;
		m_bHasData = FALSE;
		m_bErasable = FALSE;
		m_bErased = FALSE;

		CVsiWindow wndMediaType(GetDlgItem(IDC_MEDIA_TYPE_TEXT));

		if (m_bDeviceRemoved)
		{
			wndMediaType.SetWindowText(L"");
		}
		else if (m_bMediaRemoved)
		{
			wndMediaType.SetWindowText(L"No Media");
		}
		else
		{
			hr = m_pWriter->GetMediaInfo(m_strDriveLetter, &MediaInfo);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			wndMediaType.SetWindowText(MediaInfo.szMediaTypeName);

			m_bHasData = MediaInfo.bHasData;
			m_bErasable = MediaInfo.bErasable;
			m_bErased = MediaInfo.bErased;

			if ( (0 == lstrcmp(L"CD-R", MediaInfo.szMediaTypeName)) ||
				 (0 == lstrcmp(L"DVD-R", MediaInfo.szMediaTypeName)) ||
				 (0 == lstrcmp(L"DVD+R", MediaInfo.szMediaTypeName)) )
				m_bCanFinalizeMedia = TRUE;

			//plugged in open tray
			if (0 == MediaInfo.iTotalSizeBytes)
			{
				if ( (lstrlen(MediaInfo.szMediaTypeName) == 0) ||
					 (NULL != wcsstr(MediaInfo.szMediaTypeName, L"Unknown")) )
				{
					m_bMediaRemoved = TRUE;
					wndMediaType.SetWindowText(L"No Media");
				}
			}
		}

		CVsiWindow wndMediaSize = GetDlgItem(IDC_DISC_SIZE_TEXT);
		if (m_bDeviceRemoved || m_bMediaRemoved)
		{
			wndMediaSize.SetWindowText(L"");
		}
		else
		{
			WCHAR szMediaSize[64];
			WCHAR szMediaSizeText[64];
			swprintf_s(szMediaSize, _countof(szMediaSize), L"%.2f", MediaInfo.dbTotalSizeMB);
			VsiGetDoubleFormat(szMediaSize, szMediaSize, _countof(szMediaSize), 2);
			swprintf_s(szMediaSizeText, _countof(szMediaSizeText), L"%s MB", szMediaSize);
			wndMediaSize.SetWindowText(szMediaSizeText);

			if (m_dbRequiredSizeMB >= MediaInfo.dbTotalSizeMB)
				m_bSpaceAvailable = FALSE;
			else
				m_bSpaceAvailable = TRUE;
		}

		CVsiWindow wndDiscName = GetDlgItem(IDC_DISC_NAME_EDIT);
		if (m_bDeviceRemoved || m_bMediaRemoved || !m_bSpaceAvailable)
		{
			wndDiscName.SetWindowText(L"");
		}
		else
		{
			wndDiscName.SetWindowText(MediaInfo.szVolumeName);
		}

		UpdateState();

		//writing speed options
		CString strItemDisplay;
		int iIndex;
		CVsiWindow wndWriteSpeed = GetDlgItem(IDC_WRITE_SPEED_COMBO);
		wndWriteSpeed.SendMessage(CB_RESETCONTENT);

		if (!m_bDeviceRemoved && !m_bMediaRemoved)
		{
			WORD wCurrentWriteSpeed = MediaInfo.wCurrentWriteSpeed;

			WORD wWriteSpeed = 1;
			while (wWriteSpeed <= MediaInfo.wMaxWriteSpeed)
			{
				strItemDisplay.Format(L"%ix", wWriteSpeed);
				iIndex = (int)wndWriteSpeed.SendMessage(
					CB_ADDSTRING, 0, (LPARAM)(LPCTSTR)strItemDisplay);

				if (CB_ERR != iIndex)
				{
					wndWriteSpeed.SendMessage(
						CB_SETITEMDATA, iIndex, (LPARAM)wWriteSpeed);
					_ASSERT(CB_ERR != iIndex);
				}

				//insert current write speed if not in our list
				if ( (wCurrentWriteSpeed > wWriteSpeed) &&  (wCurrentWriteSpeed < (wWriteSpeed * 2)) )
				{
					strItemDisplay.Format(L"%ix", wCurrentWriteSpeed);
					iIndex = (int)wndWriteSpeed.SendMessage(
						CB_ADDSTRING, 0, (LPARAM)(LPCTSTR)strItemDisplay);

					if (CB_ERR != iIndex)
					{
						wndWriteSpeed.SendMessage(
							CB_SETITEMDATA, iIndex, (LPARAM)wCurrentWriteSpeed);
						_ASSERT(CB_ERR != iIndex);
					}

				}

				if(wCurrentWriteSpeed != MediaInfo.wMaxWriteSpeed && wWriteSpeed != MediaInfo.wMaxWriteSpeed && wWriteSpeed * 2 > MediaInfo.wMaxWriteSpeed)
				{
					wWriteSpeed = MediaInfo.wMaxWriteSpeed;
				}
				else
				{
					wWriteSpeed *= 2;
				}
			}

			// select current write speed
			strItemDisplay.Format(L"%ix", wCurrentWriteSpeed);
			if (CB_ERR != iIndex)
			{
				iIndex = (int)wndWriteSpeed.SendMessage(
					CB_FINDSTRINGEXACT, 0, (LPARAM)(LPCTSTR)strItemDisplay);
				if (CB_ERR != iIndex)
				{
					wndWriteSpeed.SendMessage(CB_SETCURSEL, iIndex, 0);
				}
			}
		}


	}
	VSI_CATCH_(_err)
	{
		hr = m_pWriter->GetDeviceInfo(m_strDriveLetter, &DeviceInfo);
		if (FAILED(hr))
		{
			m_bDeviceRemoved = TRUE;
			GetDlgItem(IDC_DEVICE_NAME_TEXT).SetWindowText(L"No Device");
			UpdateState();
			return;
		}

		hr = m_pWriter->GetMediaInfo(m_strDriveLetter, &MediaInfo);
		if (S_OK != hr)
		{
			m_bMediaRemoved = TRUE;
			GetDlgItem(IDC_MEDIA_TYPE_TEXT).SetWindowText(L"No Media");
			UpdateState();
			return;
		}
	}
}

/*********#*********#*********#*********#*********#*********#*********#*********
**
**	Function: WriteCallback
**
**	Description:
**		This is the callback function used by the progress dialog to iniate
**		the work it is reporting the progress of.
**
*********#*********#*********#*********#*********#*********#*********#*********/
DWORD CALLBACK CVsiCdDvdWriterView::WriteCallback(HWND hWnd, UINT uiState, LPVOID pData)
{
	CVsiCdDvdWriterView *pWriter = (CVsiCdDvdWriterView*)pData;

	if (VSI_PROGRESS_STATE_INIT == uiState)
	{
		// Cache progress dialog handle
		pWriter->m_hwndProgress = hWnd;
		pWriter->m_uiCountdown = 0;

		::SetWindowText(hWnd, VSI_CD_DVD_WRITER_TITLE);
		SendMessage(hWnd, WM_VSI_PROGRESS_SETMSG, 0, (LPARAM)L"Writing study data...");

		HRESULT hr = S_OK;
		DWORD dwBlocksToWrite = 0;
		DWORD dwFlags = 0;

		try
		{
//			if (BST_UNCHECKED == pWriter->IsDlgButtonChecked(IDC_FINALIZE_CHECK))
			//always finalize
			dwFlags |= VSI_WRITE_OPTION_FINALIZE;

			if (BST_CHECKED == pWriter->IsDlgButtonChecked(IDC_VERIFY_CHECK))
				dwFlags |= VSI_WRITE_OPTION_VERIFY_DATA;

			hr = pWriter->m_pWriter->Initialize(
				pWriter->m_pXmlDoc, 
				dwFlags,
				CString(MAKEINTRESOURCE(IDS_STUDY_FILE_NAME)),
				CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME)));
			VSI_CHECK_HR(hr, VSI_LOG_WARNING, NULL);

			hr = pWriter->m_pWriter->GetImageSize(&dwBlocksToWrite);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			pWriter->m_dwBlocksToWrite = dwBlocksToWrite;

#ifdef _DEBUG
			ATLTRACE(L"m_dwBlocksToWrite=[%i]\n", pWriter->m_dwBlocksToWrite);
#endif

			hr = pWriter->m_pWriter->Write(pWriter->m_strDriveLetter);
			VSI_CHECK_HR(hr, VSI_LOG_WARNING, NULL);
		}
		VSI_CATCH(hr);

		if(S_OK != hr)
		{
			SendMessage(pWriter->m_hWnd, WM_VSI_WRITE_ABORTED, 0, 0);
			return 0;
		}

		// For really small studies...
		if (dwBlocksToWrite < 1)
			dwBlocksToWrite = 1;

		// Set progress bar range to the total size of all of the studies to
		// be copied.
		SendMessage(hWnd, WM_VSI_PROGRESS_SETMAX, 100, 0);
	}
	else if (VSI_PROGRESS_STATE_CANCEL == uiState)
	{
		SendMessage(pWriter->m_hWnd, WM_VSI_WRITE_CANCELLED, 0, 0);
	}
	else if (VSI_PROGRESS_STATE_END == uiState)
	{
		// Nothing
	}

	return 0;
}

/*********#*********#*********#*********#*********#*********#*********#*********
**
**	Function: EraseCallback
**
*********#*********#*********#*********#*********#*********#*********#*********/
DWORD CALLBACK CVsiCdDvdWriterView::EraseCallback(HWND hWnd, UINT uiState, LPVOID pData)
{
	CVsiCdDvdWriterView *pWriter = (CVsiCdDvdWriterView*)pData;

	if (VSI_PROGRESS_STATE_INIT == uiState)
	{
		pWriter->m_bErasing = TRUE;
		pWriter->m_uiCountdown = 0;

		// Cache progress dialog handle
		pWriter->m_hwndProgress = hWnd;

		::SetWindowText(hWnd, VSI_CD_DVD_WRITER_TITLE);
		::SendMessage(hWnd, WM_VSI_PROGRESS_SETMSG, 0, (LPARAM)L"Erasing data");

		HRESULT hr = S_OK;

		try
		{
			hr = pWriter->m_pWriter->Erase(pWriter->m_strDriveLetter);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		VSI_CATCH(hr);
	}
	else if (VSI_PROGRESS_STATE_END == uiState)
	{
		// Nothing
		pWriter->m_bErasing = FALSE;
	}

	return 0;
}

DWORD CALLBACK CVsiCdDvdWriterView::CancelProgressCallback(int iID, LPVOID pData)
{
	CVsiCdDvdWriterView *pWriter = (CVsiCdDvdWriterView*)pData;

	if (IDYES == iID)
	{
		::PostMessage(pWriter->m_hWnd, WM_VSI_PROGRESS_WND_CLOSED, 0, 0);

		pWriter->m_bCancelWrite = TRUE;
		pWriter->m_hCancelPromptWnd = NULL;

		// Updates progress text
		if (::IsWindow(pWriter->m_hwndProgress))
		{
			CWindow wndProgress(pWriter->m_hwndProgress);
			wndProgress.SendMessage(WM_VSI_PROGRESS_SETMSG, 0, (LPARAM)L"Cancelling operation");
		}

		pWriter->m_pWriter->AbortWrite();
	}
	else if (IDNO == iID)
	{
		// Reenable Cancel button on the progress dialog
		if (::IsWindow(pWriter->m_hwndProgress))
		{
			CWindow wndProgress(pWriter->m_hwndProgress);
			wndProgress.SendMessage(WM_VSI_PROGRESS_SETCANCEL, (WPARAM)1, 0);
		}
	}

	return 0;
}

BOOL CVsiCdDvdWriterView::IsDeviceIdle()
{
	HRESULT hr = S_OK;
	DWORD dwStatus = 0;
	BOOL bIdle = TRUE; // idle by default

	try
	{
		hr = m_pWriter->GetStatus(m_strDriveLetter, &dwStatus);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (dwStatus > VSI_WRITE_IDLE)
			bIdle = FALSE;
	}
	VSI_CATCH(hr);

	return bIdle;
}

void CVsiCdDvdWriterView::UpdateState()
{
	if (m_bDeviceRemoved || m_bMediaRemoved || !m_bSpaceAvailable)
	{
		const UINT uIDs[] = {
			IDOK,
			IDC_ERASE_DISC_BUTTON,
			IDC_DISC_NAME_EDIT,
//			IDC_FINALIZE_CHECK,
			IDC_VERIFY_CHECK,
			IDC_WRITE_SPEED_COMBO
		};

		for (int i = 0; i < _countof(uIDs); i++)
		{
			GetDlgItem(uIDs[i]).EnableWindow(FALSE);
		}
	}
	else
	{
		BOOL bHasToBeErased = m_bHasData && m_bErasable && !m_bErased;
		BOOL bDirtyWritable = m_bHasData && !m_bErasable;

		GetDlgItem(IDOK).EnableWindow(m_bIdle && m_bSpaceAvailable && !bHasToBeErased && !bDirtyWritable);
		GetDlgItem(IDC_ERASE_DISC_BUTTON).EnableWindow(m_bIdle && m_bErasable && !m_bErased);
		GetDlgItem(IDC_DISC_NAME_EDIT).EnableWindow(m_bIdle && !bHasToBeErased && m_bSpaceAvailable && !bDirtyWritable);
//		GetDlgItem(IDC_FINALIZE_CHECK).EnableWindow(m_bIdle && m_bCanFinalizeMedia && m_bSpaceAvailable);

		const UINT uIDs[] = {
			IDC_VERIFY_CHECK,
			IDC_WRITE_SPEED_COMBO
		};

		for (int i = 0; i < _countof(uIDs); i++)
		{
			GetDlgItem(uIDs[i]).EnableWindow(m_bIdle && m_bSpaceAvailable && !bHasToBeErased && !bDirtyWritable);
		}
	}
}

