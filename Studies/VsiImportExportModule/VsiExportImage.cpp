/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportImage.cpp
**
**	Description:
**		Implementation of CVsiExportImage
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiRes.h>
#include <shlguid.h>
#include <ATLComTime.h>
#include <VsiSaxUtils.h>
#include <VsiWTL.h>
#include <VsiGlobalDef.h>
#include <VsiServiceProvider.h>
#include <VsiServiceKey.h>
#include <VsiCommUtl.h>
#include <VsiCommCtl.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterMode.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiAppControllerModule.h>
#include <VsiMeasurementModule.h>
#include <VsiAnalysisResultsModule.h>
#include <VsiAnalRsltXmlTags.h>
#include <VsiImportExportXmlTags.h>
#include <VsiDiagnosticsModule.h>
#include <VsiLicenseUtils.h>
#include <VsiModeUtils.h>
#include <Vsi3dModule.h>
#include "VsiReportExportStatus.h"
#include "VsiExportImageCineDlg.h"
#include "VsiExportImageFrameDlg.h"
#include "VsiExportImageDicomDlg.h"
#include "VsiExportImageReportDlg.h"
#include "VsiExportImageGraphDlg.h"
#include "VsiExportImageTableDlg.h"
#include "VsiExportImagePhysioDlg.h"
#include "VsiExportImage.h"

#define VSI_AUTOCOMPLETE_NAME	L"VsiHistoryExportImage.xml"

CVsiExportImage::CVsiExportImage() :
	m_dwFlags(0),
	m_pSubPanelCine(new CVsiExportImageCineDlg(this)),
	m_pSubPanelFrame(new CVsiExportImageFrameDlg(this)),
	m_pSubPanelDicom(new CVsiExportImageDicomDlg(this)),
	m_pSubPanelReport(new CVsiExportImageReportDlg(this)),
	m_pSubPanelGraph(new CVsiExportImageGraphDlg(this)),
	m_pSubPanelTable(new CVsiExportImageTableDlg(this)),
	m_pSubPanelPhysio(new CVsiExportImagePhysioDlg(this)),
	m_dwExportTypeMain(VSI_EXPORT_IMAGE_FLAG_CINE),
	m_bHideMsmnts(FALSE),
	m_pExportProgress(NULL),
	m_iOffset(0),
	m_i3Ds(0)
{
}

CVsiExportImage::~CVsiExportImage()
{
	_ASSERT(NULL == m_pApp);
	_ASSERT(NULL == m_pCmdMgr);
	_ASSERT(NULL == m_pPdm);
	_ASSERT(NULL == m_pModeMgr);
	_ASSERT(NULL == m_pExportProgress);
}

STDMETHODIMP CVsiExportImage::Activate(IVsiApp *pApp, HWND hwndParent)
{
	HRESULT hr = S_OK;

	try
	{
		m_pApp = pApp;
		VSI_CHECK_INTERFACE(m_pApp, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Get Command Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CMD_MGR,
			__uuidof(IVsiCommandManager),
			(IUnknown**)&m_pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Command Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CURSOR_MANAGER,
			__uuidof(IVsiCursorManager),
			(IUnknown**)&m_pCursorMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get PDM
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Cache Mode Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_MODE_MANAGER,
			__uuidof(IVsiModeManager),
			(IUnknown**)&m_pModeMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		HWND hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiExportImage::Deactivate()
{
	HRESULT hr = S_FALSE;

	if (IsWindow())
	{
		BOOL bRet = DestroyWindow();
		_ASSERT(bRet);

		hr = S_OK;
	}

	m_pModeMgr.Release();
	m_pCmdMgr.Release();
	m_pPdm.Release();
	m_pApp.Release();

	return hr;
}

STDMETHODIMP CVsiExportImage::GetWindow(HWND *phWnd)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(phWnd, VSI_LOG_ERROR, NULL);

		*phWnd = m_hWnd;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiExportImage::GetAccelerator(HACCEL *phAccel)
{
	UNREFERENCED_PARAMETER(phAccel);
	return E_NOTIMPL;
}

STDMETHODIMP CVsiExportImage::GetIsBusy(
	DWORD dwStateCurrent,
	DWORD dwState,
	BOOL bTryRelax,
	BOOL *pbBusy)
{
	UNREFERENCED_PARAMETER(dwStateCurrent);
	UNREFERENCED_PARAMETER(dwState);
	UNREFERENCED_PARAMETER(bTryRelax);

	switch (dwState)
	{
	case VSI_APP_STATE_EXPORT_IMAGE_CLOSE:
	case VSI_APP_STATE_DO_EXPORT_IMAGE:
		*pbBusy = FALSE;
		break;

	default:
		*pbBusy = !GetTopLevelParent().IsWindowEnabled();
		break;
	}

	return S_OK;
}

STDMETHODIMP CVsiExportImage::PreTranslateMessage(MSG *pMsg, BOOL *pbHandled)
{
	*pbHandled = IsDialogMessage(pMsg);

	return S_OK;
}

STDMETHODIMP CVsiExportImage::Initialize(DWORD dwFlags)
{
	m_dwFlags = dwFlags;

	return S_OK;
}

STDMETHODIMP CVsiExportImage::AddImage(IUnknown *pUnkImage, IUnknown *pUnkMode)
{
	HRESULT hr = S_OK;

	try
	{
		// pUnkImage is optional
		// pUnkMode is optional

		VSI_VERIFY((pUnkImage != NULL || pUnkMode != NULL), VSI_LOG_ERROR, NULL);

		CVsiImage item(pUnkImage, pUnkMode);
		m_listImage.push_back(item);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiExportImage::SetTable(HWND hwndTable)
{
	m_ExportTable.SetTable(hwndTable);

	return S_OK;
}

STDMETHODIMP CVsiExportImage::SetReport(IUnknown *pUnkXmlDoc)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pUnkXmlDoc, VSI_LOG_ERROR, NULL);

		m_pXmlReportDoc = pUnkXmlDoc;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiExportImage::SetAnalysisExport(IUnknown *pAnalysisExport)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pAnalysisExport, VSI_LOG_ERROR, NULL);

		m_pAnalysisExport = pAnalysisExport;
	}
	VSI_CATCH(hr);

	return hr;
}

BOOL CVsiExportImage::GetAvailablePath(LPWSTR pszBuffer, int iBufferLength)
{
	return GetSelectedPath(pszBuffer, iBufferLength);
}

int CVsiExportImage::CheckForOverwrite(LPCWSTR pszPath, LPCWSTR pszFile)
{
	int iRet = 0;
	HRESULT hr = S_OK;

	try
	{
		switch (m_dwExportTypeMain)
		{
		case VSI_EXPORT_IMAGE_FLAG_TABLE:
			{
				WCHAR szPath[MAX_PATH];
				swprintf_s(szPath, _countof(szPath), L"%s\\%s.%s",
					pszPath, pszFile, GetFileExtension());

				if (0 == _waccess(szPath, 0))
				{
					iRet = 1;
				}
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_GRAPH:
			{
				if (m_pSubPanelGraph->IsDlgButtonChecked(IDC_IE_EXPORT_CSV))
				{
					WCHAR szPath[MAX_PATH];
					swprintf_s(szPath, _countof(szPath), L"%s\\%s.csv",
						pszPath, pszFile);

					if (0 == _waccess(szPath, 0))
					{
						iRet = 1;
					}
				}

				if (m_pSubPanelGraph->IsDlgButtonChecked(IDC_IE_EXPORT_TIFF))
				{
					WCHAR szPath[MAX_PATH];
					swprintf_s(szPath, _countof(szPath), L"%s\\%s.tiff",
						pszPath, pszFile);

					if (0 == _waccess(szPath, 0))
					{
						iRet = 1;
					}
				}

				if (m_pSubPanelGraph->IsDlgButtonChecked(IDC_IE_EXPORT_BMP))
				{
					WCHAR szPath[MAX_PATH];
					swprintf_s(szPath, _countof(szPath), L"%s\\%s.bmp",
						pszPath, pszFile);

					if (0 == _waccess(szPath, 0))
					{
						iRet = 1;
					}
				}
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_CINE:
				{
					// Cine loop as one file
					WCHAR szPath[MAX_PATH];
					swprintf_s(szPath, _countof(szPath), L"%s\\%s.%s",
						pszPath, pszFile, GetFileExtension());

					if (0 == _waccess(szPath, 0))
					{
						iRet = 1;
					}
				}
			break;

		default:
			{
				CWindow wndExportType(GetTypeControl());

				int iIndex = (int)wndExportType.SendMessage(CB_GETCURSEL, NULL, NULL);
				DWORD dwType = (DWORD)wndExportType.SendMessage(CB_GETITEMDATA, iIndex, NULL);
				if (dwType == VSI_IMAGE_TIFF_LOOP)
				{
					WCHAR szPathXml[MAX_PATH];
					WCHAR szPathBMode[MAX_PATH];
					WCHAR szPathColor[MAX_PATH];

					CVsiListImageIter iter;
					for (iter = m_listImage.begin(); iter != m_listImage.end(); ++iter)
					{
						CVsiImage &item = *iter;

						BOOL bColorFrame(FALSE);
						if (NULL != item.m_pMode)
						{
							CComPtr<IVsiModeData> pModeData;
							hr = item.m_pMode->GetModeData(&pModeData);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							CComQIPtr<IVsi3dModeData> p3dData(pModeData);
							VSI_CHECK_INTERFACE(p3dData, VSI_LOG_ERROR, NULL);

							DWORD dw3dType;
							p3dData->Get3dModeType(&dw3dType);
							if (dw3dType != VSI_3DMODE_REGULAR)
								bColorFrame = TRUE;
						}
						else if (NULL != item.m_pImage)
						{
							CComVariant vMode;
							hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

							CComVariant vModeId;
							hr = m_pModeMgr->GetModePropertyFromInternalName(
								V_BSTR(&vMode),
								L"id",
								&vModeId);
							VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

							if (V_UI4(&vModeId) != VSI_SETTINGS_MODE_3D)
								bColorFrame = TRUE;
						}

						swprintf_s(szPathXml, _countof(szPathXml), L"%s\\%s.%s", pszPath, pszFile, L"xml");
						HANDLE hFind = INVALID_HANDLE_VALUE;
						WIN32_FIND_DATA fd;
						hFind = FindFirstFile(szPathXml, &fd);

						if (INVALID_HANDLE_VALUE != hFind)
						{
							iRet = 1;
						}
						else
						{
							FindClose(hFind);
							swprintf_s(szPathBMode, _countof(szPathBMode), L"%s\\%s_C1_*.%s", pszPath, pszFile, L"tif");
							hFind = FindFirstFile(szPathBMode, &fd);

							if (INVALID_HANDLE_VALUE != hFind)
							{
								iRet = 1;
							}
							else if (bColorFrame)
							{
								FindClose(hFind);
								swprintf_s(szPathColor, _countof(szPathColor), L"%s\\%s_C2_*.%s", pszPath, pszFile, L"tif");
								hFind = FindFirstFile(szPathColor, &fd);

								if (INVALID_HANDLE_VALUE != hFind)
								{
									iRet = 1;
								}
								else
								{
									FindClose(hFind);
									hFind = INVALID_HANDLE_VALUE;
								}
							}
						}
						if (1 == iRet)
						{
							FindClose(hFind);
							hFind = INVALID_HANDLE_VALUE;
							break;
						}
					}
				}
				else
				{
					// Check for overwrite for single frames
					if (1 == m_listImage.size() ||
						NULL != m_pXmlReportDoc ||
						(VSI_EXPORT_IMAGE_FLAG_REPORT & m_dwFlags))
					{
						WCHAR szPath[MAX_PATH];
						swprintf_s(szPath, _countof(szPath), L"%s\\%s.%s",
							pszPath, pszFile, GetFileExtension());

						if (0 == _waccess(szPath, 0))
						{
							iRet = 1;
						}
					}
				}
			}
			break;
		}
	}
	VSI_CATCH(hr);

	return iRet;
}

int CVsiExportImage::GetNumberOfSelectedItems()
{
	int iNum = (int)m_listImage.size();

	if (NULL != m_pXmlReportDoc || VSI_EXPORT_IMAGE_FLAG_REPORT & m_dwFlags)
		iNum = 1;

	return iNum;
}

LRESULT CVsiExportImage::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		m_bHideMsmnts = VsiGetBooleanValue<BOOL>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_MSMNT_EXPORT_HIDE, m_pPdm);

		int iFrames(0);
		int iLoops(0);
		int iAll = GetTypeCount(&iFrames, &iLoops, &m_i3Ds);

		m_wndName.SubclassWindow(GetDlgItem(IDC_IE_SAVEAS));
		m_wndName.SetAutoActivate(false);

		// Auto Complete - Keep going on error
		try
		{
			hr = m_pAutoCompleteSource.CoCreateInstance(__uuidof(VsiAutoCompleteSource));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			WCHAR pszPath[MAX_PATH];
			BOOL bRet = VsiGetApplicationDataDirectory(
				AtlFindStringResourceInstance(IDS_VSI_PRODUCT_NAME),
				pszPath);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, NULL);

			PathAppend(pszPath, VSI_AUTOCOMPLETE_NAME);

			hr = m_pAutoCompleteSource->LoadSettings(pszPath);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComPtr<IEnumString> pEnumString;
			hr = m_pAutoCompleteSource->GetSource(&pEnumString);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pAutoComplete.CoCreateInstance(CLSID_AutoComplete);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pAutoComplete->Init(GetDlgItem(IDC_IE_SAVEAS), pEnumString, NULL, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComQIPtr<IAutoComplete2> pAutoComplete2(m_pAutoComplete);
			if (pAutoComplete2 != NULL)
			{
				hr = pAutoComplete2->SetOptions(ACO_AUTOSUGGEST | ACO_UPDOWNKEYDROPSLIST);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}
		VSI_CATCH(hr);

		// Set font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
		VsiThemeRecurSetFont(m_hWnd, hFont);

		// Setup sub panels
		{
			CRect rc;
			GetDlgItem(IDC_IE_OPTIONS).GetWindowRect(&rc);
			ScreenToClient(&rc);

			GetDlgItem(IDC_IE_PANEL_CINE).DestroyWindow();
			m_pSubPanelCine->Initialize(m_pPdm);
			m_pSubPanelCine->Create(*this);
			m_pSubPanelCine->SetDlgCtrlID(IDC_IE_PANEL_CINE);

			m_pSubPanelCine->SetWindowPos(
				GetDlgItem(IDC_IE_OPTIONS), rc.left, rc.top, 0, 0, SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOREDRAW);

			GetDlgItem(IDC_IE_PANEL_FRAME).DestroyWindow();
			m_pSubPanelFrame->Initialize(m_pPdm);
			m_pSubPanelFrame->Create(*this);
			m_pSubPanelFrame->SetDlgCtrlID(IDC_IE_PANEL_FRAME);

			m_pSubPanelFrame->SetWindowPos(
				GetDlgItem(IDC_IE_OPTIONS), rc.left, rc.top, 0, 0, SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOREDRAW);

			GetDlgItem(IDC_IE_PANEL_DICOM).DestroyWindow();
			m_pSubPanelDicom->Initialize(m_pPdm);
			m_pSubPanelDicom->Create(*this);
			m_pSubPanelDicom->SetDlgCtrlID(IDC_IE_PANEL_DICOM);

			m_pSubPanelDicom->SetWindowPos(
				GetDlgItem(IDC_IE_OPTIONS), rc.left, rc.top, 0, 0, SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOREDRAW);

			GetDlgItem(IDC_IE_PANEL_REPORT).DestroyWindow();
			m_pSubPanelReport->Initialize(m_pPdm);
			m_pSubPanelReport->Create(*this);
			m_pSubPanelReport->SetDlgCtrlID(IDC_IE_PANEL_REPORT);

			m_pSubPanelReport->SetWindowPos(
				GetDlgItem(IDC_IE_OPTIONS), rc.left, rc.top, 0, 0, SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOREDRAW);

			GetDlgItem(IDC_IE_PANEL_GRAPH).DestroyWindow();
			m_pSubPanelGraph->Initialize(m_pPdm);
			m_pSubPanelGraph->Create(*this);
			m_pSubPanelGraph->SetDlgCtrlID(IDC_IE_PANEL_GRAPH);

			m_pSubPanelGraph->SetWindowPos(
				GetDlgItem(IDC_IE_OPTIONS), rc.left, rc.top, 0, 0, SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOREDRAW);

			GetDlgItem(IDC_IE_PANEL_TABLE).DestroyWindow();
			m_pSubPanelTable->Initialize(m_pPdm);
			m_pSubPanelTable->Create(*this);
			m_pSubPanelTable->SetDlgCtrlID(IDC_IE_PANEL_TABLE);

			m_pSubPanelTable->SetWindowPos(
				GetDlgItem(IDC_IE_OPTIONS), rc.left, rc.top, 0, 0, SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOREDRAW);

			GetDlgItem(IDC_IE_PANEL_PHYSIO).DestroyWindow();
			m_pSubPanelPhysio->Initialize(m_pPdm);
			m_pSubPanelPhysio->Create(*this);
			m_pSubPanelPhysio->SetDlgCtrlID(IDC_IE_PANEL_PHYSIO);

			m_pSubPanelPhysio->SetWindowPos(
				GetDlgItem(IDC_IE_OPTIONS), rc.left, rc.top, 0, 0, SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOREDRAW);
		}

		// Layout
		CComPtr<IVsiWindowLayout> pLayout;
		hr = pLayout.CoCreateInstance(__uuidof(VsiWindowLayout));
		if (SUCCEEDED(hr))
		{
			hr = pLayout->Initialize(m_hWnd, VSI_WL_FLAG_AUTO_RELEASE);
			if (SUCCEEDED(hr))
			{
				pLayout->AddControl(0, IDC_IE_NEW_FOLDER, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDOK, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDCANCEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_PATH, VSI_WL_SIZE_X);
				pLayout->AddControl(0, IDC_IE_FOLDER_TREE, VSI_WL_SIZE_XY);
				pLayout->AddControl(0, IDC_IE_TYPE_MAIN, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_CINE_LOOP, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_IMAGE, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_DICOM, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_TABLE, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_GRAPH, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_REPORT, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_PHYSIO, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_SELECTED, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_REQUIRED_LABEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_REQUIRED, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_AVAILABLE_LABEL, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_AVAILABLE, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_OPTIONS, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_PANEL_CINE, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_PANEL_FRAME, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_PANEL_DICOM, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_PANEL_REPORT, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_PANEL_GRAPH, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_PANEL_TABLE, VSI_WL_MOVE_X);
				pLayout->AddControl(0, IDC_IE_PANEL_PHYSIO, VSI_WL_MOVE_X);

				pLayout.Detach();
			}
		}

		if (0 == iAll)
		{
			if (NULL != m_pXmlReportDoc || VSI_EXPORT_IMAGE_FLAG_REPORT & m_dwFlags)
			{
				m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_REPORT;

				GetDlgItem(IDC_IE_CINE_LOOP).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_IMAGE).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_DICOM).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_GRAPH).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_PHYSIO).EnableWindow(FALSE);

				CheckDlgButton(IDC_IE_REPORT, BST_CHECKED);
			}
			else if (VSI_EXPORT_IMAGE_FLAG_GRAPH & m_dwFlags)
			{
				m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_GRAPH;

				GetDlgItem(IDC_IE_CINE_LOOP).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_IMAGE).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_DICOM).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_REPORT).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_TABLE).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_TABLE).ShowWindow(SW_HIDE);
				GetDlgItem(IDC_IE_PHYSIO).EnableWindow(FALSE);

				CheckDlgButton(IDC_IE_GRAPH, BST_CHECKED);
			}
			else
			{
				m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_TABLE;

				GetDlgItem(IDC_IE_CINE_LOOP).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_IMAGE).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_DICOM).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_REPORT).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_GRAPH).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_PHYSIO).EnableWindow(FALSE);

				CheckDlgButton(IDC_IE_TABLE, BST_CHECKED);
			}
		}
		else if (IsAllLoopsOf3dModes())
		{
			m_dwExportTypeMain = VsiGetRangeValue<DWORD>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_DEFAULT_EXPORT_TYPE, m_pPdm);

			m_dwExportTypeMain &= m_dwFlags;
			if (0 == m_dwExportTypeMain)
			{
				m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_IMAGE;
			}

			if (VSI_EXPORT_IMAGE_FLAG_PHYSIO == m_dwExportTypeMain)
			{
				m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_IMAGE;
			}

			// Disable loop, DICOM and physiological export
			GetDlgItem(IDC_IE_CINE_LOOP).EnableWindow(FALSE);

			// Enable TomTec DICOM export in Engineering Mode
			if (!VsiModeUtils::GetIsEngineerMode(m_pPdm))
			{
				GetDlgItem(IDC_IE_DICOM).EnableWindow(FALSE);
			}

			GetDlgItem(IDC_IE_PHYSIO).EnableWindow(FALSE);

			if (VSI_EXPORT_IMAGE_FLAG_CINE == m_dwExportTypeMain)
			{
				// if loop, switch to image
				CheckDlgButton(IDC_IE_IMAGE, BST_CHECKED);
			}
			else if (VSI_EXPORT_IMAGE_FLAG_IMAGE == m_dwExportTypeMain)
			{
				CheckDlgButton(IDC_IE_IMAGE, BST_CHECKED);
			}
			else if (VSI_EXPORT_IMAGE_FLAG_DICOM == m_dwExportTypeMain)
			{
				// if loop, switch to image
				CheckDlgButton(IDC_IE_IMAGE, BST_CHECKED);

				// Enable TomTec DICOM export in Engineering Mode
				if (VsiModeUtils::GetIsEngineerMode(m_pPdm))
				{
					CheckDlgButton(IDC_IE_DICOM, BST_CHECKED);
				}
			}
			else if (VSI_EXPORT_IMAGE_FLAG_REPORT == m_dwExportTypeMain)
			{
				CheckDlgButton(IDC_IE_REPORT, BST_CHECKED);
			}
			else if (VSI_EXPORT_IMAGE_FLAG_TABLE == m_dwExportTypeMain)
			{
				CheckDlgButton(IDC_IE_TABLE, BST_CHECKED);
			}
			else if (VSI_EXPORT_IMAGE_FLAG_GRAPH == m_dwExportTypeMain)
			{
				CheckDlgButton(IDC_IE_GRAPH, BST_CHECKED);
			}
		}
		else
		{
			if (0 == iLoops)
			{
				GetDlgItem(IDC_IE_CINE_LOOP).EnableWindow(FALSE);
			}
			if (0 == iFrames && 0 == iLoops && 0 == m_i3Ds)  // Can export loop as frame
			{
				GetDlgItem(IDC_IE_IMAGE).EnableWindow(FALSE);
			}

			m_dwExportTypeMain = VsiGetRangeValue<DWORD>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_DEFAULT_EXPORT_TYPE, m_pPdm);

			m_dwExportTypeMain &= m_dwFlags;

			if (0 == m_dwExportTypeMain)
			{
				m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_IMAGE;
			}

			if (0 == (VSI_EXPORT_IMAGE_FLAG_PHYSIO & m_dwFlags))
			{
				GetDlgItem(IDC_IE_PHYSIO).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_PHYSIO).ShowWindow(SW_HIDE);
			}

			if (0 == (VSI_EXPORT_IMAGE_FLAG_GRAPH & m_dwFlags))
			{
				GetDlgItem(IDC_IE_GRAPH).EnableWindow(FALSE);
				GetDlgItem(IDC_IE_GRAPH).ShowWindow(SW_HIDE);
			}
			else
			{
				// Force graph always when flag is set
				m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_GRAPH;
			}

			if (VSI_EXPORT_IMAGE_FLAG_CINE == m_dwExportTypeMain)
			{
				if (0 == iLoops)
				{
					m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_IMAGE;
				}
				else
				{
					CheckDlgButton(IDC_IE_CINE_LOOP, BST_CHECKED);
				}
			}

			if (VSI_EXPORT_IMAGE_FLAG_IMAGE == m_dwExportTypeMain)
			{
				if (0 == iFrames && 0 == iLoops)
				{
					// ???
					m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_DICOM;
				}
				else
				{
					CheckDlgButton(IDC_IE_IMAGE, BST_CHECKED);
				}
			}

			if (VSI_EXPORT_IMAGE_FLAG_PHYSIO == m_dwExportTypeMain)
			{
				if (0 == iFrames && 0 == iLoops)
				{
					// ???
					m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_DICOM;
				}
				else
				{
					CheckDlgButton(IDC_IE_PHYSIO, BST_CHECKED);
				}
			}

			if (VSI_EXPORT_IMAGE_FLAG_DICOM == m_dwExportTypeMain)
			{
				CheckDlgButton(IDC_IE_DICOM, BST_CHECKED);
			}
			else if (VSI_EXPORT_IMAGE_FLAG_REPORT == m_dwExportTypeMain)
			{
				CheckDlgButton(IDC_IE_REPORT, BST_CHECKED);
			}
			else if (VSI_EXPORT_IMAGE_FLAG_TABLE == m_dwExportTypeMain)
			{
				CheckDlgButton(IDC_IE_TABLE, BST_CHECKED);
			}
			else if (VSI_EXPORT_IMAGE_FLAG_GRAPH == m_dwExportTypeMain)
			{
				CheckDlgButton(IDC_IE_GRAPH, BST_CHECKED);
			}
		}

		if (0 == (VSI_EXPORT_IMAGE_FLAG_GRAPH & m_dwFlags))
		{
			GetDlgItem(IDC_IE_GRAPH).ShowWindow(SW_HIDE);
			GetDlgItem(IDC_IE_GRAPH).EnableWindow(FALSE);
		}

		if (0 == (VSI_EXPORT_IMAGE_FLAG_TABLE & m_dwFlags))
		{
			GetDlgItem(IDC_IE_TABLE).ShowWindow(SW_HIDE);
			GetDlgItem(IDC_IE_TABLE).EnableWindow(FALSE);
		}

		hr = m_pFExplorer.CoCreateInstance(__uuidof(VsiFolderExplorer));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Sets file name
		if (m_listImage.size() == 1)
		{
			CVsiImage &item = *m_listImage.begin();

			SYSTEMTIME stDate;
			CComVariant vDate;
			CComVariant vLabel;

			if (item.m_pImage != NULL)
			{
				hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vDate);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vLabel);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
			else if (item.m_pMode != NULL)
			{
				CComPtr<IVsiModeData> pModeData;

				hr = item.m_pMode->GetModeData(&pModeData);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComQIPtr<IVsiPropertyBag> pProps(pModeData);
				VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

				hr = pProps->ReadId(VSI_PROP_MODE_DATA_ACQ_START_TIME, &vDate);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				// Set image
				CComQIPtr<IVsiPropertyBag> pModeProps(item.m_pMode);
				VSI_CHECK_INTERFACE(pModeProps, VSI_LOG_ERROR, NULL);

				CComVariant vImage;
				hr = pModeProps->ReadId(VSI_PROP_MODE_IMAGE, &vImage);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					CComQIPtr<IVsiImage> pImage(V_UNKNOWN(&vImage));
					VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

					hr = pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vLabel);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}
			else
			{
				VSI_FAIL(VSI_LOG_ERROR, NULL);
			}

			COleDateTime date(vDate);
			date.GetAsSystemTime(stDate);

			SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

			CString strName;
			if (VT_BSTR == V_VT(&vLabel) && (lstrlen(V_BSTR(&vLabel)) > 0))
			{
				// Check to see if the label has special characters
				CString strImageLabel(V_BSTR(&vLabel));
				LPCWSTR ppszSpecial[] = {
					L"/", L"\\", L":", L"*", L"?", L"\"", L"<", L">", L"|"};

				for(int i=0; i<_countof(ppszSpecial); ++i)
				{
					strImageLabel.Replace(ppszSpecial[i], L"_");
				}

				strName.Format(L"%s_%04d-%02d-%02d-%02d-%02d-%02d",
					strImageLabel, stDate.wYear, stDate.wMonth, stDate.wDay,
					stDate.wHour, stDate.wMinute, stDate.wSecond);
			}
			else
			{
				strName.Format(L"%04d-%02d-%02d-%02d-%02d-%02d",
					stDate.wYear, stDate.wMonth, stDate.wDay,
					stDate.wHour, stDate.wMinute, stDate.wSecond);
			}

			SetDlgItemText(IDC_IE_SAVEAS, strName);
		}
		else
		{
			// Date
			SYSTEMTIME stDate;
			GetSystemTime(&stDate);
			SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

			CString strName;
			strName.Format(L"%04d-%02d-%02d-%02d-%02d-%02d",
				stDate.wYear, stDate.wMonth, stDate.wDay,
				stDate.wHour, stDate.wMinute, stDate.wSecond);

			SetDlgItemText(IDC_IE_SAVEAS, strName);
		}

		// Refresh UI
		UpdateTypeMain(TRUE);

		// Refresh tree
		SetDlgItemText(IDC_IE_PATH, L"Loading...");
		PostMessage(WM_VSI_REFRESH_TREE);

		// Sets focus
		PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)m_wndName, TRUE);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiExportImage::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		if (m_pFExplorer != NULL)
		{
			CComHeapPtr<OLECHAR> pszPathName;
			hr = m_pFExplorer->GetSelectedNamespacePath(&pszPathName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_pFExplorer->Uninitialize();
			m_pFExplorer.Release();

			CVsiRange<LPCWSTR> pParamPath;
			hr = m_pPdm->GetParameter(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
				VSI_PARAMETER_SYS_PATH_EXPORT_IMAGE, &pParamPath);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			pParamPath.SetValue(pszPathName);
		}
	}
	VSI_CATCH(hr)

	return 0;
}

LRESULT CVsiExportImage::OnWindowPosChanged(UINT, WPARAM, LPARAM lp, BOOL &bHandled)
{
	WINDOWPOS *pPos = (WINDOWPOS*)lp;

	HRESULT hr = S_OK;

	try
	{
		// Updates preview list when turn visible
		if (SWP_SHOWWINDOW & pPos->flags)
		{
			hr = m_pCursorMgr->SetState(VSI_CURSOR_STATE_ON, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
	}
	VSI_CATCH(hr);
	bHandled = FALSE;

	// Act as if nothing happened
	return 1;
}

LRESULT CVsiExportImage::OnRefreshTree(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_FE_FLAG dwFlags(VSI_FE_FLAG_NONE);
		CComPtr<IFolderFilter> pFolderFilter;

		DWORD dwState = VsiGetBitfieldValue<DWORD>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_STATE, m_pPdm);
		if (0 < (VSI_SYS_STATE_SYSTEM & dwState) &&
			0 == (VSI_SYS_STATE_ENGINEER_MODE & dwState))
		{
			// For Sierra system, shows removable and network drive, do not show CD ROM
			dwFlags = VSI_FE_FLAG_SHOW_REMOVEABLE |
				VSI_FE_FLAG_SHOW_NETWORK | VSI_FE_FLAG_SHOW_NOCDROM;

			// Install custom filter
			hr = pFolderFilter.CoCreateInstance(__uuidof(VsiFolderFilter));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else
		{
			// For workstation and eng mode, shows Desktop
			dwFlags = VSI_FE_FLAG_SHOW_DESKTOP |
				VSI_FE_FLAG_EXPLORER | VSI_FE_FLAG_SHOW_NOCDROM;
		}

		hr = m_pFExplorer->Initialize(
			GetDlgItem(IDC_IE_FOLDER_TREE),
			dwFlags | VSI_FE_FLAG_IGNORE_LAST_FOLDER,
			pFolderFilter);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Set selection
		PostMessage(WM_VSI_SET_SELECTION);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiExportImage::OnSetSelection(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	CWaitCursor wait;

	try
	{
		// Set default selected folder
		CVsiRange<LPCWSTR> pParamPath;
		hr = m_pPdm->GetParameter(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_SYS_PATH_EXPORT_IMAGE, &pParamPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiParameterRange> pRange(pParamPath.m_pParam);
		VSI_CHECK_INTERFACE(pRange, VSI_LOG_ERROR, NULL);

		CComVariant vPath;
		hr = pRange->GetValue(&vPath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = m_pFExplorer->SetSelectedNamespacePath(V_BSTR(&vPath));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (S_OK != hr)
		{
			UpdateUI();
		}
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiExportImage::OnExportImage(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	BOOL bValidType = TRUE;

	try
	{
		// Progress
		if (m_listImage.size() > 1)
		{
			CString strMsg;
			strMsg.Format(
				L"Exporting Image - %d/%d",
				m_pExportProgress->m_iImage + 1,
				m_listImage.size());
			m_pExportProgress->m_wndProgress.SendMessage(
				WM_VSI_PROGRESS_SETMSG, 0, (LPARAM)(LPCWSTR)strMsg);
		}

		DWORD dwExportResInitial = VsiGetRangeValue<DWORD>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_EXPORT_RES_INITIAL, m_pPdm);
		double dbExportScaleHigh = VsiGetRangeValue<double>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_EXPORT_SCALE_HIGH, m_pPdm);
		double dbExportScaleMed = VsiGetRangeValue<double>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_EXPORT_SCALE_MED, m_pPdm);
		DWORD dwAviQuality = VsiGetDiscreteValue<DWORD>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
			VSI_PARAMETER_SYS_EXPORT_AVI_QUALITY, m_pPdm);

		CWindow wndExportType(GetTypeControl());

		int iIndex = (int)wndExportType.SendMessage(CB_GETCURSEL, NULL, NULL);
		DWORD dwType = (DWORD)wndExportType.SendMessage(CB_GETITEMDATA, iIndex, NULL);

		DWORD dwFcc(0);

		// Get FCC code / only required for AVI types
		if (VSI_IMAGE_AVI_CUSTOM == dwType)
		{
			if (iIndex >= 0 && iIndex < _countof(m_pSubPanelCine->m_pdwFccHandlerType))
			{
				dwFcc = m_pSubPanelCine->m_pdwFccHandlerType[iIndex];
			}
		}

		WCHAR szFileName[MAX_PATH];
		m_wndName.GetWindowText(szFileName, _countof(szFileName));

		LPCWSTR pszExt = GetFileExtension();

		// Test to see if the file name exists. If it does, keep adding digits
		// until we get something unique.

		CVsiImage &item = *m_pExportProgress->m_iterExport;

		// If image is NULL, we must be exporting from a mode directly so try getting it this way
		if (NULL == item.m_pImage)
		{
			CComPtr<IVsiMode> pMode = item.m_pMode;
			if (NULL != pMode)
			{
				CComPtr<IVsiApp> pApp;
				pMode->GetApp(&pApp);
				if (NULL != pApp)
				{
					CComQIPtr<IVsiServiceProvider> pServiceProvider(pApp);
					VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiSession> pSession;
					pServiceProvider->QueryService(VSI_SERVICE_SESSION_MGR, __uuidof(IVsiSession), (IUnknown**)&pSession);
					if (NULL != pSession)
					{
						hr = pSession->GetImage(VSI_SLOT_ACTIVE, &item.m_pImage);
					}
				}
			}
		}

		WCHAR szPath[MAX_PATH];
		CComVariant vModeId;
		CString strModeName;

		CComVariant vFrame(false);
		CComVariant vRf(false);
		BOOL b3D(FALSE);
		VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_NONE);

		if (NULL != item.m_pImage)
		{
			item.m_pImage->GetErrorCode(&dwErrorCode);

			if (VSI_IMAGE_ERROR_NONE == dwErrorCode)
			{
				hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_FRAME_TYPE, &vFrame);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_RF_DATA, &vRf);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComVariant vModeName;
				hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vModeName);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				CComVariant v3d;
				hr = m_pModeMgr->GetModePropertyFromInternalName(
					V_BSTR(&vModeName),
					L"3d",
					&v3d);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				if (VARIANT_FALSE != V_BOOL(&v3d))
				{
					b3D = TRUE;
				}

				strModeName = V_BSTR(&vModeName);
			}
		}
		else if (NULL != item.m_pMode)
		{
			CComQIPtr<IVsiPropertyBag> pProps(item.m_pMode);
			VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

			CComVariant vRoot;
			hr = pProps->ReadId(VSI_PROP_MODE_ROOT, &vRoot);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			BOOL bIsFrame = VsiGetBooleanValue<BOOL>(
				V_BSTR(&vRoot), VSI_PDM_GROUP_GLOBAL,
				VSI_PARAMETER_SETTINGS_SAVED_FRAME, m_pPdm);
			vFrame = (FALSE != bIsFrame);

			BOOL bRfFeature = VsiGetBooleanValue<BOOL>(
				V_BSTR(&vRoot), VSI_PDM_GROUP_MODE,
				VSI_PARAMETER_SETTINGS_RF_FEATURE, m_pPdm);
			vRf = (FALSE != bRfFeature);

			hr = pProps->ReadId(VSI_PROP_MODE_MODE_ID, &vModeId);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CComVariant v3d;
			hr = m_pModeMgr->GetModePropertyFromId(V_UI4(&vModeId),	L"3d", &v3d);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			if (VARIANT_FALSE != V_BOOL(&v3d))
			{
				b3D = TRUE;
			}

			CComVariant vModeName;
			hr = m_pModeMgr->GetModeNameFromIdNoFlags(V_I4(&vModeId), V_BSTR(&vRoot), &vModeName);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			strModeName = V_BSTR(&vModeName);
		}
		else
		{
			_ASSERT(0);
		}

		if (VSI_IMAGE_ERROR_NONE == dwErrorCode)
		{
			if (m_listImage.size() > 1)
			{
				// This type isn't valid with cine formats
				switch (dwType)
				{
				case VSI_IMAGE_AVI_CUSTOM:
				case VSI_IMAGE_WAV:
					if (VARIANT_FALSE != V_BOOL(&vFrame))
					{
						bValidType = FALSE;
					}
					break;
				}

				// If RF type selected then we need to have a RF frame
				// We allow raw types to go through
				if (dwType == VSI_IMAGE_IQ_DATA_LOOP ||
					dwType == VSI_IMAGE_IQ_DATA_FRAME ||
					dwType == VSI_IMAGE_RF_DATA_LOOP ||
					dwType == VSI_IMAGE_RF_DATA_FRAME)
				{
					if (VARIANT_FALSE == V_BOOL(&vRf))
					{
						bValidType = FALSE;
					}
				}

				if (b3D)
				{
					if (dwType == VSI_DICOM_JPEG_BASELINE ||
						dwType == VSI_DICOM_RLE_LOSSLESS ||
						dwType == VSI_DICOM_JPEG_LOSSLESS ||
						dwType == VSI_EXPORT_TYPE_PHYSIO)
					{
						bValidType = FALSE;
					}

					if (!VsiModeUtils::GetIsEngineerMode(m_pPdm))
					{
						if (dwType == VSI_DICOM_IMPLICIT_VR_LITTLE_ENDIAN ||
							dwType == VSI_DICOM_EXPLICIT_VR_LITTLE_ENDIAN)
						{
							bValidType = FALSE;
						}
					}
				}

				if (dwType == VSI_IMAGE_TIFF_LOOP && FALSE == b3D)
				{
					bValidType = FALSE;
				}

				CComVariant vName;
				hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				SYSTEMTIME stDate;
				CComVariant vDate;
				hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vDate);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				COleDateTime date(vDate);
				date.GetAsSystemTime(stDate);

				SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

				CString strTime;
				strTime.Format(L"%04d-%02d-%02d-%02d-%02d-%02d",
					stDate.wYear, stDate.wMonth, stDate.wDay,
					stDate.wHour, stDate.wMinute, stDate.wSecond);

				if (NULL != *szFileName)
				{
					// Find unique file name
					BOOL bNotUnique = TRUE;

					if (NULL != *V_BSTR(&vName))
					{
						int iUniqueDigit(1);

						// Check to see if the label has special characters
						CString strImageLabel(V_BSTR(&vName));
						{
							LPCWSTR ppszSpecial[] = {
								L"/", L"\\", L":", L"*", L"?", L"\"", L"<", L">", L"|"};

							for (int i=0; i<_countof(ppszSpecial); ++i)
							{
								strImageLabel.Replace(ppszSpecial[i], L"_");
							}
						}

						do
						{
							swprintf_s(szPath, _countof(szPath), L"%s\\%s_%s-%s_%d.%s",
								(LPCWSTR)m_strExportFolder, szFileName, (LPCWSTR)strImageLabel, (LPCWSTR)strTime,
								iUniqueDigit++, pszExt);

							bNotUnique = (_waccess(szPath, 0) == 0);
						}
						while (bNotUnique);
					}
					else
					{
						int iUniqueDigit(1);

						do
						{
							swprintf_s(szPath, _countof(szPath), L"%s\\%s_%s_%d.%s",
								(LPCWSTR)m_strExportFolder, szFileName, (LPCWSTR)strTime,
								iUniqueDigit++, pszExt);

							bNotUnique = (_waccess(szPath, 0) == 0);
						}
						while (bNotUnique);
					}
				}
				else
				{
					// Find unique file name
					BOOL bNotUnique = TRUE;

					if (NULL != *V_BSTR(&vName))
					{
						int iUniqueDigit(1);

						do
						{
							swprintf_s(szPath, _countof(szPath), L"%s\\%s-%s_%d.%s",
								(LPCWSTR)m_strExportFolder, V_BSTR(&vName), (LPCWSTR)strTime,
								iUniqueDigit++, pszExt);

							bNotUnique = (_waccess(szPath, 0) == 0);
						}
						while (bNotUnique);
					}
					else
					{
						int iUniqueDigit(1);

						do
						{
							swprintf_s(szPath, _countof(szPath), L"%s\\%s_%d.%s",
								(LPCWSTR)m_strExportFolder, strTime,
								iUniqueDigit++, pszExt);

							bNotUnique = (_waccess(szPath, 0) == 0);
						}
						while (bNotUnique);
					}
				}
			}
			else
			{
				swprintf_s(szPath, _countof(szPath), L"%s\\%s.%s",
					(LPCWSTR)m_strExportFolder, szFileName, pszExt);
			}
		}
		else
		{
			bValidType = FALSE;
		}

		if (bValidType)
		{
			CComPtr<IVsiPropertyBag> pProperties;
			hr = pProperties.CoCreateInstance(__uuidof(VsiPropertyBag));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Setup
			{
				// Allow all exports to access progress
				CComVariant vProgress((ULONG)(ULONG_PTR)(HWND)m_pExportProgress->m_wndProgress);
				hr = pProperties->Write(VSI_EXPORT_PROP_PROGRESS, &vProgress);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Pass the xml document for error reporting
				CComVariant vReport(m_pExportProgress->m_pXmlDoc);
				hr = pProperties->Write(VSI_EXPORT_PROP_REPORT, &vReport);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (item.m_pImage != NULL)
				{
					// Log
					{
						CComHeapPtr<OLECHAR> pszImageFile;
						item.m_pImage->GetDataPath(&pszImageFile);

						VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Exporting %s image: from %s", (LPWSTR)(LPCWSTR)strModeName, (LPCWSTR)pszImageFile));
					}

					CComVariant vImage(item.m_pImage);
					pProperties->Write(VSI_EXPORT_PROP_IMAGE, &vImage);

					// Acquisition date
					CComVariant vDate;
					hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vDate);
					if (S_OK == hr && VT_NULL != V_VT(&vDate))
					{
						SYSTEMTIME stDate;

						COleDateTime date(vDate);
						date.GetAsSystemTime(stDate);

						SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

						date = stDate;

						vDate.Clear();
						V_VT(&vDate) = VT_DATE;
						V_DATE(&vDate) = DATE(date);

						hr = pProperties->Write(L"createdDate", &vDate);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}

					// Display mode name
					CComVariant vModeName;
					hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vModeName);
					if (S_OK == hr && VT_NULL != V_VT(&vModeName))
					{
						hr = pProperties->Write(L"modeName", &vModeName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}

					// Name
					CComVariant vName;
					hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK != hr || (0 == *V_BSTR(&vName)))
					{
						CString strName;
						strName.Format(L"Image %d", m_pExportProgress->m_iImage + 1);
						vName = (LPCWSTR)strName;
					}

					hr = pProperties->Write(L"name", &vName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				if (item.m_pMode != NULL)
				{
					if (NULL == item.m_pImage)
					{
						// Log
						VSI_LOG_MSG(VSI_LOG_INFO, L"Exporting image: from mode");

						CString strName;
						strName.Format(L"Image %d", m_pExportProgress->m_iImage + 1);
						CComVariant vName((LPCWSTR)strName);

						hr = pProperties->Write(L"name", &vName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}

					CComVariant vMode(item.m_pMode);
					pProperties->Write(L"mode", &vMode);
				}

				CComVariant vType(dwType);
				pProperties->Write(L"type", &vType);

				// Get fcc code / only required for AVI types
				if (VSI_IMAGE_AVI_CUSTOM == dwType)
				{
					CComVariant vFcc(dwFcc);
					pProperties->Write(L"fcc", &vFcc);
				}

				// General quality settings
				CComVariant vQuality(dwAviQuality);
				pProperties->Write(L"quality", &vQuality);

				// Write expected approximate resolution
				switch (dwType & VSI_IMAGE_EXPORT_MASK)
				{
				case VSI_IMAGE_AVI_CUSTOM:
					{
						CComVariant vYRes(dwExportResInitial);
						pProperties->Write(L"yRes", &vYRes);

						// AVI's come in multi resolution
						if (VSI_SYS_EXPORT_AVI_QUALITY_HIGH == dwAviQuality)
						{
							CComVariant vScale(dbExportScaleHigh);
							pProperties->Write(L"scale", &vScale);
						}
						else
						{
							CComVariant vScale(dbExportScaleMed);
							pProperties->Write(L"scale", &vScale);
						}
					}
					break;

				case VSI_DICOM_IMPLICIT_VR_LITTLE_ENDIAN:
				case VSI_DICOM_EXPLICIT_VR_LITTLE_ENDIAN:
				case VSI_DICOM_JPEG_BASELINE:
				case VSI_DICOM_RLE_LOSSLESS:
				case VSI_DICOM_JPEG_LOSSLESS:
					{
						CComVariant vYRes(dwExportResInitial);
						pProperties->Write(L"yRes", &vYRes);

						CComVariant vScale(dbExportScaleHigh);
						pProperties->Write(L"scale", &vScale);
					}
					break;
				}

				CComVariant vPath(szPath);
				pProperties->Write(L"path", &vPath);

				if (VSI_EXPORT_IMAGE_FLAG_DICOM == m_dwExportTypeMain)
				{
					BOOL bExportRegions = BST_CHECKED == m_pSubPanelDicom->IsDlgButtonChecked(IDC_IE_REGIONS);
					CComVariant vExportRegions(FALSE != bExportRegions);
					hr = pProperties->Write(VSI_EXPORT_PROP_REGIONS, &vExportRegions);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}

			// Log
			VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Exporting image: to %s", szPath));

			// This is the actual command that performs the export
			CComVariant v(pProperties.p);
			hr = m_pCmdMgr->InvokeCommand(ID_CMD_DO_EXPORT_IMAGE, &v);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Checks return
			CComVariant vReturn;
			m_pExportProgress->m_bError = FALSE;
			hr = pProperties->Read(VSI_EXPORT_PROP_RETURN, &vReturn);
			if ((S_OK != hr) || (VARIANT_FALSE == V_BOOL(&vReturn)))
				m_pExportProgress->m_bError = TRUE;

			// Checks forceReport
			CComVariant vForce;
			hr = pProperties->Read(L"forceReport", &vForce);
			if ((S_OK != hr) || (VARIANT_TRUE == V_BOOL(&vForce)))
				m_pExportProgress->m_bForceReport = TRUE;
		}
		else  // Not valid type
		{
			CComPtr<IXMLDOMNode> pImages;
			CComPtr<IXMLDOMElement> pElmImage;

			hr = m_pExportProgress->m_pXmlDoc->selectSingleNode(CComBSTR(VSI_IMAGE_XML_ELM_IMAGES), &pImages);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			hr = m_pExportProgress->m_pXmlDoc->createElement(CComBSTR(VSI_IMAGE_XML_ELM_IMAGE), &pElmImage);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

			hr = pElmImage->setAttribute(
				CComBSTR(VSI_IMAGE_XML_ATT_EXPORT_STATUS),
				CComVariant(VSI_SEI_EXPORT_NOTIMPL));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComVariant vName;
			hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK != hr || (0 == *V_BSTR(&vName)))
			{
				CString strName;
				strName.Format(L"Image %d", m_pExportProgress->m_iImage + 1);
				vName = (LPCWSTR)strName;
			}

			hr = pElmImage->setAttribute(CComBSTR(VSI_IMAGE_XML_ATT_NAME), vName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pImages->appendChild(pElmImage, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_pExportProgress->m_bForceReport = TRUE;
		}
	}
	VSI_CATCH(hr);

	PostMessage(WM_VSI_EXPORT_IMAGE_COMPLETED);

	return 0;
}

LRESULT CVsiExportImage::OnExportImageCompleted(UINT, WPARAM, LPARAM, BOOL&)
{
	++m_pExportProgress->m_iImage;
	++m_pExportProgress->m_iterExport;

	if (m_pExportProgress->m_iterExport == m_listImage.end() ||
		m_pExportProgress->m_bCancel)
	{
		// dump all left over images into a special place in the report so we can inform about them
		if (m_pExportProgress->m_iterExport != m_listImage.end() && m_pExportProgress->m_bCancel)
		{
			HRESULT hr = S_OK;

			try
			{
				CComPtr<IXMLDOMNode> pImages;
				hr = m_pExportProgress->m_pXmlDoc->selectSingleNode(CComBSTR(VSI_IMAGE_XML_ELM_IMAGES), &pImages);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				int imageIndex = m_pExportProgress->m_iImage;
				CVsiListImageIter iter;
				for (iter = m_pExportProgress->m_iterExport; iter != m_listImage.end(); iter++)
				{
					imageIndex++;
					CVsiImage &item = *iter;
					CComPtr<IXMLDOMElement> pElmImage;
					hr = m_pExportProgress->m_pXmlDoc->createElement(CComBSTR(VSI_IMAGE_XML_ELM_IMAGE), &pElmImage);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

					hr = pImages->appendChild(pElmImage, NULL);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Name
					CComVariant vName;
					hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK != hr || (0 == *V_BSTR(&vName)))
					{
						CString strName;
						strName.Format(L"Image %d", imageIndex);
						vName = (LPCWSTR)strName;
					}
					hr = pElmImage->setAttribute(CComBSTR(VSI_IMAGE_XML_ATT_NAME), vName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pElmImage->setAttribute(CComBSTR(VSI_IMAGE_XML_ATT_EXPORT_STATUS), CComVariant(VSI_SEI_EXPORT_CANCELED));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComVariant vDate;
					hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &vDate);
					if (S_OK == hr && VT_NULL != V_VT(&vDate))
					{
						SYSTEMTIME stDate;

						COleDateTime date(vDate);
						date.GetAsSystemTime(stDate);

						SystemTimeToTzSpecificLocalTime(NULL, &stDate, &stDate);

						date = stDate;

						vDate.Clear();
						V_VT(&vDate) = VT_DATE;
						V_DATE(&vDate) = DATE(date);

						hr = pElmImage->setAttribute(CComBSTR(VSI_IMAGE_XML_ATT_CREATED_DATE), vDate);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}

					CComVariant vModeName;
					hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_MODE_DISPLAY, &vModeName);
					if (S_OK == hr && VT_NULL != V_VT(&vModeName))
					{
						hr = pElmImage->setAttribute(CComBSTR(VSI_IMAGE_XML_ATT_MODE_NAME), vModeName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}
			}
			VSI_CATCH(hr);
		}

		// Flush this volume
		VsiFlushVolume(m_strExportFolder);

		// Force progress to end
		{
			// make sure the progress is at max
			int iMax = (int)m_pExportProgress->m_wndProgress.SendMessage(
				WM_VSI_PROGRESS_GETMAX);

			m_pExportProgress->m_wndProgress.SendMessage(
				WM_VSI_PROGRESS_SETPOS, iMax);
		}

		Sleep(500);  // Show the progress for another 0.5sec

		// Shut progress panel down
		m_pExportProgress->m_wndProgress.SendMessage(WM_CLOSE);

		if (m_pExportProgress->m_bError || m_pExportProgress->m_bForceReport)
		{
			// Report errors
			CVsiReportExportStatus statusReport(
				m_pExportProgress->m_pXmlDoc,
				CVsiReportExportStatus::REPORT_TYPE_IMAGE_EXPORT);
			statusReport.DoModal(GetTopLevelParent());
		}

		delete m_pExportProgress;
		m_pExportProgress = NULL;
	}
	else
	{
		//if (VSI_EXPORT_IMAGE_FLAG_CINE != m_dwExportTypeMain)
		{
			// Each export image comprises of a number of steps - only increment single for images
			// which have one step - cine loops comprise of multiple steps
			m_pExportProgress->m_wndProgress.SendMessage(
				WM_VSI_PROGRESS_SETPOS, m_pExportProgress->m_iImage, 0);
		}

		// Pump some messages so that the user can press the Cancel button
		MSG msg;
		while (PeekMessage(&msg, NULL, 0, 0, TRUE))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}

		// Do the next image in the queue
		PostMessage(WM_VSI_EXPORT_IMAGE);
	}

	return 0;
}

LRESULT CVsiExportImage::OnBnClickedOK(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Export Image - OK");

	// Do export
	HRESULT hr = S_OK;

	try
	{
		CWindow wndExportType(GetTypeControl());
		int iIndex = (int)wndExportType.SendMessage(CB_GETCURSEL, NULL, NULL);
		if (iIndex < 0)
		{
			// Shouldn't happen
			return 0;
		}

		CString strTargetPath;
		CString strTargetFile;
		CVsiImportExportUIHelper::VSI_NAME_PARSE ret =
			CVsiImportExportUIHelper::ParseUserSuppliedName(strTargetPath, strTargetFile);
		if (VSI_NAME_PARSE_BAD_NAME == ret)
		{
			m_wndName.ActivateTips(CVsiPathNameEditWnd::VSI_TIPS_NAME);
			return 0;
		}
		else if (VSI_NAME_PARSE_INCOMPLETE_PATH == ret)
		{
			return 0;
		}
		else if (VSI_NAME_PARSE_COMPLETE_PATH == ret)
		{
			CWaitCursor wait;

			m_iAvailableSize = VsiGetSpaceAvailableAtPath(strTargetPath);
			UpdateSizeUI();
		}

		if (m_iAvailableSize < m_iRequiredSize)
		{
			// Not enough space
			ShowNotEnoughSpaceMessage();
			return 0;
		}

		// Read only test
		if (CheckReadOnly(strTargetPath))
		{
			ShowReadOnlyMessage();
			return 0;
		}

		// Overwrite test
		if (CheckForOverwrite(strTargetPath, strTargetFile) > 0)
		{
			// Confirm overwrite
			int iRet = ShowOverwriteMessage();
			if (IDYES != iRet)
			{
				return 0;
			}
		}

		m_strExportFolder = strTargetPath;

		m_wndName.SetWindowText(strTargetFile);

		// Will automatically exit so don't let operator press any buttons
		GetDlgItem(IDOK).EnableWindow(FALSE);
		GetDlgItem(IDCANCEL).EnableWindow(FALSE);
		GetDlgItem(IDC_IE_NEW_FOLDER).EnableWindow(FALSE);

		// Save Auto Complete settings
		if (!strTargetFile.IsEmpty() && NULL != m_pAutoCompleteSource)
		{
			WCHAR pszPath[MAX_PATH];
			BOOL bRet = VsiGetApplicationDataDirectory(
				AtlFindStringResourceInstance(IDS_VSI_PRODUCT_NAME),
				pszPath);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, NULL);

			PathAppend(pszPath, VSI_AUTOCOMPLETE_NAME);

			hr = m_pAutoCompleteSource->Push(strTargetFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pAutoCompleteSource->SaveSettings(pszPath, 10);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		DWORD dwType = (DWORD)wndExportType.SendMessage(CB_GETITEMDATA, iIndex, NULL);

		// Saves last used type
		if (VSI_EXPORT_IMAGE_FLAG_CINE == m_dwExportTypeMain)
		{
			hr = m_pSubPanelCine->SaveSettings();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (VSI_EXPORT_IMAGE_FLAG_IMAGE == m_dwExportTypeMain)
		{
			hr = m_pSubPanelFrame->SaveSettings();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (VSI_EXPORT_IMAGE_FLAG_DICOM == m_dwExportTypeMain)
		{
			hr = m_pSubPanelDicom->SaveSettings();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (VSI_EXPORT_IMAGE_FLAG_REPORT == m_dwExportTypeMain)
		{
			hr = m_pSubPanelReport->SaveSettings();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (VSI_EXPORT_IMAGE_FLAG_GRAPH == m_dwExportTypeMain)
		{
			hr = m_pSubPanelGraph->SaveSettings();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (VSI_EXPORT_IMAGE_FLAG_TABLE == m_dwExportTypeMain)
		{
			hr = m_pSubPanelTable->SaveSettings();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}
		else if (VSI_EXPORT_IMAGE_FLAG_PHYSIO == m_dwExportTypeMain)
		{
			hr = m_pSubPanelPhysio->SaveSettings();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		// Save export type
		VsiSetRangeValue<DWORD>(m_dwExportTypeMain,
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
			VSI_PARAMETER_SYS_DEFAULT_EXPORT_TYPE, m_pPdm);

		switch (dwType & VSI_IMAGE_EXPORT_MASK)
		{
		case VSI_IMAGE_TIFF_FULL:
		case VSI_IMAGE_BMP_FULL:
		case VSI_IMAGE_TIFF_IMAGE:
		case VSI_IMAGE_BMP_IMAGE:
		case VSI_IMAGE_AVI_CUSTOM:
		case VSI_IMAGE_IQ_DATA_LOOP:
		case VSI_IMAGE_IQ_DATA_FRAME:
		case VSI_IMAGE_RF_DATA_LOOP:
		case VSI_IMAGE_RF_DATA_FRAME:
		case VSI_IMAGE_RAW_DATA_LOOP:
		case VSI_IMAGE_RAW_DATA_FRAME:
		case VSI_IMAGE_WAV:
		case VSI_DICOM_IMPLICIT_VR_LITTLE_ENDIAN:
		case VSI_DICOM_EXPLICIT_VR_LITTLE_ENDIAN:
		case VSI_DICOM_JPEG_BASELINE:
		case VSI_DICOM_RLE_LOSSLESS:
		case VSI_DICOM_JPEG_LOSSLESS:
		case VSI_IMAGE_TIFF_LOOP:
			{
				VSI_LOG_MSG(VSI_LOG_INFO, L"Export images: started");

				int iNumItems = (int)m_listImage.size();
				VSI_VERIFY(iNumItems > 0, VSI_LOG_ERROR, L"Invalid number of items in export");

				// Allow commands from this progress
				VsiProgressDialogBox(
					NULL,
					(LPCWSTR)MAKEINTRESOURCE(VSI_PROGRESS_TEMPLATE_MPC),
					GetTopLevelParent(),
					ExportImageCallback,
					this);

				VSI_LOG_MSG(VSI_LOG_INFO, L"Export images: ended");
			}
			break;

		case VSI_EXPORT_TABLE:
			{
				VSI_LOG_MSG(VSI_LOG_INFO, L"Export table: started");

				WCHAR szPath[MAX_PATH];
				swprintf_s(szPath, L"%s\\%s.%s",
					strTargetPath, strTargetFile, GetFileExtension());

				VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Exporting table: %s", szPath));

				hr = m_ExportTable.DoExport(szPath);
				if (FAILED(hr))
				{
					CString strMessage(L"Table export failed.");
					if (HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) == hr)
					{
						strMessage += L"\n\nAccess Denied.";
					}
					else if (HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) == hr)
					{
						strMessage += L"\n\nFile is in use.";
					}
					// Report failure
					VsiMessageBox(
						GetTopLevelParent(),
						strMessage,
						VsiGetRangeValue<CString>(
							VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
							VSI_PARAMETER_SYS_PRODUCT_NAME,
							m_pPdm),
						MB_OK | MB_ICONWARNING);
				}

				VSI_LOG_MSG(VSI_LOG_INFO, L"Export table: ended");
			}
			break;

		case VSI_COMBINED_REPORT:
			{
				VSI_LOG_MSG(VSI_LOG_INFO, L"Export report: started");

				WCHAR szPath[MAX_PATH];
				swprintf_s(szPath, _countof(szPath), L"%s\\%s.%s",
					strTargetPath, strTargetFile, GetFileExtension());

				VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Exporting report: %s", szPath));

				// Hacking export to create Measurement Report remotely
				if ((VSI_EXPORT_IMAGE_FLAG_REPORT & m_dwFlags) && (NULL == m_pXmlReportDoc))
				{
					hr = CreateMsmntReport(m_pApp, &m_pXmlReportDoc, m_listImage);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				CComVariant vXmlDoc(m_pXmlReportDoc);
				hr = m_ExportReport.DoExport(szPath, &vXmlDoc, m_pApp);

				if (FAILED(hr))
				{
					// Delete file
					DeleteFile(szPath);

					CString strMessage(L"Failed to export the measurement report.");
					if (HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) == hr)
					{
						strMessage += L"\n\nAccess Denied.";
					}
					else if (HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) == hr)
					{
						strMessage += L"\n\nFile is in use.";
					}
					// Report failure
					VsiMessageBox(
						GetTopLevelParent(),
						strMessage,
						VsiGetRangeValue<CString>(
							VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
							VSI_PARAMETER_SYS_PRODUCT_NAME,
							m_pPdm),
						MB_OK | MB_ICONWARNING);

					UpdateUI();

					return 0;
				}

				VSI_LOG_MSG(VSI_LOG_INFO, L"Export report: ended");
			}
			break;

		case VSI_EXPORT_GRAPH:
			{
				VSI_LOG_MSG(VSI_LOG_INFO, L"Export graph: started");

				// Should we export graph parameters
				BOOL bExportGraphParam = m_pSubPanelGraph->IsDlgButtonChecked(IDC_IE_EXPORT_GRAPH_PARAM);

				if (m_pSubPanelGraph->IsDlgButtonChecked(IDC_IE_EXPORT_CSV))
				{
					WCHAR szPath[MAX_PATH];
					swprintf_s(szPath, _countof(szPath), L"%s\\%s.csv",
						strTargetPath.GetString(), strTargetFile.GetString());

					VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Exporting CSV: %s", szPath));

					m_pAnalysisExport->ExportAsCSV(szPath);
				}
				if (m_pSubPanelGraph->IsDlgButtonChecked(IDC_IE_EXPORT_TIFF))
				{
					WCHAR szPath[MAX_PATH];
					swprintf_s(szPath, _countof(szPath), L"%s\\%s.tiff",
						strTargetPath.GetString(), strTargetFile.GetString());

					VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Exporting TIFF: %s", szPath));

					m_pAnalysisExport->ExportAsImage(
						szPath,
						VSI_ANALYSIS_EXPORT_FLAGS_TIFF | (bExportGraphParam ? VSI_ANALYSIS_EXPORT_FLAGS_GRAPH_PARAM : 0));
				}
				if (m_pSubPanelGraph->IsDlgButtonChecked(IDC_IE_EXPORT_BMP))
				{
					WCHAR szPath[MAX_PATH];
					swprintf_s(szPath, _countof(szPath), L"%s\\%s.bmp", strTargetPath, strTargetFile);

					VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Exporting BMP: %s", szPath));

					m_pAnalysisExport->ExportAsImage(
						szPath,
						VSI_ANALYSIS_EXPORT_FLAGS_BMP | (bExportGraphParam ? VSI_ANALYSIS_EXPORT_FLAGS_GRAPH_PARAM : 0));
				}

				VSI_LOG_MSG(VSI_LOG_INFO, L"Export graph: ended");
			}
			break;

		case VSI_EXPORT_TYPE_PHYSIO:
			{
				VSI_LOG_MSG(VSI_LOG_INFO, L"Export physio: started");

				int iNumItems = (int)m_listImage.size();
				VSI_VERIFY(iNumItems > 0, VSI_LOG_ERROR, L"Invalid number of items in export");

				// Allow commands from this progress
				VsiProgressDialogBox(
					NULL,
					(LPCWSTR)MAKEINTRESOURCE(VSI_PROGRESS_TEMPLATE_MPC),
					GetTopLevelParent(),
					ExportImageCallback,
					this);

				VSI_LOG_MSG(VSI_LOG_INFO, L"Export physio: ended");
			}
			break;

		default:
			_ASSERT(0);
		}

		m_pCmdMgr->InvokeCommand(ID_CMD_EXPORT_CLOSE, NULL);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiExportImage::OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Export Image - Cancel");

	m_pCmdMgr->InvokeCommand(ID_CMD_EXPORT_CLOSE, NULL);

	return 0;
}

LRESULT CVsiExportImage::OnBnClickedTypeMain(WORD, WORD, HWND, BOOL&)
{
	UpdateTypeMain();

	return 0;
}

LRESULT CVsiExportImage::OnFileNameChanged(WORD, WORD, HWND, BOOL&)
{
	const MSG *pMsg = GetCurrentMessage();

	CWindow wnPanel(GetPanel(m_dwExportTypeMain));

	return wnPanel.SendMessage(pMsg->message, pMsg->wParam, pMsg->lParam);
}

LRESULT CVsiExportImage::OnTreeSelChanged(int, LPNMHDR, BOOL&)
{
	if (m_pFExplorer != NULL)
	{
		UpdateAvailableSpace();
		UpdateSizeUI();
		UpdateUI();
	}
	return 0;
}

DWORD CALLBACK CVsiExportImage::ExportImageCallback(HWND hWnd, UINT uiState, LPVOID pData)
{
	CVsiExportImage *pThis = static_cast<CVsiExportImage*>(pData);

	if (VSI_PROGRESS_STATE_INIT == uiState)
	{
		::SetWindowText(hWnd, L"Image Export");
		::SendMessage(hWnd, WM_VSI_PROGRESS_SETMAX, pThis->m_listImage.size(), 0);
		::SendMessage(hWnd, WM_VSI_PROGRESS_SETMSG, 0, (LPARAM)L"Exporting Image...");

		HRESULT hr(S_OK);

		try
		{
			pThis->m_pExportProgress = new CVsiExportProgress();
			pThis->m_pExportProgress->m_wndProgress = hWnd;

			// Prepare the xml document that is to be used for error reporting
			hr = VsiCreateDOMDocument(&pThis->m_pExportProgress->m_pXmlDoc);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating DOMDocument");

			// Add Root element
			CComPtr<IXMLDOMElement> pElmRoot;
			hr = pThis->m_pExportProgress->m_pXmlDoc->createElement(
				CComBSTR(VSI_IMAGE_XML_ELM_IMAGES), &pElmRoot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pThis->m_pExportProgress->m_pXmlDoc->appendChild(pElmRoot, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			pThis->m_pExportProgress->m_iterExport = pThis->m_listImage.begin();

			pThis->PostMessage(WM_VSI_EXPORT_IMAGE);
		}
		VSI_CATCH(hr);

		if (FAILED(hr))
		{
			::PostMessage(hWnd, WM_CLOSE, 0, 0);
		}
	}
	else if (VSI_PROGRESS_STATE_CANCEL == uiState)
	{
		pThis->m_pExportProgress->m_bCancel = TRUE;
	}
	else if (VSI_PROGRESS_STATE_END == uiState)
	{
		// Nothing
	}

	return 0;
}

CWindow CVsiExportImage::GetPanel(DWORD dwExportType)
{
	CWindow wndPanel;

	switch (dwExportType)
	{
	case VSI_EXPORT_IMAGE_FLAG_CINE:
		wndPanel = *m_pSubPanelCine;
		break;
	case VSI_EXPORT_IMAGE_FLAG_IMAGE:
		wndPanel = *m_pSubPanelFrame;
		break;
	case VSI_EXPORT_IMAGE_FLAG_DICOM:
		wndPanel = *m_pSubPanelDicom;
		break;
	case VSI_EXPORT_IMAGE_FLAG_REPORT:
		wndPanel = *m_pSubPanelReport;
		break;
	case VSI_EXPORT_IMAGE_FLAG_GRAPH:
		wndPanel = *m_pSubPanelGraph;
		break;
	case VSI_EXPORT_IMAGE_FLAG_TABLE:
		wndPanel = *m_pSubPanelTable;
		break;
	case VSI_EXPORT_IMAGE_FLAG_PHYSIO:
		wndPanel = *m_pSubPanelPhysio;
		break;
	}

	return wndPanel;
}

CWindow CVsiExportImage::GetTypeControl()
{
	CWindow wndPanel(GetPanel(m_dwExportTypeMain));
	return wndPanel.GetDlgItem(IDC_IE_TYPE);
}

int CVsiExportImage::GetTypeCount(int *piFrames, int *piLoops, int *pi3Ds) throw(...)
{
	HRESULT hr;

	if (NULL != piFrames || NULL != piLoops)
	{
		int iFrames = 0;
		int iLoops = 0;
		int i3Ds = 0;

		CVsiListImageIter iter;
		for (iter = m_listImage.begin(); iter != m_listImage.end(); ++iter)
		{
			CVsiImage &item = *iter;

			if (NULL != item.m_pImage)
			{
				VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_NONE);
				item.m_pImage->GetErrorCode(&dwErrorCode);
				if (VSI_IMAGE_ERROR_NONE == dwErrorCode)
				{
					CComVariant vFrame;
					hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_FRAME_TYPE, &vFrame);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					if (VARIANT_FALSE != V_BOOL(&vFrame))
					{
						++iFrames;
					}
					else
					{
						CComVariant vMode;
						hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						CComVariant v3d;
						hr = m_pModeMgr->GetModePropertyFromInternalName(
							V_BSTR(&vMode),
							L"3d",
							&v3d);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						if (VARIANT_FALSE == V_BOOL(&v3d))
						{
							++iLoops;
						}
						else
						{
							++i3Ds;
						}
					}
				}
			}
			else if (NULL != item.m_pMode)
			{
				CComQIPtr<IVsiPropertyBag> pProps(item.m_pMode);
				VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

				CComVariant vRoot;
				hr = pProps->ReadId(VSI_PROP_MODE_ROOT, &vRoot);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				BOOL bIsFrame = VsiGetBooleanValue<BOOL>(
					V_BSTR(&vRoot), VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_SETTINGS_SAVED_FRAME, m_pPdm);
				if (bIsFrame)
				{
					++iFrames;
				}
				else
				{
					CComQIPtr<IVsiPropertyBag> pProps(item.m_pMode);
					VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

					CComVariant vModeId;
					hr = pProps->ReadId(VSI_PROP_MODE_MODE_ID, &vModeId);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					CComVariant v3d;
					hr = m_pModeMgr->GetModePropertyFromId(
						V_UI4(&vModeId),
						L"3d",
						&v3d);
					VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

					if (VARIANT_FALSE == V_BOOL(&v3d))
					{
						++iLoops;
					}
					else
					{
						++i3Ds;
					}
				}
			}
		}

		if (NULL != piFrames)
			*piFrames = iFrames;

		if (NULL != piLoops)
			*piLoops = iLoops;

		if (NULL != pi3Ds)
			*pi3Ds = i3Ds;
	}

	return (int)m_listImage.size();
}

void CVsiExportImage::UpdateTypeMain(BOOL bForceUpdate)
{
	DWORD dwTypeMainNew = VSI_EXPORT_IMAGE_FLAG_NONE;

	// Updates combo
	if (IsDlgButtonChecked(IDC_IE_CINE_LOOP))
	{
		dwTypeMainNew = VSI_EXPORT_IMAGE_FLAG_CINE;
	}
	else if (IsDlgButtonChecked(IDC_IE_IMAGE))
	{
		dwTypeMainNew = VSI_EXPORT_IMAGE_FLAG_IMAGE;
	}
	else if (IsDlgButtonChecked(IDC_IE_DICOM))
	{
		dwTypeMainNew = VSI_EXPORT_IMAGE_FLAG_DICOM;
	}
	else if (IsDlgButtonChecked(IDC_IE_REPORT))
	{
		dwTypeMainNew = VSI_EXPORT_IMAGE_FLAG_REPORT;
	}
	else if (IsDlgButtonChecked(IDC_IE_TABLE))
	{
		dwTypeMainNew = VSI_EXPORT_IMAGE_FLAG_TABLE;
	}
	else if (IsDlgButtonChecked(IDC_IE_GRAPH))
	{
		dwTypeMainNew = VSI_EXPORT_IMAGE_FLAG_GRAPH;
	}
	else if (IsDlgButtonChecked(IDC_IE_PHYSIO))
	{
		dwTypeMainNew = VSI_EXPORT_IMAGE_FLAG_PHYSIO;
	}

	if (bForceUpdate || dwTypeMainNew != m_dwExportTypeMain)
	{
		DWORD dwTypeOld = m_dwExportTypeMain;
		m_dwExportTypeMain = (dwTypeMainNew & m_dwFlags);

		if (0 == m_dwExportTypeMain)
		{
			// Just incase
			m_dwExportTypeMain = VSI_EXPORT_IMAGE_FLAG_CINE;
		}

		CWindow wnPanel(GetPanel(m_dwExportTypeMain));

		CWindow wndSaveAs(wnPanel.GetDlgItem(IDC_IE_SAVEAS_DUMMY));

		CRect rc;
		wndSaveAs.GetWindowRect(&rc);
		wnPanel.ScreenToClient(&rc);

		switch (m_dwExportTypeMain)
		{
		case VSI_EXPORT_IMAGE_FLAG_CINE:
			{
				SetWindowText(L"Export Image");

				if (m_listImage.size() > 1)
				{
					CString strSelected;
					strSelected.Format(L"%d Images Selected", m_listImage.size());
					SetDlgItemText(IDC_IE_SELECTED, strSelected);
				}
				else if (m_listImage.size() == 1)
				{
					SetDlgItemText(IDC_IE_SELECTED, L"1 Image Selected");
				}
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_IMAGE:
			{
				SetWindowText(L"Export Image");

				if (m_listImage.size() > 1)
				{
					CString strSelected;
					strSelected.Format(L"%d Images Selected", m_listImage.size());
					SetDlgItemText(IDC_IE_SELECTED, strSelected);
				}
				else if (m_listImage.size() == 1)
				{
					SetDlgItemText(IDC_IE_SELECTED, L"1 Image Selected");
				}
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_DICOM:
			{
				SetWindowText(L"Export Image");

				if (m_listImage.size() > 1)
				{
					CString strSelected;
					strSelected.Format(L"%d Images Selected", m_listImage.size());
					SetDlgItemText(IDC_IE_SELECTED, strSelected);
				}
				else if (m_listImage.size() == 1)
				{
					SetDlgItemText(IDC_IE_SELECTED, L"1 Image Selected");
				}
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_PHYSIO:
			{
				SetWindowText(L"Export Physiological Data");

				if (m_listImage.size() > 1)
				{
					CString strSelected;
					strSelected.Format(L"%d Images Selected", m_listImage.size());
					SetDlgItemText(IDC_IE_SELECTED, strSelected);
				}
				else if (m_listImage.size() == 1)
				{
					SetDlgItemText(IDC_IE_SELECTED, L"1 Image Selected");
				}
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_REPORT:
			{
				SetWindowText(L"Export Report");
				SetDlgItemText(IDC_IE_SELECTED, L"Information");
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_TABLE:
			{
				SetWindowText(L"Export Table");
				SetDlgItemText(IDC_IE_SELECTED, L"Information");
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_GRAPH:
			{
				SetWindowText(L"Export Graph");
				SetDlgItemText(IDC_IE_SELECTED, L"Information");
			}
			break;

		default:
			_ASSERT(0);
		}

		m_wndName.SetParent(wnPanel);
		m_wndName.SetWindowPos(wndSaveAs, rc.left, rc.top, rc.Width(), rc.Height(), SWP_NOACTIVATE);

		HDWP hwdp = BeginDeferWindowPos(2);

		hwdp = GetPanel(m_dwExportTypeMain).DeferWindowPos(
			hwdp, NULL, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_SHOWWINDOW);
		if (dwTypeOld != m_dwExportTypeMain)
		{
			hwdp = GetPanel(dwTypeOld).DeferWindowPos(
				hwdp, NULL, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_HIDEWINDOW);
		}

		EndDeferWindowPos(hwdp);

		UpdateWindow();

		// Refresh UI
		UpdateRequiredSize();
	}
}

void CVsiExportImage::UpdateRequiredSize()
{
	m_iRequiredSize = static_cast<INT64>(EstimateSize());

	// Enable the dialog buttons based on the current selection.
	UpdateSizeUI();
	UpdateUI();
}

void CVsiExportImage::UpdateUI()
{
	CVsiImportExportUIHelper<CVsiExportImage>::UpdateUI();

	BOOL bOkayEnabled = TRUE;

	switch (m_dwExportTypeMain)
	{
	case VSI_EXPORT_IMAGE_FLAG_CINE:
		{
			bOkayEnabled = m_pSubPanelCine->UpdateUI();
		}
		break;

	case VSI_EXPORT_IMAGE_FLAG_IMAGE:
		{
			bOkayEnabled = m_pSubPanelFrame->UpdateUI();
		}
		break;

	case VSI_EXPORT_IMAGE_FLAG_DICOM:
		{
			bOkayEnabled = m_pSubPanelDicom->UpdateUI();
		}
		break;

	case VSI_EXPORT_IMAGE_FLAG_REPORT:
		{
			bOkayEnabled = m_pSubPanelReport->UpdateUI();
		}
		break;

	case VSI_EXPORT_IMAGE_FLAG_GRAPH:
		{
			bOkayEnabled = m_pSubPanelGraph->UpdateUI();
		}
		break;

	case VSI_EXPORT_IMAGE_FLAG_TABLE:
		{
			bOkayEnabled = m_pSubPanelTable->UpdateUI();
		}
		break;

	case VSI_EXPORT_IMAGE_FLAG_PHYSIO:
		{
			bOkayEnabled = m_pSubPanelPhysio->UpdateUI();
		}
		break;

	default:
		_ASSERT(0);
	}

	CString strFileName;
	m_wndName.GetWindowText(strFileName);

	if (strFileName.IsEmpty())
	{
		bOkayEnabled = FALSE;
	}
	else
	{
		if (CVsiImportExportUIHelper::IsAbsolutePath())
		{
			bOkayEnabled = TRUE;
		}
		else if (bOkayEnabled)
		{
			CComHeapPtr<OLECHAR> pszFolderName;
			HRESULT hr = m_pFExplorer->GetSelectedFilePath(&pszFolderName);
			bOkayEnabled = ((S_OK == hr) && (*pszFolderName != 0));
		}
	}

	GetDlgItem(IDOK).EnableWindow(bOkayEnabled);
}

int CVsiExportImage::SetTextBoxLimit()
{
	int iRemainingCharacters(0);

	int iPathLength = lstrlen(GetPath());

	iRemainingCharacters = VSI_MAX_FOLDER_LENGTH - iPathLength;
	if (iRemainingCharacters < 0)
		iRemainingCharacters = 0;
	else if (iRemainingCharacters > 100)
		iRemainingCharacters = 100;

	// Limits text to help prevent overrun
	m_wndName.SendMessage(EM_SETLIMITTEXT, iRemainingCharacters);

	return iRemainingCharacters;
}

LPCWSTR CVsiExportImage::GetFileExtension()
{
	LPCWSTR pszExt = NULL;

	switch (m_dwExportTypeMain)
	{
	case VSI_EXPORT_IMAGE_FLAG_CINE:
	case VSI_EXPORT_IMAGE_FLAG_IMAGE:
	case VSI_EXPORT_IMAGE_FLAG_DICOM:
	case VSI_EXPORT_IMAGE_FLAG_PHYSIO:
	case VSI_EXPORT_IMAGE_FLAG_REPORT:
	case VSI_EXPORT_IMAGE_FLAG_TABLE:
		{
			CWindow wndExportType(GetTypeControl());

			int iIndex = (int)wndExportType.SendMessage(CB_GETCURSEL, NULL, NULL);
			DWORD dwType = (DWORD)wndExportType.SendMessage(CB_GETITEMDATA, iIndex, NULL);

			switch (dwType)
			{
			case VSI_IMAGE_TIFF_LOOP:
			case VSI_IMAGE_TIFF_FULL:
				pszExt = L"tif";
				break;
			case VSI_IMAGE_BMP_FULL:
				pszExt = L"bmp";
				break;
			case VSI_IMAGE_TIFF_IMAGE:
				pszExt = L"tif";
				break;
			case VSI_IMAGE_BMP_IMAGE:
				pszExt = L"bmp";
				break;

			case VSI_IMAGE_RAW_DATA_FRAME:
			case VSI_IMAGE_RAW_DATA_LOOP:
				pszExt = VSI_RF_FILE_EXTENSION_RAWXML;			// Multiple file extensions will be used for this case
				break;
			case VSI_IMAGE_RF_DATA_FRAME:
			case VSI_IMAGE_RF_DATA_LOOP:
				pszExt = VSI_RF_FILE_EXTENSION_RFXML;			// Multiple file extensions will be used for this case
				break;
			case VSI_IMAGE_IQ_DATA_FRAME:
			case VSI_IMAGE_IQ_DATA_LOOP:
				pszExt = VSI_RF_FILE_EXTENSION_IQXML;			// Multiple file extensions will be used for this case
				break;
			case VSI_IMAGE_AVI_CUSTOM:
				{
					DWORD dwFcc = m_pSubPanelCine->m_pdwFccHandlerType[iIndex];
					if (dwFcc == VSI_4CC_ANIMGIF)
					{
						pszExt = L"gif";
					}
					else
					{
						pszExt = L"avi";
					}
				}
				break;
			case VSI_IMAGE_WAV:
				pszExt = L"wav";
				break;
			case VSI_COMBINED_REPORT:
				pszExt = L"csv";
				break;
			case VSI_EXPORT_TABLE:
				pszExt = L"txt";
				break;
			case VSI_DICOM_IMPLICIT_VR_LITTLE_ENDIAN:
			case VSI_DICOM_EXPLICIT_VR_LITTLE_ENDIAN:
			case VSI_DICOM_JPEG_BASELINE:
			case VSI_DICOM_RLE_LOSSLESS:
			case VSI_DICOM_JPEG_LOSSLESS:
				pszExt = L"dcm";
				break;
			case VSI_EXPORT_TYPE_PHYSIO:
				pszExt = L"physio.csv";
				break;
			default:
				_ASSERT(0);
				break;
			}
		}
		break;

	case VSI_EXPORT_IMAGE_FLAG_GRAPH:
		break;
	}

	return pszExt;
}

BOOL CVsiExportImage::IsAllLoopsOfModes(ULONG mode) throw(...)
{
	HRESULT hr;

	CVsiListImageIter iter;
	for (iter = m_listImage.begin(); iter != m_listImage.end(); ++iter)
	{
		CVsiImage &item = *iter;

		if (item.m_pImage != NULL)
		{
			CComVariant vMode;
			hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CComVariant vModeId;
			hr = m_pModeMgr->GetModePropertyFromInternalName(
				V_BSTR(&vMode),
				L"id",
				&vModeId);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			if (V_UI4(&vModeId) != mode)
			{
				return FALSE;
			}
		}
		else if (item.m_pMode != NULL)
		{
			CComQIPtr<IVsiPropertyBag> pProps(item.m_pMode);
			VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

			CComVariant vModeId;
			hr = pProps->ReadId(VSI_PROP_MODE_MODE_ID, &vModeId);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			if (V_UI4(&vModeId) != mode)
			{
				return FALSE;
			}
		}
	}

	return TRUE;
}

BOOL CVsiExportImage::IsAllLoopsOf3dModes() throw(...)
{
	HRESULT hr = S_OK;

	CVsiListImageIter iter;
	for (iter = m_listImage.begin(); iter != m_listImage.end(); ++iter)
	{
		CVsiImage &item = *iter;

		if (item.m_pImage != NULL)
		{
			CComVariant vMode;
			hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CComVariant v3d;
			hr = m_pModeMgr->GetModePropertyFromInternalName(
				V_BSTR(&vMode),
				L"3d",
				&v3d);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			if (VARIANT_FALSE == V_BOOL(&v3d))
			{
				return FALSE;
			}
		}
		else if (item.m_pMode != NULL)
		{
			CComQIPtr<IVsiPropertyBag> pProps(item.m_pMode);
			VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

			CComVariant vModeId;
			hr = pProps->ReadId(VSI_PROP_MODE_MODE_ID, &vModeId);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CComVariant v3d;
			hr = m_pModeMgr->GetModePropertyFromId(
				V_UI4(&vModeId),
				L"3d",
				&v3d);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			if (VARIANT_FALSE == V_BOOL(&v3d))
			{
				return FALSE;
			}
		}
	}

	return TRUE;
}

BOOL CVsiExportImage::IsAnyOf3dVolumeModes() throw(...)
{
	HRESULT hr = S_OK;

	CVsiListImageIter iter;
	for (iter = m_listImage.begin(); iter != m_listImage.end(); ++iter)
	{
		CVsiImage &item = *iter;

		if (item.m_pImage != NULL)
		{
			CComVariant vMode;
			hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CComVariant v3d;
			hr = m_pModeMgr->GetModePropertyFromInternalName(
				V_BSTR(&vMode),
				L"3d",
				&v3d);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			if (VARIANT_FALSE != V_BOOL(&v3d))
			{
				return TRUE;
			}
		}
		else if (item.m_pMode != NULL)
		{
			CComQIPtr<IVsiPropertyBag> pProps(item.m_pMode);
			VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

			CComVariant vModeId;
			hr = pProps->ReadId(VSI_PROP_MODE_MODE_ID, &vModeId);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CComVariant v3d;
			hr = m_pModeMgr->GetModePropertyFromId(
				V_UI4(&vModeId),
				L"3d",
				&v3d);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			if (VARIANT_FALSE != V_BOOL(&v3d))
			{
				return TRUE;
			}
		}
	}

	return FALSE;
}

BOOL CVsiExportImage::IsAnyOfModes(ULONG mode) throw(...)
{
	HRESULT hr = S_OK;

	CVsiListImageIter iter;
	for (iter = m_listImage.begin(); iter != m_listImage.end(); ++iter)
	{
		CVsiImage &item = *iter;

		if (item.m_pImage != NULL)
		{
			CComVariant vMode;
			hr = item.m_pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CComVariant vModeId;
			hr = m_pModeMgr->GetModePropertyFromInternalName(
				V_BSTR(&vMode),
				L"id",
				&vModeId);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			if (V_UI4(&vModeId) == mode)
			{
				return TRUE;
			}
		}
		else if (item.m_pMode != NULL)
		{
			CComQIPtr<IVsiPropertyBag> pProps(item.m_pMode);
			VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

			CComVariant vModeId;
			hr = pProps->ReadId(VSI_PROP_MODE_MODE_ID, &vModeId);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			if (V_UI4(&vModeId) == mode)
			{
				return TRUE;
			}
		}
	}

	return FALSE;
}

// Returns true if this RF type is supported by this image
BOOL CVsiExportImage::IsThisOfRf(IVsiImage *pImage, IVsiMode *pMode, VSI_RF_EXPORT_TYPE rfType, BOOL *pbEkv) throw(...)
{
	HRESULT hr = S_OK;
	BOOL bIsOfRf(FALSE);
	DWORD dwMode(VSI_SETTINGS_MODE_NULL);
	BOOL bIsOxyPaMode(FALSE);
	BOOL bIsCineStroed(FALSE);

	if (NULL != pbEkv)
		*pbEkv = FALSE;

	if (pImage != NULL)
	{
		CComVariant vRf(false);
		CComVariant vEkv(false);

		bIsCineStroed = TRUE; // This image has at least been cine stored

		// Property only set on cine store, thus we cannot export RF from a non
		// cine stored object (frame store ok for m and pw)
		// this is how we want it to work
		hr = pImage->GetProperty(VSI_PROP_IMAGE_RF_DATA, &vRf);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		if (hr == S_OK)
		{
			if (V_BOOL(&vRf) == VARIANT_TRUE)
				bIsOfRf = TRUE;
		}

		// If this is EKV then RF exports are not permitted
		hr = pImage->GetProperty(VSI_PROP_IMAGE_EKV, &vEkv);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		if (hr == S_OK)
		{
			if (V_BOOL(&vEkv) == VARIANT_TRUE)
			{
				bIsOfRf = FALSE;

				if (NULL != pbEkv)
					*pbEkv = TRUE;
			}
		}

		// Get mode
		{
			CComVariant vMode;
			hr = pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			if (hr == S_OK)
			{
				CComVariant vModeId;
				hr = m_pModeMgr->GetModePropertyFromInternalName(
					V_BSTR(&vMode),
					L"id",
					&vModeId);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				dwMode = V_UI4(&vModeId);
			}
		}

		// Check to see if this is OXY Pa mode - no IQ, only RAW
		if (bIsOfRf && (dwMode == VSI_SETTINGS_MODE_PA_MODE_3D || dwMode == VSI_SETTINGS_MODE_PA_MODE))
		{
			CComVariant vAcqType(0);

			hr = pImage->GetProperty(VSI_PROP_IMAGE_PA_ACQ_TYPE, &vAcqType);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			if (hr == S_OK)
			{
				if (V_UI4(&vAcqType) == VSI_PA_MODE_ACQUISITION_OXYHEMO)
					bIsOxyPaMode = TRUE;
			}
		}
	}
	else if (pMode != NULL)
	{
		// When we have a loaded loop we need to get the image from here
		bIsCineStroed = TRUE;

		CComQIPtr<IVsiPropertyBag> pProps(pMode);
		VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

		CComVariant vRoot;
		hr = pProps->ReadId(VSI_PROP_MODE_ROOT, &vRoot);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		bIsOfRf = VsiGetBooleanValue<BOOL>(
			V_BSTR(&vRoot), VSI_PDM_GROUP_MODE,
			VSI_PARAMETER_SETTINGS_RF_FEATURE, m_pPdm);

		// If EKV this is not RF
		BOOL bEkv = VsiGetBooleanValue<BOOL>(
			V_BSTR(&vRoot), VSI_PDM_GROUP_B_MODE,
			VSI_PARAMETER_EKV_LOOP, m_pPdm);

		if (bEkv)
		{
			bIsOfRf = FALSE;

			if (NULL != pbEkv)
				*pbEkv = TRUE;
		}

		CComVariant vModeId;
		hr = pProps->ReadId(VSI_PROP_MODE_MODE_ID, &vModeId);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		dwMode = V_I4(&vModeId);

		// Check to see if this is OXY Pa mode - no IQ, only RAW
		if (bIsOfRf && (dwMode == VSI_SETTINGS_MODE_PA_MODE_3D || dwMode == VSI_SETTINGS_MODE_PA_MODE))
		{
			CString strRoot;
			VsiModeUtils::GetModeRoot(pMode, strRoot);

			DWORD dwAcqModeCurrent = VsiGetDiscreteValue<DWORD>(
				strRoot, VSI_PDM_GROUP_PA_MODE,
				VSI_PARAMETER_PA_MODE_ACQUISITION_MODE, m_pPdm);
			bIsOxyPaMode = (dwAcqModeCurrent == VSI_PA_MODE_ACQUISITION_OXYHEMO);
		}
	}

	// Check to see if the type is valid for this mode
	if (bIsOfRf)
	{
		switch (rfType)
		{
			case VSI_RF_EXPORT_IQ:
				{
					// 3D Volume does not support IQ
					switch (dwMode)
					{
					case VSI_SETTINGS_MODE_3D:
					case VSI_SETTINGS_MODE_POWER_3D:
					case VSI_SETTINGS_MODE_COLOR_3D:
					case VSI_SETTINGS_MODE_CONTRAST_3D:
					case VSI_SETTINGS_MODE_PA_MODE_3D:
						bIsOfRf = FALSE;
						break;
					}

					// Only allow ADV Contrast IQ if we are in engineering mode
					BOOL bEngineerMode = VsiModeUtils::GetIsEngineerMode(m_pPdm);
					if (dwMode == VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED && !bEngineerMode)
					{
						bIsOfRf = FALSE;
					}

					// Only allow PA mode IQ if not in OXY/Hemo sub-mode
					if (bIsOxyPaMode)
					{
						bIsOfRf = FALSE;
					}
				}
				break;

			case VSI_RF_EXPORT_RF:
				{
					// Only the following modes support RF
					if (!(dwMode == VSI_SETTINGS_MODE_B_MODE ||
						dwMode == VSI_SETTINGS_MODE_M_MODE ||
						dwMode == VSI_SETTINGS_MODE_AM_MODE ||
						dwMode == VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL ||
						dwMode == VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED))
					{
						bIsOfRf = FALSE;
					}
				}
				break;

			case VSI_RF_EXPORT_RAW:
				{
					// Only the following modes support RAW
					if (!(dwMode == VSI_SETTINGS_MODE_B_MODE ||
						dwMode == VSI_SETTINGS_MODE_M_MODE ||
						dwMode == VSI_SETTINGS_MODE_AM_MODE ||
						dwMode == VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL ||
						dwMode == VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED ||
						dwMode == VSI_SETTINGS_MODE_COLOR_DOPPLER ||
						dwMode == VSI_SETTINGS_MODE_POWER_DOPPLER ||
						dwMode == VSI_SETTINGS_MODE_PA_MODE))
					{
						bIsOfRf = FALSE;
					}
				}
				break;
		}
	}
	else
	{
		// This image is not a true RF acquired image - see if raw is supported
		if (bIsCineStroed)
		{
			if (rfType == VSI_RF_EXPORT_RAW)
			{
				bIsOfRf = TRUE;

				// Only the following modes support RAW
				if (!(dwMode == VSI_SETTINGS_MODE_B_MODE ||
					dwMode == VSI_SETTINGS_MODE_M_MODE ||
					dwMode == VSI_SETTINGS_MODE_AM_MODE ||
					dwMode == VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL ||
					dwMode == VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED ||
					dwMode == VSI_SETTINGS_MODE_COLOR_DOPPLER ||
					dwMode == VSI_SETTINGS_MODE_POWER_DOPPLER ||
					dwMode == VSI_SETTINGS_MODE_PA_MODE))
				{
					bIsOfRf = FALSE;
				}
			}
		}
	}

	return bIsOfRf;
}

// Returns true if all images in the set are of the RF type
// For this to return true at least one image in the set must be of this type
// If the image is not cine saved, then this returns false
BOOL CVsiExportImage::IsAllOfRf(VSI_RF_EXPORT_TYPE rfType) throw(...)
{
	BOOL bResult(FALSE);

	CVsiListImageIter iter;
	for (iter = m_listImage.begin(); iter != m_listImage.end(); ++iter)
	{
		CVsiImage &item = *iter;

		BOOL bIsRf = IsThisOfRf(item.m_pImage, item.m_pMode, rfType, NULL);
		if (!bIsRf)
		{
			return FALSE;
		}
		else
		{
			// At least one is of this type
			bResult = TRUE;
		}
	}

	return bResult;
}

// Returns true if all images in the set are of the EKV type
BOOL CVsiExportImage::IsAllOfEKV() throw(...)
{
	BOOL bResult(FALSE);

	CVsiListImageIter iter;
	for (iter = m_listImage.begin(); iter != m_listImage.end(); ++iter)
	{
		CVsiImage &item = *iter;

		BOOL bEkv(FALSE);
		BOOL bIsRf = IsThisOfRf(item.m_pImage, item.m_pMode, VSI_RF_EXPORT_RAW, &bEkv);
		if (!bEkv || !bIsRf)
		{
			return FALSE;
		}
		else if (bEkv && bIsRf)
		{
			// At least one is of this type
			bResult = TRUE;
		}
	}

	return bResult;
}

//// If we remove the VXML we have a valid file extension
///
/// <summary>
///		Estimate the size of the RF data for this object. If it is a frame,
///		open the files and try to read the first frame header
///
/// </summary>
///
/// <returns>Total number of bytes</returns>
double CVsiExportImage::EstimateRFSize(IVsiImage *pImage, IVsiMode *pMode, BOOL bFrame, VSI_RF_EXPORT_TYPE rfType)
{
	HRESULT hr = S_OK;
	CString strPath;
	double dbSize(0);

	try
	{
		if (pImage != NULL)
		{
			CComHeapPtr<OLECHAR> pszStudyFile;
			hr = pImage->GetDataPath(&pszStudyFile);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			strPath = pszStudyFile;
			strPath.GetBufferSetLength(strPath.GetLength() - 4);
		}
		else if (pMode != NULL)
		{
			// When we have a loaded loop we need to get the image from here

			CComQIPtr<IVsiPropertyBag> pProps(pMode);
			VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

			CComVariant vImage;
			hr = pProps->ReadId(VSI_PROP_MODE_IMAGE, &vImage);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (hr == S_OK)
			{
				CComQIPtr<IVsiImage> pImage(V_UNKNOWN(&vImage));
				VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

				CComHeapPtr<OLECHAR> pszStudyFile;
				hr = pImage->GetDataPath(&pszStudyFile);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				strPath = pszStudyFile;
				strPath.GetBufferSetLength(strPath.GetLength() - 4);
			}
		}

		// Check to see if this is valid
		if (strPath.GetLength() > 0)
		{
			LARGE_INTEGER liSize;

			CString pstrRfFiles[60];
			int iFilesOut = VsiModeUtils::GetRfFileNames(
				strPath,
				pstrRfFiles,
				_countof(pstrRfFiles),
				-1); // all types

			if (bFrame)
			{
				for (int i=0; i<iFilesOut; ++i)
				{
					// Make sure this is not a physio data file
					BOOL bIsPhysio = pstrRfFiles[i].Find(VSI_RF_FILE_EXTENSION_PHYSIO) > 0;
					BOOL bIsXML = pstrRfFiles[i].Find(VSI_RF_FILE_EXTENSION_XML) > 0;

					if (VsiGetFileExists(pstrRfFiles[i]) && !bIsPhysio && !bIsXML)
					{
						HANDLE hFile = CreateFile(pstrRfFiles[i], GENERIC_READ, FILE_SHARE_READ, NULL,
							OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
						if (INVALID_HANDLE_VALUE != hFile)
						{
							VSI_RF_FILE_HEADER fileHeader;
							ZeroMemory(&fileHeader, sizeof(VSI_RF_FILE_HEADER));

							DWORD dwRead;
							ReadFile(hFile, &fileHeader, sizeof(VSI_RF_FILE_HEADER), &dwRead, NULL);
							VSI_WIN32_VERIFY(sizeof(VSI_RF_FILE_HEADER) == (int)dwRead, VSI_LOG_ERROR, L"Failed to write data buffer");

							// Size in the first frame
							if (fileHeader.dwVersion == VSI_DOWNLOAD_RF_VERSION_3 ||
								fileHeader.dwVersion == VSI_DOWNLOAD_RF_VERSION_2)
							{
								VSI_RF_FRAME_HEADER frameHeader;
								ZeroMemory(&frameHeader, sizeof(VSI_RF_FRAME_HEADER));

								ReadFile(hFile, &frameHeader, sizeof(VSI_RF_FRAME_HEADER), &dwRead, NULL);
								VSI_WIN32_VERIFY(sizeof(VSI_RF_FRAME_HEADER) == (int)dwRead, VSI_LOG_ERROR, L"Failed to read data buffer");

								dbSize += sizeof(VSI_RF_FRAME_HEADER) + sizeof(VSI_RF_FILE_HEADER);
								if (rfType == VSI_RF_EXPORT_IQ)
								{
									dbSize += frameHeader.dwPacketSize;
								}
								else if (rfType == VSI_RF_EXPORT_RF)
								{
									dbSize += frameHeader.dwPacketSize * 16;
								}
								else if (rfType == VSI_RF_EXPORT_RAW)
								{
									// This is approximately right - at least it is not smaller
									dbSize += frameHeader.dwPacketSize;
								}
							}

							CloseHandle(hFile);
							hFile = INVALID_HANDLE_VALUE;
						}
					}
				}
			}
			else
			{
				// get size of entire file - assume loop is whole file, may be much less, don't worry about details
				for (int i=0; i<iFilesOut; ++i)
				{
					// Make sure this is not a physio data file
					BOOL bIsPhysio = pstrRfFiles[i].Find(VSI_RF_FILE_EXTENSION_PHYSIO) > 0;
					BOOL bIsXML = pstrRfFiles[i].Find(VSI_RF_FILE_EXTENSION_XML) > 0;
					BOOL bIsDopplerIQ = pstrRfFiles[i].Find(VSI_RF_FILE_EXTENSION_PWMODE) > 0;

					if (VsiGetFileExists(pstrRfFiles[i]) && !bIsPhysio && !bIsXML)
					{
						VsiGetFileSize(pstrRfFiles[i], &liSize);
						if (rfType == VSI_RF_EXPORT_IQ)
						{
							dbSize += liSize.QuadPart;
						}
						else if (rfType == VSI_RF_EXPORT_RF)
						{
							dbSize += liSize.QuadPart * 16;
						}
						else if (rfType == VSI_RF_EXPORT_RAW)
						{
							// This is approximately right - at least it is not smaller
							dbSize += liSize.QuadPart;
						}

						if (bIsDopplerIQ)
						{
							CString dopplerPhysio(strPath);
							dopplerPhysio += L"pimg";
							if (VsiGetFileExists(dopplerPhysio))
							{
								VsiGetFileSize(dopplerPhysio, &liSize);
								dbSize += liSize.QuadPart;
							}
						}
					}
				}				
			}

			// Add 4 k for RFXML (approximate)
			dbSize += 4 * 1024;
		}
		else //we are just going to fake the estimate at this point
		{
			CComPtr<IVsiModeData> pModeData;

			hr = pMode->GetModeData(&pModeData);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			double dbLength;
			hr = pModeData->GetLength(&dbLength);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Check is frame based
			CComQIPtr<IVsiModeFrameData> pFrameData(pModeData);
			if (NULL == pFrameData)
			{
				// The export size about 0.5M per 5 seconds
				dbLength = dbLength / 10000.0;
			}

			// Default used 1M per frame
			dbSize = 1048576 * dbLength;
			if (rfType == VSI_RF_EXPORT_RF)
			{
				dbSize = dbSize * 32;
			}
			else if (rfType == VSI_RF_EXPORT_IQ)
			{
				dbSize = dbSize * 2;
			}
		}
	}
	VSI_CATCH(hr);

	// Add 10% extra to size
	return (dbSize * 1.1);
}

/// <summary>
///		Try to estimate how much total disk space will be required to store all
///		of the selected images based on the currently selected file type. The
///		value is in bytes.
///
///		NOTE: the numbers used for this estimate were determine empirically on
///		one system on a sunny day. Your millage may vary.
///
/// </summary>
///
/// <returns>Total number of bytes</returns>
double CVsiExportImage::EstimateSize()
{
	HRESULT hr = S_OK;
	double dbEstimatedSize(0);
	double dbDurationMS = 0.0;
	double dbCompressionFactor = 1.0;
	double dbNumFrames = 1.0;
	DWORD dwMode = VSI_SETTINGS_MODE_NULL;
	CComVariant vLength;
	double dbLength = 0.0;
	int iCx = 0, iCy = 0;

	try
	{
		CWindow wndExportType(GetTypeControl());

		switch (m_dwExportTypeMain)
		{
		case VSI_EXPORT_IMAGE_FLAG_CINE:
		case VSI_EXPORT_IMAGE_FLAG_IMAGE:
		case VSI_EXPORT_IMAGE_FLAG_DICOM:
			{
				int iIndex = (int)wndExportType.SendMessage(CB_GETCURSEL, NULL, NULL);
				VSI_VERIFY(CB_ERR != iIndex, VSI_LOG_ERROR, NULL);

				DWORD dwType = (DWORD)wndExportType.SendMessage(CB_GETITEMDATA, iIndex, NULL);

				if (VSI_DICOM_RLE_LOSSLESS == dwType)
					dbCompressionFactor = 12.0 / 15.0;
				else if (VSI_DICOM_JPEG_BASELINE == dwType)
					dbCompressionFactor = 1.0 / 5.0;

				CVsiListImageIter iter;
				for (iter = m_listImage.begin(); iter != m_listImage.end(); ++iter)
				{
					CVsiImage &item = *iter;
					IVsiImage *pImage = item.m_pImage;
					IVsiMode *pMode = item.m_pMode;

					// Determine if this image is an RF image
					BOOL bIsRfIQ = IsThisOfRf(pImage, pMode, VSI_RF_EXPORT_IQ, NULL);
					BOOL bIsRfRf = IsThisOfRf(pImage, pMode, VSI_RF_EXPORT_RF, NULL);
					BOOL bIsRfRaw = IsThisOfRf(pImage, pMode, VSI_RF_EXPORT_RAW, NULL);

					if (NULL != pMode && NULL == pImage)
					{
						CString strRoot, strGroup;
						VsiModeUtils::GetModeRootGroup(pMode, strRoot, strGroup);

						CComQIPtr<IVsiPropertyBag> pModeProps(item.m_pMode);
						VSI_CHECK_INTERFACE(pModeProps, VSI_LOG_ERROR, NULL);

						CComVariant vModeId;
						hr = pModeProps->ReadId(VSI_PROP_MODE_MODE_ID, &vModeId);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						dwMode = V_I4(&vModeId);

						CComPtr<IVsiModeView> pModeView;
						hr = item.m_pMode->GetModeView(&pModeView);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComQIPtr<IVsiModeExport> pModeExport(pModeView);
						VSI_CHECK_INTERFACE(pModeExport, VSI_LOG_ERROR, NULL);

						hr = pModeExport->CaptureImage(
							NULL, NULL,
							&iCx, &iCy,
							24,
							VSI_MODE_CAPTURE_HEADER | VSI_MODE_CAPTURE_DICOM_IMAGE |
								VSI_MODE_CAPTURE_GET_SIZE_ONLY, NULL, NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IVsiModeData> pModeData;
						hr = item.m_pMode->GetModeData(&pModeData);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						hr = pModeData->GetLength(&dbLength);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComVariant vFrame;
						hr = m_pModeMgr->GetModePropertyFromId(
							dwMode,
							L"frameBased",
							&vFrame);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						CComVariant v3d;
						hr = m_pModeMgr->GetModePropertyFromId(
							dwMode,
							L"3d",
							&v3d);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						if (VARIANT_FALSE != V_BOOL(&vFrame))
						{
							dbNumFrames = (double)dbLength;
						}
						else if (VARIANT_FALSE != V_BOOL(&v3d))
						{
							CComQIPtr<IVsi3dModeData> p3dData(pModeData);
							VSI_CHECK_INTERFACE(p3dData, VSI_LOG_ERROR, NULL);

							int iCx, iCy, iCz;
							p3dData->GetCubeSize(&iCx, &iCy, &iCz);

							if (VSI_IMAGE_TIFF_LOOP == dwType ||
								VSI_DICOM_IMPLICIT_VR_LITTLE_ENDIAN == dwType ||
								VSI_DICOM_EXPLICIT_VR_LITTLE_ENDIAN == dwType ||
								VSI_DICOM_JPEG_BASELINE == dwType ||
								VSI_DICOM_RLE_LOSSLESS == dwType ||
								VSI_DICOM_JPEG_LOSSLESS == dwType)
							{
								dbEstimatedSize = iCx * iCy * iCz;
							}
							else
							{
								dbEstimatedSize = iCx * iCy;
							}
						}
					}
					else
					{
						VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_NONE);
						pImage->GetErrorCode(&dwErrorCode);
						if (VSI_IMAGE_ERROR_NONE != dwErrorCode)
							continue;

						#pragma MESSAGE("Get image size - iCx and iCy")
						iCx = 640;
						iCy = 640;

						// Get image duration
						CComVariant vMode;
						hr = pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComVariant vModeId;
						hr = m_pModeMgr->GetModePropertyFromInternalName(
							V_BSTR(&vMode),
							L"id",
							&vModeId);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						dwMode = V_UI4(&vModeId);
						if (VSI_SETTINGS_MODE_NULL == dwMode)
							continue;

						// Duration
						hr = pImage->GetProperty(VSI_PROP_IMAGE_LENGTH, &vLength);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						dbLength = V_R8(&vLength);

						CComVariant vFrame;
						hr = m_pModeMgr->GetModePropertyFromId(
							dwMode,
							L"frameBased",
							&vFrame);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						CComVariant v3d;
						hr = m_pModeMgr->GetModePropertyFromId(
							dwMode,
							L"3d",
							&v3d);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						if (VARIANT_FALSE == V_BOOL(&vFrame))
						{
							dbLength *= 1000.0;
						}
						else if (VARIANT_FALSE != V_BOOL(&v3d) ||
							VSI_SETTINGS_MODE_PA_MODE == dwMode)
						{
							// use minimal step size to calculate number of frames
							dbNumFrames = V_R8(&vLength) / 0.032;

							if (VSI_IMAGE_TIFF_LOOP == dwType ||
								VSI_DICOM_IMPLICIT_VR_LITTLE_ENDIAN == dwType ||
								VSI_DICOM_EXPLICIT_VR_LITTLE_ENDIAN == dwType ||
								VSI_DICOM_JPEG_BASELINE == dwType ||
								VSI_DICOM_RLE_LOSSLESS == dwType ||
								VSI_DICOM_JPEG_LOSSLESS == dwType)
							{
								dbEstimatedSize = iCx * iCy * dbNumFrames;
							}
							else
							{
								dbEstimatedSize = iCx * iCy;
							}
						}
						else
						{
							dbNumFrames = V_R8(&vLength);
						}
					}

					dbDurationMS = dbLength;

					// Duration can be a little short depending on rounding
					dbDurationMS += 25;

					// header
					switch (dwType)
					{
						case VSI_IMAGE_TIFF_LOOP:
							{
								if (NULL != pImage)
								{
									LARGE_INTEGER liFile = {0};
									CString strFile;
									CComVariant vName;

									CComHeapPtr<OLECHAR> pszFile;
									hr = pImage->GetDataPath(&pszFile);

									WCHAR szRawFile[MAX_PATH];
									lstrcpyn(szRawFile, pszFile, _countof(szRawFile));

									LPWSTR pszDot = wcsrchr(szRawFile, L'.');
									VSI_VERIFY(NULL != pszDot, VSI_LOG_ERROR, NULL);

									lstrcpy(pszDot + 1, VSI_FILE_EXT_IMAGE_RAW_3D);

									VsiGetFileSize(szRawFile, &liFile);
									dbEstimatedSize += liFile.LowPart;
								}
								// If not, it was calculated before (by the memory size)
							}
						break;

						case VSI_IMAGE_TIFF_FULL:
							dbEstimatedSize += 2314240;				// checked
							break;

						case VSI_IMAGE_BMP_FULL:
							dbEstimatedSize += 3084288;				// checked
							break;

						case VSI_IMAGE_TIFF_IMAGE:
							{
								switch (dwMode)
								{
								case VSI_SETTINGS_MODE_COLOR_DOPPLER:
								case VSI_SETTINGS_MODE_POWER_DOPPLER:
								case VSI_SETTINGS_MODE_TISSUE_DOPPLER:
								case VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL:
								case VSI_SETTINGS_MODE_B_MODE:
								case VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED:
								case VSI_SETTINGS_MODE_PA_MODE:
									dbEstimatedSize += 1466368;		// updated
									break;

								case VSI_SETTINGS_MODE_M_MODE:
								case VSI_SETTINGS_MODE_AM_MODE:
									dbEstimatedSize += 1368180;		// updated
									break;

								case VSI_SETTINGS_MODE_PW_DOPPLER:
								case VSI_SETTINGS_MODE_PW_TISSUE_DOPPLER:
									dbEstimatedSize += 1392180;		// updated
									break;

								case VSI_SETTINGS_MODE_3D:
								case VSI_SETTINGS_MODE_POWER_3D:
								case VSI_SETTINGS_MODE_COLOR_3D:
								case VSI_SETTINGS_MODE_CONTRAST_3D:
								case VSI_SETTINGS_MODE_ADV_CONTRAST_3D:
								case VSI_SETTINGS_MODE_PA_MODE_3D:
									dbEstimatedSize += 1785856;		// checked
									break;

								default:
									_ASSERT(false);
									break;
								}
							}
							break;

						case VSI_DICOM_IMPLICIT_VR_LITTLE_ENDIAN:
						case VSI_DICOM_EXPLICIT_VR_LITTLE_ENDIAN:
						case VSI_DICOM_JPEG_BASELINE:
						case VSI_DICOM_RLE_LOSSLESS:
						case VSI_DICOM_JPEG_LOSSLESS:
							{
								if (!VsiModeUtils::GetIsEngineerMode(m_pPdm))
								{
									// We only support TomTec DICOM in Engineering Mode
									if (VSI_SETTINGS_MODE_3D == dwMode ||
										VSI_SETTINGS_MODE_POWER_3D == dwMode ||
										VSI_SETTINGS_MODE_COLOR_3D == dwMode ||
										VSI_SETTINGS_MODE_CONTRAST_3D == dwMode ||
										VSI_SETTINGS_MODE_ADV_CONTRAST_3D == dwMode ||
										VSI_SETTINGS_MODE_PA_MODE_3D == dwMode)
									{
										break;
									}
								}

								//if (!IsDICOMCompressedFileType())
								//	m_bFixedSize = TRUE;

								if (VSI_SETTINGS_MODE_M_MODE == dwMode ||
									VSI_SETTINGS_MODE_AM_MODE == dwMode ||
									VSI_SETTINGS_MODE_PW_TISSUE_DOPPLER == dwMode ||
									VSI_SETTINGS_MODE_PW_DOPPLER == dwMode)
								{
									DWORD dwPlaybackRate = VsiGetRangeValue<DWORD>(
										VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
										VSI_PARAMETER_SYS_DEFAULT_EXPORT_DICOM_PLAYBACK_RATE, m_pPdm);
									dbNumFrames = (int)ceil(dbDurationMS * dwPlaybackRate / 1000.0);
								}

								dbEstimatedSize += 5120; // header
								switch (dwMode)
								{
								case VSI_SETTINGS_MODE_COLOR_DOPPLER:
								case VSI_SETTINGS_MODE_POWER_DOPPLER:
								case VSI_SETTINGS_MODE_TISSUE_DOPPLER:
								case VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL:
								case VSI_SETTINGS_MODE_B_MODE:
								case VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED:
								case VSI_SETTINGS_MODE_PA_MODE:
								case VSI_SETTINGS_MODE_M_MODE:
								case VSI_SETTINGS_MODE_AM_MODE:
								case VSI_SETTINGS_MODE_PW_DOPPLER:
								case VSI_SETTINGS_MODE_PW_TISSUE_DOPPLER:
									dbEstimatedSize += dbNumFrames * (iCx * iCy * 3 + 2097152) * dbCompressionFactor;
									break;

								default:
									break;
								}
							}
							break;

						case VSI_IMAGE_BMP_IMAGE:
							{
								m_bFixedSize = TRUE;

								switch (dwMode)
								{
								case VSI_SETTINGS_MODE_COLOR_DOPPLER:
								case VSI_SETTINGS_MODE_POWER_DOPPLER:
								case VSI_SETTINGS_MODE_TISSUE_DOPPLER:
								case VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL:
								case VSI_SETTINGS_MODE_B_MODE:
								case VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED:
								case VSI_SETTINGS_MODE_PA_MODE:
									dbEstimatedSize += 1953792;		// updated
									break;

								case VSI_SETTINGS_MODE_M_MODE:
								case VSI_SETTINGS_MODE_AM_MODE:
									dbEstimatedSize += 1826816;		// updated
									break;

								case VSI_SETTINGS_MODE_PW_DOPPLER:
								case VSI_SETTINGS_MODE_PW_TISSUE_DOPPLER:
									dbEstimatedSize += 1859584;		// updated
									break;

								case VSI_SETTINGS_MODE_3D:
								case VSI_SETTINGS_MODE_POWER_3D:
								case VSI_SETTINGS_MODE_COLOR_3D:
								case VSI_SETTINGS_MODE_CONTRAST_3D:
								case VSI_SETTINGS_MODE_ADV_CONTRAST_3D:
								case VSI_SETTINGS_MODE_PA_MODE_3D:
									dbEstimatedSize += 2384872;		// checked
									break;

								default:
									_ASSERT(false);
									break;
								}
							}
							break;

						case VSI_IMAGE_WAV:
							{
								// For all intents and purposes this is exact, or very close
								m_bFixedSize = TRUE;

								switch (dwMode)
								{
								case VSI_SETTINGS_MODE_B_MODE:
								case VSI_SETTINGS_MODE_COLOR_DOPPLER:
								case VSI_SETTINGS_MODE_POWER_DOPPLER:
								case VSI_SETTINGS_MODE_TISSUE_DOPPLER:
								case VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL:
								case VSI_SETTINGS_MODE_M_MODE:
								case VSI_SETTINGS_MODE_AM_MODE:
								case VSI_SETTINGS_MODE_3D:
								case VSI_SETTINGS_MODE_POWER_3D:
								case VSI_SETTINGS_MODE_COLOR_3D:
								case VSI_SETTINGS_MODE_CONTRAST_3D:
								case VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED:
								case VSI_SETTINGS_MODE_ADV_CONTRAST_3D:
								case VSI_SETTINGS_MODE_PA_MODE:
									// this image can't be exported as a WAV file
									break;
								case VSI_SETTINGS_MODE_PW_DOPPLER:
								case VSI_SETTINGS_MODE_PW_TISSUE_DOPPLER:
									dbEstimatedSize += dbDurationMS * 192.0 + 44;
									break;

								// other modes cannot be exported as wave and should be caught
								default:
									_ASSERT(false);
									break;
								}
							}
							break;

						case VSI_IMAGE_IQ_DATA_FRAME:
						{
							if (bIsRfIQ)
							{
								dbEstimatedSize += EstimateRFSize(pImage, pMode, TRUE, VSI_RF_EXPORT_IQ);
							}
						}
						break;
						case VSI_IMAGE_RF_DATA_FRAME:
						{
							if (bIsRfRf)
							{
								dbEstimatedSize += EstimateRFSize(pImage, pMode, TRUE, VSI_RF_EXPORT_RF);
							}
						}
						break;
						case VSI_IMAGE_RAW_DATA_FRAME:
						{
							if (bIsRfRaw)
							{
								dbEstimatedSize += EstimateRFSize(pImage, pMode, TRUE, VSI_RF_EXPORT_RAW);
							}
						}
						break;

						case VSI_IMAGE_IQ_DATA_LOOP:
						{
							if (bIsRfIQ)
							{
								dbEstimatedSize += EstimateRFSize(pImage, pMode, FALSE, VSI_RF_EXPORT_IQ);
							}
						}
						break;
						case VSI_IMAGE_RF_DATA_LOOP:
						{
							if (bIsRfRf)
							{
								dbEstimatedSize += EstimateRFSize(pImage, pMode, FALSE, VSI_RF_EXPORT_RF);
							}
						}
						break;
						case VSI_IMAGE_RAW_DATA_LOOP:
						{
							if (bIsRfRaw)
							{
								dbEstimatedSize += EstimateRFSize(pImage, pMode, FALSE, VSI_RF_EXPORT_RAW);
							}
						}
						break;

						case VSI_IMAGE_AVI_CUSTOM:
						{
							DWORD dwFcc = m_pSubPanelCine->m_pdwFccHandlerType[iIndex];

							switch(dwFcc)
							{
								case VSI_4CC_ANIMGIF:
								{
									// Estimate number of approx frames at 30 fps
									int iFrames = (int)(dbDurationMS / 33.0);

									switch (dwMode)
									{
										case VSI_SETTINGS_MODE_M_MODE:
										case VSI_SETTINGS_MODE_AM_MODE:
										case VSI_SETTINGS_MODE_PW_DOPPLER:
										case VSI_SETTINGS_MODE_PW_TISSUE_DOPPLER:
											dbEstimatedSize += iFrames * 5300 + 1024;
											break;

										case VSI_SETTINGS_MODE_3D:
										case VSI_SETTINGS_MODE_POWER_3D:
										case VSI_SETTINGS_MODE_COLOR_3D:
										case VSI_SETTINGS_MODE_CONTRAST_3D:
										case VSI_SETTINGS_MODE_ADV_CONTRAST_3D:
										case VSI_SETTINGS_MODE_PA_MODE_3D:
											// this image can't be exported as an Uncompressed AVI file
											break;

										case VSI_SETTINGS_MODE_B_MODE:
										case VSI_SETTINGS_MODE_COLOR_DOPPLER:
										case VSI_SETTINGS_MODE_POWER_DOPPLER:
										case VSI_SETTINGS_MODE_TISSUE_DOPPLER:
										case VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL:
										case VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED:
										case VSI_SETTINGS_MODE_PA_MODE:
											// this image can be exported as an Uncompressed AVI
											// file only if it is not a single-frame image
											// Most images are not this large, however a 300 frame contrast loop was (large FOV)
											if (dbDurationMS != 1.0)
												dbEstimatedSize += iFrames * 150000 + 1024;
											break;

										default:
											_ASSERT(false);
											break;
									}
									break;
								}
								break;

								case VSI_4CC_UNCOMP:
								{
									switch (dwMode)
									{
										case VSI_SETTINGS_MODE_M_MODE:
										case VSI_SETTINGS_MODE_AM_MODE:
											dbEstimatedSize += dbDurationMS * 54361.81 + 17203.2;
											break;

										case VSI_SETTINGS_MODE_PW_DOPPLER:
										case VSI_SETTINGS_MODE_PW_TISSUE_DOPPLER:
											dbEstimatedSize += dbDurationMS * 59727.13 + 17203.2;
											break;

										case VSI_SETTINGS_MODE_3D:
										case VSI_SETTINGS_MODE_POWER_3D:
										case VSI_SETTINGS_MODE_COLOR_3D:
										case VSI_SETTINGS_MODE_CONTRAST_3D:
										case VSI_SETTINGS_MODE_ADV_CONTRAST_3D:
										case VSI_SETTINGS_MODE_PA_MODE_3D:
											// this image can't be exported as an Uncompressed AVI file
											break;

										case VSI_SETTINGS_MODE_B_MODE:
										case VSI_SETTINGS_MODE_COLOR_DOPPLER:
										case VSI_SETTINGS_MODE_POWER_DOPPLER:
										case VSI_SETTINGS_MODE_TISSUE_DOPPLER:
										case VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL:
										case VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED:
										case VSI_SETTINGS_MODE_PA_MODE:
											// this image can be exported as an Uncompressed AVI
											// file only if it is not a single-frame image
											if (dbDurationMS != 1.0)
												dbEstimatedSize += (dbDurationMS * 1356595.2) + 17203.2;
											break;

										default:
											_ASSERT(false);
											break;
									}
									break;
								}
								break;

								case VSI_4CC_COMP_MSVIDEO1:
								{
									switch (dwMode)
									{
										case VSI_SETTINGS_MODE_B_MODE:
										case VSI_SETTINGS_MODE_COLOR_DOPPLER:
										case VSI_SETTINGS_MODE_POWER_DOPPLER:
										case VSI_SETTINGS_MODE_TISSUE_DOPPLER:
										case VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL:
										case VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED:
										case VSI_SETTINGS_MODE_PA_MODE:
											// this image can be exported as an MSVideo1 AVI
											// file only if it is not a single-frame image
											if (dbDurationMS != 1.0)
												dbEstimatedSize += (dbDurationMS * 199884.8) + 7372.8;
											break;

										case VSI_SETTINGS_MODE_M_MODE:
										case VSI_SETTINGS_MODE_AM_MODE:
											dbEstimatedSize += dbDurationMS * 6533.2;
											break;

										case VSI_SETTINGS_MODE_PW_DOPPLER:
										case VSI_SETTINGS_MODE_PW_TISSUE_DOPPLER:
											dbEstimatedSize += dbDurationMS * 3918;
											break;

										case VSI_SETTINGS_MODE_3D:
										case VSI_SETTINGS_MODE_POWER_3D:
										case VSI_SETTINGS_MODE_COLOR_3D:
										case VSI_SETTINGS_MODE_CONTRAST_3D:
										case VSI_SETTINGS_MODE_ADV_CONTRAST_3D:
										case VSI_SETTINGS_MODE_PA_MODE_3D:
											// this image can't be exported as an MSVideo1 AVI file
											break;

										default:
											_ASSERT(false);
											break;
									}
									break;
								}
								break;

								case VSI_4CC_COMPCINEPAK:
								default:
								{
									switch (dwMode)
									{
										case VSI_SETTINGS_MODE_B_MODE:
										case VSI_SETTINGS_MODE_COLOR_DOPPLER:
										case VSI_SETTINGS_MODE_POWER_DOPPLER:
										case VSI_SETTINGS_MODE_TISSUE_DOPPLER:
										case VSI_SETTINGS_MODE_CONTRAST_MODE_FUNDAMENTAL:
										case VSI_SETTINGS_MODE_CONTRAST_MODE_ADVANCED:
										case VSI_SETTINGS_MODE_PA_MODE:
											// this image can be exported as a Cinepak AVI
											// file only if it is not a single-frame image
											if (dbDurationMS != 1.0)
												dbEstimatedSize += (dbDurationMS * 87654.4) + 18022.4;
											break;

										case VSI_SETTINGS_MODE_M_MODE:
										case VSI_SETTINGS_MODE_AM_MODE:
											dbEstimatedSize += dbDurationMS * 2225.5;
											break;

										case VSI_SETTINGS_MODE_PW_DOPPLER:
										case VSI_SETTINGS_MODE_PW_TISSUE_DOPPLER:
											dbEstimatedSize += dbDurationMS * 2082.3;
											break;

										case VSI_SETTINGS_MODE_3D:
										case VSI_SETTINGS_MODE_POWER_3D:
										case VSI_SETTINGS_MODE_COLOR_3D:
										case VSI_SETTINGS_MODE_CONTRAST_3D:
										case VSI_SETTINGS_MODE_ADV_CONTRAST_3D:
										case VSI_SETTINGS_MODE_PA_MODE_3D:
											// this image can't be exported as a Cinepak AVI file
											break;

										default:
											_ASSERT(false);
											break;
									}
									break;
								}
								break;
							}
							break;
						}
						break;

						case VSI_COMBINED_REPORT:
						{
							dbEstimatedSize += 1000;
						}
						break;

						case VSI_EXPORT_TABLE:
						{
							dbEstimatedSize += 1000;
						}
						break;

						default:
							_ASSERT(0);
							break;
					}
				}
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_REPORT:
			{
				dbEstimatedSize = 1024;
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_TABLE:
			{
				dbEstimatedSize = 1024;
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_GRAPH:
			{
				dbEstimatedSize = 0;
				if (m_pSubPanelGraph->IsDlgButtonChecked(IDC_IE_EXPORT_CSV))
				{
					dbEstimatedSize += 10240;				// checked
				}
				if (m_pSubPanelGraph->IsDlgButtonChecked(IDC_IE_EXPORT_BMP))
				{
					dbEstimatedSize += 3084288;				// checked
				}
				if (m_pSubPanelGraph->IsDlgButtonChecked(IDC_IE_EXPORT_TIFF))
				{
					dbEstimatedSize += 2314240;				// checked
				}
			}
			break;

		case VSI_EXPORT_IMAGE_FLAG_PHYSIO:
			{
				CVsiListImageIter iter;
				for (iter = m_listImage.begin(); iter != m_listImage.end(); ++iter)
				{
					CVsiImage &item = *iter;
					IVsiImage *pImage = item.m_pImage;
					IVsiMode *pMode = item.m_pMode;

					if (NULL != pMode)
					{
						CString strRoot, strGroup;
						VsiModeUtils::GetModeRootGroup(pMode, strRoot, strGroup);

						CComPtr<IVsiModeData> pModeData;
						hr = item.m_pMode->GetModeData(&pModeData);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComPtr<IVsiPhysiologicalRawData> pPhysioRawData;
						hr = pModeData->GetPhysiologicalData(__uuidof(IVsiPhysiologicalRawData), (IUnknown **)&pPhysioRawData);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting raw data");

						VSI_PHYSIO_DATA_SET dataSet;
						pPhysioRawData->GetDataSet(&dataSet);

						dbEstimatedSize += (dataSet.m_dwFrames * dataSet.m_dwFrameSamples * 64);
					}
					else if (NULL != pImage)
					{
						VSI_IMAGE_ERROR dwErrorCode(VSI_IMAGE_ERROR_NONE);
						pImage->GetErrorCode(&dwErrorCode);
						if (VSI_IMAGE_ERROR_NONE != dwErrorCode)
							continue;

						// Get image duration
						CComVariant vMode;
						hr = pImage->GetProperty(VSI_PROP_IMAGE_MODE, &vMode);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComVariant vModeId;
						hr = m_pModeMgr->GetModePropertyFromInternalName(
							V_BSTR(&vMode),
							L"id",
							&vModeId);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						dwMode = V_UI4(&vModeId);
						if (VSI_SETTINGS_MODE_NULL == dwMode)
							continue;

						CComHeapPtr<OLECHAR> pszFile;
						hr = pImage->GetDataPath(&pszFile);

						WCHAR szPhysioFile[MAX_PATH];
						errno_t err = wcscpy_s(szPhysioFile, pszFile);

						VsiPathRemoveExtension(szPhysioFile);
						err = wcscat_s(szPhysioFile, L".");
						VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");

						err = wcscat_s(szPhysioFile, VSI_FILE_EXT_IMAGE_RAW_PHYSIO);
						VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");

						LARGE_INTEGER liSize;
						VsiGetFileSize(szPhysioFile, &liSize);
						dbEstimatedSize += (liSize.LowPart * 5);
					}
				}

			}
			break;

		default:
			_ASSERT(0);
		}
	}
	VSI_CATCH(hr);

	// Round it to KB / MB
	if (dbEstimatedSize >= 1024)
	{
		if (dbEstimatedSize < (1024 * 1024))
		{
			double dbSizeMax = (double)(dbEstimatedSize) / 1024.0;
			if (dbSizeMax < 1.0)
			{
				dbSizeMax = 1.0;
			}

			dbSizeMax = (int)(dbSizeMax + 1.0);

			dbEstimatedSize = dbSizeMax * 1024; // Round m_iRequiredSize to KB
		}
		else if (dbEstimatedSize < (1024 * 1024 * 1024))
		{
			double dbSizeMax = (double)(dbEstimatedSize) / (1024.0 * 1024.0);
			if (dbSizeMax < 1.0)
			{
				dbSizeMax = 1.0;
			}

			dbSizeMax = (int)(dbSizeMax + 1.0);

			dbEstimatedSize = dbSizeMax * 1024 * 1024; // Round m_iRequiredSize to MB
		}
	}

	return dbEstimatedSize;
}

HRESULT CVsiExportImage::CreateMsmntReport(IVsiApp *pApp, IXMLDOMDocument **ppXmlMsmntDoc, CVsiListImage &imageList)
{
	HRESULT hr = S_OK;
	CComPtr<IVsiAnalysisSet> pAnalysisSet;
	CComPtr<IXMLDOMDocument> pXmlDoc;

	try
	{
		VSI_CHECK_MEM(ppXmlMsmntDoc, VSI_LOG_ERROR, NULL);

		// Create XML DOM document with Measurement Report
		hr = VsiCreateDOMDocument(&pXmlDoc);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating DOMDocument");

		CComPtr<IXMLDOMElement> pAnalysisResults;

		CComPtr<IXMLDOMProcessingInstruction> pPI;
		hr = pXmlDoc->createProcessingInstruction(CComBSTR(L"xml"), CComBSTR(L"version='1.0'"), &pPI);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pXmlDoc->appendChild(pPI, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Create and attach the root element
		hr = pXmlDoc->createElement(CComBSTR(VSI_MSMNT_XML_ELE_ANAL_RSLTS), &pAnalysisResults);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		hr = pAnalysisResults->setAttribute(CComBSTR(VSI_MSMNT_XML_ATT_VERSION), CComVariant(VSI_MSMNT_FILE_VERSION));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failed to write version");
		hr = pXmlDoc->appendChild(pAnalysisResults, NULL);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failed to append root element to document");

		hr = pAnalysisSet.CoCreateInstance(__uuidof(VsiAnalysisSet));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failed to create Analysis Set");

		// Initialize it with an output xml report file
		CComVariant vXmlDoc(pXmlDoc);
		hr = pAnalysisSet->Initialize(pApp, &vXmlDoc);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failed to initialize Analysis Set");

		// Loops selected images
		CComPtr<IVsiStudy> pStudy;
		CComPtr<IVsiSeries> pSeries;
		CComPtr<IVsiImage> pImage;
		CComPtr<IVsiMode> pMode;
		if (imageList.size() > 0)
		{
			CVsiListImageIter iter;
			for (iter = imageList.begin(); iter != imageList.end(); ++iter)
			{
				pStudy.Release();
				pSeries.Release();
				pImage.Release();

				CVsiImage &item = *iter;
				pImage = item.m_pImage;
				pMode = item.m_pMode;

				if (NULL != pImage)
				{
					hr = pImage->GetSeries(&pSeries);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pSeries->GetStudy(&pStudy);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					VSI_STUDY_ERROR dwError;
					hr = pStudy->GetErrorCode(&dwError);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (VSI_STUDY_ERROR_LOAD_XML == dwError)
					{
						WCHAR szSupport[500];
						VsiGetTechSupportInfo(szSupport, _countof(szSupport));

						CString strMsg(MAKEINTRESOURCE(IDS_STUDY_OPEN_ERROR));
						strMsg += szSupport;

						VsiMessageBox(
							GetTopLevelParent(),
							strMsg,
							VsiGetRangeValue<CString>(
								VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
								VSI_PARAMETER_SYS_PRODUCT_NAME,
								m_pPdm),
							MB_OK | MB_ICONEXCLAMATION);

						return 0;
					}

					// Copy image to analysis report
					hr = pAnalysisSet->CopyImageFromAnalysisMsmntInfo(pImage);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						L"Failure copying image to the analysis file");
				}
				else if (NULL != pMode)
				{
					CComPtr<IVsiModeData> pModeData;
					hr = pMode->GetModeData(&pModeData);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IUnknown> pUnkMsmntSet;
					hr = pModeData->GetMsmntSet(&pUnkMsmntSet);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComQIPtr<IVsiMsmntSet> pMsmntSet(pUnkMsmntSet);
					VSI_CHECK_INTERFACE(pMsmntSet, VSI_LOG_ERROR, NULL);

					hr = pMsmntSet->GenerateAnalysisReport(pAnalysisSet);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Copy curves from measurement file
					{
						CComQIPtr<IVsiPropertyBag> pModeProps(pMode);
						VSI_CHECK_INTERFACE(pModeProps, VSI_LOG_ERROR, NULL);

						CComVariant vImage;
						hr = pModeProps->ReadId(VSI_PROP_MODE_IMAGE, &vImage);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (S_OK == hr)
						{
							CComPtr<IXMLDOMDocument> pXmlDoc;
							hr = pAnalysisSet->GetReportXml(&pXmlDoc);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							CComQIPtr<IVsiImage> pImage(V_UNKNOWN(&vImage));
							VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

							hr = CopyCurves(pXmlDoc, pImage, pAnalysisSet);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
					}
				}
			}
		}

		if ((NULL != pAnalysisSet) && (imageList.size() > 0))
		{
			hr = pAnalysisSet->ProcessMsmnts();
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Process calculations for series with individual images
			// Those series are the series with unprocessed calculations
			// FALSE parameter means only marked series are going to be processed
			hr = pAnalysisSet->ProcessCalculations(FALSE, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure to process outstanding calculations");
		}

		// Debugging work
		#ifdef _DEBUG
		{
			WCHAR szPath[MAX_PATH];
			GetModuleFileName(NULL, szPath, _countof(szPath));
			LPWSTR pszSpt = wcsrchr(szPath, L'\\');
			lstrcpy(pszSpt + 1, VSI_FOLDER_DEBUG);
			lstrcat(szPath, L"\\AnalysisReport.xml");
			pXmlDoc->save(CComVariant(szPath));
		}
		#endif

		*ppXmlMsmntDoc = pXmlDoc.Detach();
	}
	VSI_CATCH(hr);

	return hr;
}

HRESULT CVsiExportImage::CopyCurves(
	IXMLDOMDocument *pXmlAnalysisReportDoc,
	IVsiImage *pImage,
	IVsiAnalysisSet *pAnalysisSet)
{
	HRESULT hr = S_OK;
	CComPtr<IXMLDOMDocument> pXmlMsmntDoc;

	try
	{
		VSI_CHECK_MEM(pXmlAnalysisReportDoc, VSI_LOG_ERROR, NULL);
		VSI_CHECK_MEM(pImage, VSI_LOG_ERROR, NULL);
		VSI_CHECK_MEM(pAnalysisSet, VSI_LOG_ERROR, NULL);

		// Get series path
		CComHeapPtr<OLECHAR> pszImagePath;
		hr = pImage->GetDataPath(&pszImagePath);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Get Output series element
		CComPtr<IXMLDOMNodeList> pOutputSeriesList;
		hr = pXmlAnalysisReportDoc->getElementsByTagName(VSI_MSMNT_REPORT_XML_ELE_SERIES, &pOutputSeriesList);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComPtr<IXMLDOMNode> pOutputSeriesNode;
		hr = pOutputSeriesList->get_item(0, &pOutputSeriesNode);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComQIPtr<IXMLDOMElement> pOutputSeriesElement(pOutputSeriesNode);
		VSI_CHECK_INTERFACE(pOutputSeriesElement, VSI_LOG_ERROR, NULL);

		// Get Image ID
		CComVariant vImageId;
		hr = pImage->GetProperty(VSI_PROP_IMAGE_ID, &vImageId);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		WCHAR szMasterMsmntPath[MAX_PATH];
		wcscpy_s(szMasterMsmntPath, _countof(szMasterMsmntPath), pszImagePath);
		PathRemoveFileSpec(szMasterMsmntPath);
		wcscat_s(szMasterMsmntPath, _countof(szMasterMsmntPath), L"\\" VSI_FILE_MEASUREMENT);

		hr = pAnalysisSet->ReadCurveInstance(
			szMasterMsmntPath,
			pOutputSeriesElement,
			V_BSTR(&vImageId));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiExportImage::ExportImage(IUnknown *pApp, IUnknown *pUnkImage, IUnknown *pUnkMode, LPCOLESTR pszPath)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_VERIFY((pApp != NULL || pUnkImage != NULL || pUnkMode != NULL), VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiApp> pVsiApp = pApp;
		VSI_CHECK_INTERFACE(pVsiApp, VSI_LOG_ERROR, NULL);

		CVsiImage item(pUnkImage, pUnkMode);
		CVsiListImage listImage;
		listImage.push_back(item);

		CComPtr<IXMLDOMDocument> pXmlReportDoc;
		hr = CreateMsmntReport(pVsiApp, &pXmlReportDoc, listImage);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vXmlDoc(pXmlReportDoc);
		hr = m_ExportReport.DoExport(pszPath, &vXmlDoc, pVsiApp);
		if (FAILED(hr))
		{
			// Delete file
			DeleteFile(pszPath);

			CString strMessage(L"Failed to export the measurement report.");
			if (HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) == hr)
			{
				strMessage += L"\n\nAccess Denied.";
			}
			else if (HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) == hr)
			{
				strMessage += L"\n\nFile is in use.";
			}
			// Report failure
			VsiMessageBox(
				NULL,
				strMessage,
				VsiGetRangeValue<CString>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_SYS_PRODUCT_NAME,
					m_pPdm),
				MB_OK | MB_ICONWARNING);
			return 0;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiExportImage::ExportReport(IUnknown *pApp, IUnknown *pUnkXmlDoc, LPCOLESTR pszPath)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_VERIFY(pApp != NULL, VSI_LOG_ERROR, NULL);

		CComQIPtr<IVsiApp> pVsiApp = pApp;
		VSI_CHECK_INTERFACE(pVsiApp, VSI_LOG_ERROR, NULL);

		CComQIPtr<IXMLDOMDocument> pXmlReportDoc(pUnkXmlDoc);

		CComVariant vXmlDoc(pXmlReportDoc);
		hr = m_ExportReport.DoExport(pszPath, &vXmlDoc, pVsiApp);
		if (FAILED(hr))
		{
			// Delete file
			DeleteFile(pszPath);

			CString strMessage(L"Failed to export the measurement report.");
			if (HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) == hr)
			{
				strMessage += L"\n\nAccess Denied.";
			}
			else if (HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) == hr)
			{
				strMessage += L"\n\nFile is in use.";
			}
			// Report failure
			VsiMessageBox(
				NULL,
				strMessage,
				VsiGetRangeValue<CString>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_GLOBAL,
					VSI_PARAMETER_SYS_PRODUCT_NAME,
					m_pPdm),
				MB_OK | MB_ICONWARNING);
			return 0;
		}
	}
	VSI_CATCH(hr);

	return hr;
}