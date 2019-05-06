/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiChangeMsmntPackageDlg.cpp
**
**	Description:
**		Implementation of the CVsiChangeMsmntPackageDlg class
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiResProduct.h>
#include <VsiSaxUtils.h>
#include <VsiCommUtl.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiServiceKey.h>
#include <VsiPdmModule.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterUtils.h>
#include <VsiParameterMode.h>
#include <VsiParamMeasurement.h>
#include <VsiMeasurementModule.h>
#include <VsiStudyXml.h>
#include <VsiConfigXML.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterMode.h>
#include <VsiParameterSystem.h>
#include "VsiChangeMsmntPackageDlg.h"


CVsiChangeMsmntPackageDlg::CVsiChangeMsmntPackageDlg(
	IVsiApp *pApp, int iImage, IVsiEnumImage *pEnumImages)
:
	m_pApp(pApp),
	m_iImageCount(iImage),
	m_pEnumImages(pEnumImages)
{
}

CVsiChangeMsmntPackageDlg::~CVsiChangeMsmntPackageDlg()
{
}

LRESULT CVsiChangeMsmntPackageDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;

	try
	{
		// Set font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
		VsiThemeRecurSetFont(m_hWnd, hFont);

		// Sets header and footer
		HIMAGELIST hilHeader(NULL);
		VsiThemeGetImageList(VSI_THEME_IL_HEADER_NORMAL, &hilHeader);
		VsiThemeDialogBox(*this, TRUE, TRUE, hilHeader);

		// Set message
		CString strMsg;
		if (1 == m_iImageCount)
		{
			strMsg = L"Change the measurement package for the selected image using the options below.";
		}
		else
		{
			strMsg.Format(L"Change the measurement package for the %d selected images using the options below.", m_iImageCount);
		}

		SetDlgItemText(IDC_CMPKG_MESSAGE, strMsg);

		// Packages
		CWindow wndPackages(GetDlgItem(IDC_CMPKG_PACKAGES));

		wndPackages.SendMessage(CB_SETCUEBANNER, 0, (LPARAM)L"Select a measurement package");

		// Query Measurement Manager
		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiMsmntManager> pMsmntMgr;

		hr = pServiceProvider->QueryService(
			VSI_SERVICE_MEASUREMENT_MANAGER,
			__uuidof(IVsiMsmntManager),
			(IUnknown**)&pMsmntMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiEnumMsmntPackage> pEnumPackages;
		hr = pMsmntMgr->GetPackages(&pEnumPackages);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiMsmntPackage> pPackage;
		while (S_OK == pEnumPackages->Next(1, &pPackage, NULL))
		{
			BOOL bEnabled(FALSE);
			hr = pPackage->GetIsEnabled(&bEnabled);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (bEnabled)
			{
				// Get the description and add it to the drop down
				CComHeapPtr<OLECHAR> pszDescription;
				hr = pPackage->GetDescription(&pszDescription);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				int iIndex = (int)wndPackages.SendMessage(CB_ADDSTRING, 0, (LPARAM)(LPWSTR)pszDescription);
				if (CB_ERR != iIndex)
				{
					wndPackages.SendMessage(CB_SETITEMDATA, iIndex, (LPARAM)pPackage.Detach());
				}
			}

			pPackage.Release();
		}
	}
	VSI_CATCH(hr);

	return 1;
}

LRESULT CVsiChangeMsmntPackageDlg::OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
{
	CWindow wndPackages(GetDlgItem(IDC_CMPKG_PACKAGES));

	int iCount = (int)wndPackages.SendMessage(CB_GETCOUNT, 0, 0);
	for (int i = 0; i < iCount; ++i)
	{
		IVsiMsmntPackage* pPackage = (IVsiMsmntPackage*)wndPackages.SendMessage(CB_GETITEMDATA, i);
		if (NULL != pPackage)
		{
			pPackage->Release();
		}
	}
	wndPackages.SendMessage(CB_RESETCONTENT);

	return 0;
}

/// <summary>
///	Clicked on the X. Assume they meant "Cancel"
/// </summary>
LRESULT CVsiChangeMsmntPackageDlg::OnClose(UINT, WPARAM, LPARAM, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Switch Measurement Package - Cancel");

	EndDialog(IDCANCEL);

	return 0;
}

/// <summary>
///	Clicked on OK. Get the new package name and close the dialog
/// </summary>
LRESULT CVsiChangeMsmntPackageDlg::OnCmdOk(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Switch Measurement Package - Switch");

	HRESULT hr(S_OK);
	int iFailed(0);
	int iLocked(0);

	try
	{
		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiSession> pSession;

		hr = pServiceProvider->QueryService(VSI_SERVICE_SESSION_MGR, __uuidof(IVsiSession),	(IUnknown**)&pSession);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiPdm> pPdm;

		hr = pServiceProvider->QueryService(VSI_SERVICE_PDM, __uuidof(IVsiPdm),	(IUnknown**)&pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		BOOL bLockAll = VsiGetBooleanValue<BOOL>(
			VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
			VSI_PARAMETER_SYS_LOCK_ALL, pPdm);

		// Get selected package
		CWindow wndPackages(GetDlgItem(IDC_CMPKG_PACKAGES));

		int index = (int)wndPackages.SendMessage(CB_GETCURSEL);
		VSI_VERIFY(CB_ERR != index, VSI_LOG_ERROR, NULL);

		CComVariant vPackage(L"");
		CComVariant vSystemPackage(L"");

		if (CB_ERR != index)
		{
			IVsiMsmntPackage* pPackage = (IVsiMsmntPackage*)wndPackages.SendMessage(CB_GETITEMDATA, index);
			VSI_CHECK_INTERFACE(pPackage, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszPackageFileName;
			hr = pPackage->GetPackageFileName(&pszPackageFileName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			vPackage = pszPackageFileName;

			CComPtr<IVsiMsmntPackage> pSysPackage;
			hr = pPackage->GetSystemPackage(&pSysPackage);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (NULL != pSysPackage)
			{
				CComHeapPtr<OLECHAR> pszSysPackageFileName;
				hr = pSysPackage->GetPackageFileName(&pszSysPackageFileName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				vSystemPackage = pszSysPackageFileName;
			}
			else
			{
				vSystemPackage = vPackage;
			}
		}

		VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(
			L"Switching measurement package: %s (system - %s)", V_BSTR(&vPackage), V_BSTR(&vSystemPackage)));

		// DO switch
		CComBSTR bstrSelParams(
			L"//" VSI_IMAGE_XML_ELM_SETTINGS L"/" 
			VSI_CFG_ELM_ROOTS L"/" VSI_CFG_ELM_ROOT L"[@" VSI_PDM_ATT_NAME L"=\'" VSI_PDM_ROOT_REVIEW L"\']/"
			VSI_CFG_ELM_GROUPS L"/" VSI_CFG_ELM_GROUP L"[@" VSI_PDM_ATT_NAME L"=\'" VSI_PDM_GROUP_GLOBAL L"\']/"
			VSI_CFG_ELM_PARAMETERS);

		CComBSTR bstrSelPackage(
			VSI_CFG_ELM_PARAMETER L"[@" VSI_PDM_ATT_NAME L"=\'"
			VSI_PARAMETER_SETTINGS_MSMNT_PACKAGE L"\']/" L"settings/@value");

		CComBSTR bstrSelSysPackage(
			VSI_CFG_ELM_PARAMETER L"[@" VSI_PDM_ATT_NAME L"=\'"
			VSI_PARAMETER_SETTINGS_MSMNT_PACKAGE_SYSTEM L"\']/" L"settings/@value");

		CComPtr<IXMLDOMDocument> pXmlDoc;
		hr = VsiCreateDOMDocument(&pXmlDoc);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiImage> pImage;
		while (S_OK == m_pEnumImages->Next(1, &pImage, NULL))
		{
			try
			{
				CComVariant vLocked(false);

				if (bLockAll)
				{
					CComPtr<IVsiStudy> pStudy;
					hr = pImage->GetStudy(&pStudy);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr && VARIANT_FALSE != V_BOOL(&vLocked))
					{
						// Locked - skip
						++iLocked;
					}
				}

				if (VARIANT_FALSE == V_BOOL(&vLocked))
				{
					CComVariant vImageNs;
					hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vImageNs);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					int iSlot(-1);
					hr = pSession->GetIsImageLoaded(V_BSTR(&vImageNs), &iSlot);
					if (S_OK == hr)
					{
						CComPtr<IUnknown> pUnkMode;
						hr = pSession->GetMode(iSlot, &pUnkMode);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(
							L"Switching measurement package: image loaded - %s", V_BSTR(&vImageNs)));

						CComQIPtr<IVsiPropertyBag> pProps(pUnkMode);
						VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

						CComVariant vRoot;
						hr = pProps->ReadId(VSI_PROP_MODE_ROOT, &vRoot);
						VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

						VsiSetRangeValue<LPCWSTR>(V_BSTR(&vPackage), V_BSTR(&vRoot),
							VSI_PDM_GROUP_MODE, VSI_PARAMETER_SETTINGS_MSMNT_PACKAGE, pPdm);

						VsiSetRangeValue<LPCWSTR>(V_BSTR(&vSystemPackage), V_BSTR(&vRoot),
							VSI_PDM_GROUP_MODE, VSI_PARAMETER_SETTINGS_MSMNT_PACKAGE_SYSTEM, pPdm);

						VsiSetParamEvent(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS,
							VSI_PARAMETER_MSMNT_PACKAGE_CHANGED, pPdm);
					}
					else
					{
						CComHeapPtr<OLECHAR> pszFile;
						hr = pImage->GetDataPath(&pszFile);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(
							L"Switching measurement package: image - %s", (LPCWSTR)pszFile));

						CComVariant vFile((LPCWSTR)pszFile);

						VARIANT_BOOL bLoaded = VARIANT_FALSE;
						hr = pXmlDoc->load(vFile, &bLoaded);
						if (hr != S_OK || bLoaded == VARIANT_FALSE)
						{
							VSI_FAIL(VSI_LOG_ERROR, NULL);
						}

						CComPtr<IXMLDOMNode> pXmlNodeParams;
						hr = pXmlDoc->selectSingleNode(bstrSelParams, &pXmlNodeParams);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (S_OK == hr)
						{
							CComPtr<IXMLDOMNode> pNodePackage;
							hr = pXmlNodeParams->selectSingleNode(bstrSelPackage, &pNodePackage);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							if (S_OK == hr)	
							{
								hr = pNodePackage->put_nodeValue(vPackage);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							}

							CComPtr<IXMLDOMNode> pNodeSysPackage;
							hr = pXmlNodeParams->selectSingleNode(bstrSelSysPackage, &pNodeSysPackage);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							if (S_OK == hr)	
							{
								hr = pNodeSysPackage->put_nodeValue(vSystemPackage);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							}

							hr = pXmlDoc->save(vFile);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
					}
				}
			}
			VSI_CATCH_(err)
			{
				VSI_LOG_MSG(VSI_LOG_INFO, L"Switch measurement package failed");

				++iFailed;
			}

			pImage.Release();
		}

		if (0 < iFailed || 0 < iLocked)
		{
			CString strMsg;
			UINT nIcon(MB_ICONWARNING);

			if (0 < iFailed)
			{
				if (iFailed == m_iImageCount)
				{
					strMsg = L"Unable to switch measurement package.\r\n";
				}
				else if (1 == iFailed)
				{
					strMsg = L"Unable to switch measurement package for 1 image.\r\n";
				}
				else
				{
					strMsg.Format(L"Unable to switch measurement package for %d images.\r\n", iFailed);
				}

				nIcon = MB_ICONERROR;
			}

			if (0 < iLocked)
			{
				CString strLocked;
				strLocked.Format(
					(1 == iLocked) ?
						L"%d image is locked and was skipped.\r\n" :
						L"%d images are locked and were skipped.\r\n",
					iLocked);

				strMsg += strLocked;
			}

			VsiMessageBox(*this, strMsg, L"Switch Measurement Package", MB_OK | nIcon);
		}

		EndDialog(IDOK);
	}
	VSI_CATCH(hr);

	return 0;
}

/// <summary>
///	Clicked on No. Just close the dialog.
/// </summary>
LRESULT CVsiChangeMsmntPackageDlg::OnCmdCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Switch Measurement Package - Cancel");

	EndDialog(IDCANCEL);

	return 0;
}

LRESULT CVsiChangeMsmntPackageDlg::OnPackageSelChange(WORD, WORD, HWND, BOOL&)
{
	CWindow wndPackages(GetDlgItem(IDC_CMPKG_PACKAGES));

	int index = (int)wndPackages.SendMessage(CB_GETCURSEL);
	if (CB_ERR != index)
	{
		GetDlgItem(IDOK).EnableWindow(TRUE);
	}

	return 0;
}
