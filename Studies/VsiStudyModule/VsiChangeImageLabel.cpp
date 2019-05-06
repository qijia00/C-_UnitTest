/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiChangeImageLabel.cpp
**
**	Description:
**		Implementation of CVsiChangeImageLabel
**
*******************************************************************************/

#include "stdafx.h"
#include <shlguid.h>
#include <VsiServiceProvider.h>
#include <VsiCommUtlLib.h>
#include <VsiCommUtl.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiPdmModule.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiServiceKey.h>
#include <VsiAppViewModule.h>
#include <VsiModeModule.h>
#include <VsiResProduct.h>
#include "VsiStudyXml.h"
#include "VsiChangeImageLabel.h"


#define VSI_AUTOCOMPLETE_NAME L"HistoryImageLabel.xml"


CVsiChangeImageLabel::CVsiChangeImageLabel() :
	m_bCommit(TRUE)
{
}

CVsiChangeImageLabel::~CVsiChangeImageLabel()
{
}

STDMETHODIMP CVsiChangeImageLabel::Activate(
	IVsiApp *pApp,
	HWND hwndParent,
	IUnknown *pUnkItem)
{
	HRESULT hr(S_OK);
	HWND hwnd(NULL);

	try
	{
		VSI_CHECK_ARG(pApp, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(pUnkItem, VSI_LOG_ERROR, NULL);

		m_pApp = pApp;
		m_pUnkItem = pUnkItem;

		hwnd = Create(hwndParent);
		VSI_WIN32_VERIFY(NULL != hwnd, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiChangeImageLabel::Deactivate()
{
	BOOL bCommited(FALSE);

	if (IsWindow())
	{
		HRESULT hr;

		try
		{
			if (m_bCommit)
			{
				hr = CommitLabel();
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (S_OK == hr)
				{
					bCommited = TRUE;
				}
			}
		}
		VSI_CATCH(hr);

		DestroyWindow();
	}

	return bCommited ? S_OK : S_FALSE;
}

STDMETHODIMP CVsiChangeImageLabel::PreTranslateAccel(MSG *pMsg, BOOL *pbHandled)
{
	*pbHandled = IsDialogMessage(pMsg);

	return S_OK;
}

LRESULT CVsiChangeImageLabel::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr;

	try
	{
		// Sets font
		HFONT hFont;
		VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
		VsiThemeRecurSetFont(m_hWnd, hFont);

		// Sets header and footer
		HIMAGELIST hilHeader(NULL);
		VsiThemeGetImageList(VSI_THEME_IL_HEADER_NORMAL, &hilHeader);
		VsiThemeDialogBox(*this, TRUE, TRUE, hilHeader);

		SendDlgItemMessage(IDC_IL_IMAGE_LABEL, EM_SETLIMITTEXT, VSI_IMAGE_NAME_MAX);

		// Auto Complete - Keep going on error
		try
		{
			hr = m_pAutoComplete.CoCreateInstance(CLSID_AutoComplete);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

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

			hr = m_pAutoComplete->Init(GetDlgItem(IDC_IL_IMAGE_LABEL), pEnumString, NULL, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
			CComQIPtr<IAutoComplete2> pAutoComplete2(m_pAutoComplete);
			if (pAutoComplete2 != NULL)
			{
				hr = pAutoComplete2->SetOptions(
					ACO_AUTOSUGGEST | ACO_UPDOWNKEYDROPSLIST);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}
		VSI_CATCH(hr);

		// Set label
		CComQIPtr<IVsiImage> pImage(m_pUnkItem);
		if (NULL != pImage)
		{
			CComVariant vLabel;
			hr = pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vLabel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (VT_EMPTY != V_VT(&vLabel))
				SetDlgItemText(IDC_IL_IMAGE_LABEL, V_BSTR(&vLabel));
		}
		else
		{
			CComQIPtr<IVsiModeData> pModeData(m_pUnkItem);
			if (NULL != pModeData)
			{
				CComQIPtr<IVsiPropertyBag> pProps(pModeData);
				VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

				CComVariant vLabel;
				hr = pProps->ReadId(VSI_PROP_MODE_DATA_NAME, &vLabel);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (VT_EMPTY != V_VT(&vLabel))
					SetDlgItemText(IDC_IL_IMAGE_LABEL, V_BSTR(&vLabel));
			}
			else
			{
				// Error
			}
		}

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Adds pre-translate message filter
		{
			CComPtr<IVsiMessageFilter> pMsgFilter;

			// Get App View
			hr = pServiceProvider->QueryService(
				VSI_SERVICE_APP_VIEW,
				__uuidof(IVsiMessageFilter),
				(IUnknown**)&pMsgFilter);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pMsgFilter->AddPreTranslateAccelFilter(
				static_cast<IVsiPreTranslateAccelFilter*>(this));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComQIPtr<IVsiAppView> pAppView(pMsgFilter);
			pAppView->UnblockControlPanel();
		}

		// Get PDM
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Query Cursor Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CURSOR_MANAGER,
			__uuidof(IVsiCursorManager),
			(IUnknown **)&m_pCursorMgr);
		VSI_CHECK_INTERFACE(m_pCursorMgr, VSI_LOG_ERROR, L"Failure getting cursor manager interface");

		// Set cursor state
		m_pCursorMgr->SetState(VSI_CURSOR_STATE_ON, NULL);

		// Clip mouse cursor to window
		VSI_APP_MODE_FLAGS dwAppMode(VSI_APP_MODE_NONE);
		m_pApp->GetAppMode(&dwAppMode);
		CenterMouseCursor(dwAppMode, IDOK);
	}
	VSI_CATCH(hr);

	ShowWindow(SW_SHOW);

	return 1;
}

LRESULT CVsiChangeImageLabel::OnDestroy(UINT, WPARAM, LPARAM, BOOL& bHandled)
{
	HRESULT hr = S_OK;

	try
	{
		// Removes pre-translate message filter
		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiMessageFilter> pMsgFilter;

		// Get App View
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_APP_VIEW,
			__uuidof(IVsiMessageFilter),
			(IUnknown**)&pMsgFilter);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pMsgFilter->RemovePreTranslateAccelFilter(
			static_cast<IVsiPreTranslateAccelFilter*>(this));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH(hr);

	m_pUnkItem.Release();

	m_pCursorMgr.Release();

	m_pPdm.Release();

	m_pApp.Release();

	bHandled = FALSE;

	return 0;
}

LRESULT CVsiChangeImageLabel::OnCmdOk(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Image Label - OK");

	HRESULT hr;

	try
	{
		hr = CommitLabel();
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CloseWindow(TRUE);
	}
	VSI_CATCH(hr);

	return 0;
}

LRESULT CVsiChangeImageLabel::OnCmdCancel(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Image Label - Cancel");

	CloseWindow(FALSE);

	return 0;
}

HRESULT CVsiChangeImageLabel::CommitLabel()
{
	HRESULT hr(S_FALSE);
	BOOL bCommited = FALSE;

	try
	{
		WCHAR szLabel[VSI_IMAGE_NAME_MAX + 1];

		GetDlgItemText(IDC_IL_IMAGE_LABEL, szLabel, _countof(szLabel));

		// Save Auto Complete settings
		if (*szLabel != 0 && NULL != m_pAutoCompleteSource)
		{
			WCHAR pszPath[MAX_PATH];
			BOOL bRet = VsiGetApplicationDataDirectory(
				AtlFindStringResourceInstance(IDS_VSI_PRODUCT_NAME),
				pszPath);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, NULL);

			PathAppend(pszPath, VSI_AUTOCOMPLETE_NAME);

			hr = m_pAutoCompleteSource->Push(szLabel);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = m_pAutoCompleteSource->SaveSettings(pszPath, 10);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
		}

		CComVariant vLabel(szLabel);

		// Get session and slot
		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiSession> pSession;
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_SESSION_MGR,
			__uuidof(IVsiSession),
			(IUnknown**)&pSession);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		// Set label
		CComQIPtr<IVsiImage> pImage(m_pUnkItem);
		if (NULL != pImage)
		{
			CComVariant vLabelOld;
			hr = pImage->GetProperty(VSI_PROP_IMAGE_NAME, &vLabelOld);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (vLabel != vLabelOld)
			{
				bCommited = TRUE;

				hr = pImage->SetProperty(VSI_PROP_IMAGE_NAME, &vLabel);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Get mode data
				CComVariant vNs;
				hr = pImage->GetProperty(VSI_PROP_IMAGE_NS, &vNs);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				int iSlot = 0;
				HRESULT hrLoaded = pSession->GetIsItemLoaded(V_BSTR(&vNs), &iSlot);
				if (S_OK == hrLoaded)
				{
					CComPtr<IUnknown> pUnkMode;
					hr = pSession->GetMode(iSlot, &pUnkMode);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComQIPtr<IVsiMode> pMode(pUnkMode);
					VSI_CHECK_INTERFACE(pMode, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiModeData> pModeData;
					hr = pMode->GetModeData(&pModeData);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComQIPtr<IVsiPropertyBag> pProps(pModeData);
					VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

					hr = pProps->WriteId(VSI_PROP_MODE_DATA_NAME, &vLabel);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}

				CComQIPtr<IVsiPersistImage> pPersist(pImage);
				VSI_CHECK_INTERFACE(pPersist, VSI_LOG_ERROR, NULL);

				hr = pPersist->ResaveProperties(NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				VsiSetParamEvent(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
					VSI_PARAMETER_EVENTS_GENERAL_IMAGE_PROP_UPDATE, m_pPdm);
			}
		}
		else
		{
			CComQIPtr<IVsiModeData> pModeData(m_pUnkItem);
			if (NULL != pModeData)
			{
				CComQIPtr<IVsiPropertyBag> pProps(pModeData);
				VSI_CHECK_INTERFACE(pProps, VSI_LOG_ERROR, NULL);

				CComVariant vLabelOld;
				hr = pProps->ReadId(VSI_PROP_MODE_DATA_NAME, &vLabelOld);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (vLabel != vLabelOld)
				{
					bCommited = TRUE;

					hr = pProps->WriteId(VSI_PROP_MODE_DATA_NAME, &vLabel);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Get image
					int iSlotActive = 0;
					hr = pSession->GetActiveSlot(&iSlotActive);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CComPtr<IVsiImage> pImageActive;
					hr = pSession->GetImage(iSlotActive, &pImageActive);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					// Resave new label
					if (NULL != pImageActive)
					{
						hr = pImageActive->SetProperty(VSI_PROP_IMAGE_NAME, &vLabel);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						CComQIPtr<IVsiPersistImage> pPersist(pImageActive);
						VSI_CHECK_INTERFACE(pPersist, VSI_LOG_ERROR, NULL);

						hr = pPersist->ResaveProperties(NULL);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					}
				}
			}
			else
			{
				// Error
			}
		}
	}
	VSI_CATCH(hr);

	return SUCCEEDED(hr) ? (bCommited ? S_OK : S_FALSE) : hr;
}

void CVsiChangeImageLabel::CloseWindow(BOOL bOk)
{
	HRESULT hr = S_OK;
	CComPtr<IVsiCommandManager> pCmdMgr;

	try
	{
		m_bCommit = bOk;

		CComQIPtr<IVsiServiceProvider> pServiceProvider(m_pApp);
		VSI_CHECK_INTERFACE(pServiceProvider, VSI_LOG_ERROR, NULL);

		// Get Command Manager
		hr = pServiceProvider->QueryService(
			VSI_SERVICE_CMD_MGR,
			__uuidof(IVsiCommandManager),
			(IUnknown**)&pCmdMgr);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComPtr<IVsiPropertyBag> pProps;
		hr = pProps.CoCreateInstance(__uuidof(VsiPropertyBag));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vCloseCode(FALSE != bOk);
		hr = pProps->Write(L"closecode", &vCloseCode);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CComVariant vParams(pProps.p);
		pCmdMgr->InvokeCommand(ID_CMD_IMAGE_LABEL, &vParams);  // Close panel
	}
	VSI_CATCH(hr);
}

