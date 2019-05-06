/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudy.h
**
**	Description:
**		Declaration and Implementation of the CVsiChangeStudyInfoDlg
**
*******************************************************************************/

#pragma once

#include <VsiWtl.h>
#include <VsiOperatorModule.h>
#include <VsiStudyModule.h>
#include <VsiToolTipHelper.h>
#include <VsiSelectOperator.h>
#include <VsiStudyAccessCombo.h>


class CVsiChangeStudyInfoDlg :
	public CDialogImpl<CVsiChangeStudyInfoDlg>,
	public CVsiToolTipHelper<CVsiChangeStudyInfoDlg>
{
private:
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;
	CComPtr<IVsiStudyManager> m_pStudyMgr;
	CComPtr<IVsiSession> m_pSession;
	CComPtr<IVsiEnumStudy> m_pEnumStudy;
	CVsiSelectOperator m_wndOperatorOwner;

	CVsiStudyAccessCombo m_studyAccess;

	BOOL m_bAdvanced;

	typedef enum
	{
		VSI_TIPS_NONE = 0,
		VSI_TIPS_OWNER,
		VSI_TIPS_STUDY_SHARE,
	} VSI_TIPS;  // Tips state

	VSI_TIPS m_TipsState;

public:
	WORD IDD;

	CVsiChangeStudyInfoDlg(
		IVsiPdm *pPdm,
		IVsiOperatorManager *pOperatorMgr,
		IVsiStudyManager *pStudyMgr,
		IVsiSession *pSessionObj,
		IVsiEnumStudy *pEnumStudy,
		BOOL bAdvanced)
	:
		m_pPdm(pPdm),
		m_pOperatorMgr(pOperatorMgr),
		m_pStudyMgr(pStudyMgr),
		m_pSession(pSessionObj),
		m_pEnumStudy(pEnumStudy),
		m_bAdvanced(bAdvanced),
		m_TipsState(VSI_TIPS_NONE)
	{
		IDD = bAdvanced ? IDD_CHANGE_STUDIES_INFO_ADV : IDD_CHANGE_STUDIES_INFO;
	}

	~CVsiChangeStudyInfoDlg()
	{
	}

protected:
	BEGIN_MSG_MAP(CVsiChangeStudyInfoDlg)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		COMMAND_HANDLER(IDC_CSI_OPERATOR_CHANGE_OWNER, BN_CLICKED, OnBnClickedChangeOwner)
		COMMAND_HANDLER(IDC_CSI_OPERATOR_CHANGE_ACCESS, BN_CLICKED, OnBnClickedChangeAccess)
		COMMAND_ID_HANDLER(IDC_CSI_OPERATOR_NAME, OnSelChange)
		COMMAND_ID_HANDLER(IDC_CSI_STUDY_ACCESS, OnSelChange)
		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)

		NOTIFY_CODE_HANDLER(TTN_GETDISPINFO, OnTipsGetDisplayInfo);

		CHAIN_MSG_MAP(CVsiToolTipHelper<CVsiChangeStudyInfoDlg>)
		CHAIN_MSG_MAP_MEMBER(m_wndOperatorOwner)
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
	{
		HRESULT hr(S_OK);

		try
		{
			VSI_VERIFY(NULL != m_pEnumStudy, VSI_LOG_ERROR, NULL);

			// Sets font
			HFONT hFont;
			VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
			VsiThemeRecurSetFont(m_hWnd, hFont);

			// Sets header and footer
			HIMAGELIST hilHeader(NULL);
			VsiThemeGetImageList(VSI_THEME_IL_HEADER_NORMAL, &hilHeader);
			VsiThemeDialogBox(*this, TRUE, TRUE, hilHeader);

			GetDlgItem(IDC_CSI_MESSAGE).SetWindowText(CString(MAKEINTRESOURCE(IDS_STYBRW_CHANGE_STUDIES_INFO)));

			if (m_bAdvanced)
			{
				hr = m_wndOperatorOwner.Initialize(
					FALSE,
					FALSE,
					FALSE,
					GetDlgItem(IDC_CSI_OPERATOR_NAME),
					m_pOperatorMgr,
					NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				m_wndOperatorOwner.Fill(NULL);
			}

			m_studyAccess.Initialize(GetDlgItem(IDC_CSI_STUDY_ACCESS));
			m_studyAccess.EnableWindow(!m_bAdvanced);

			m_studyAccess.SendMessage(CB_SETCUEBANNER, 0, (LPARAM)L"Select a Sharing setting");

			PostMessage(WM_NEXTDLGCTL, (WPARAM)(HWND)GetDlgItem(IDCANCEL), TRUE);
		}
		VSI_CATCH(hr);

		if (FAILED(hr))
		{
			EndDialog(IDCANCEL);
		}

		return 0;
	}

	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&)
	{
		m_wndOperatorOwner.Uninitialize();

		return 0;
	}

	LRESULT OnBnClickedChangeOwner(WORD, WORD, HWND, BOOL&)
	{
		DeactivateTips();

		BOOL bOwnerChecked = BST_CHECKED == IsDlgButtonChecked(IDC_CSI_OPERATOR_CHANGE_OWNER);
		m_wndOperatorOwner.EnableWindow(bOwnerChecked);

		BOOL bAccessChecked = BST_CHECKED == IsDlgButtonChecked(IDC_CSI_OPERATOR_CHANGE_ACCESS);
		GetDlgItem(IDOK).EnableWindow(bOwnerChecked || bAccessChecked);

		return 0;
	}

	LRESULT OnBnClickedChangeAccess(WORD, WORD, HWND, BOOL&)
	{
		DeactivateTips();

		BOOL bOwnerChecked = BST_CHECKED == IsDlgButtonChecked(IDC_CSI_OPERATOR_CHANGE_OWNER);
		BOOL bAccessChecked = BST_CHECKED == IsDlgButtonChecked(IDC_CSI_OPERATOR_CHANGE_ACCESS);

		GetDlgItem(IDC_CSI_STUDY_ACCESS).EnableWindow(bAccessChecked);

		GetDlgItem(IDOK).EnableWindow(bOwnerChecked || bAccessChecked);

		return 0;
	}

	LRESULT OnSelChange(WORD, WORD, HWND, BOOL &bHandled)
	{
		DeactivateTips();

		bHandled = FALSE;

		return 0;
	}

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&)
	{
		VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Change Study Access - OK");

		HRESULT hr(S_OK);

		try
		{
			DeactivateTips();

			CString strOwnerOperator;
			bool bContinue(true);

			BOOL bOwnerChecked = BST_CHECKED == IsDlgButtonChecked(IDC_CSI_OPERATOR_CHANGE_OWNER);
			if (bOwnerChecked)
			{
				int index = (int)m_wndOperatorOwner.SendMessage(CB_GETCURSEL);
				if (CB_ERR == index)
				{
					m_wndOperatorOwner.SetFocus();

					ActivateTips(VSI_TIPS_OWNER);

					bContinue = false;
				}
				else
				{
					m_wndOperatorOwner.GetWindowText(strOwnerOperator);
				}
			}

			BOOL bAccessChecked = FALSE;
			if (m_bAdvanced)
			{
				bAccessChecked = BST_CHECKED == IsDlgButtonChecked(IDC_CSI_OPERATOR_CHANGE_ACCESS);
			}
			else
			{
				bAccessChecked = TRUE;
			}

			if (bContinue && bAccessChecked)
			{
				CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE type = m_studyAccess.GetStudyAccess();

				if (CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_NONE == type)
				{
					GetDlgItem(IDC_CSI_STUDY_ACCESS).SetFocus();

					ActivateTips(VSI_TIPS_STUDY_SHARE);

					bContinue = false;
				}
			}

			if (bContinue)
			{
				CWaitCursor wait;

				// XML for UI update
				CComPtr<IXMLDOMDocument> pUpdateXmlDoc;
				hr = VsiCreateDOMDocument(&pUpdateXmlDoc);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating DOMDocument");

				CComPtr<IXMLDOMElement> pElmUpdate;
				hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_UPDATE), &pElmUpdate);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pUpdateXmlDoc->appendChild(pElmUpdate, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				CComPtr<IXMLDOMElement> pElmRefresh;
				hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_REFRESH), &pElmRefresh);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pElmUpdate->appendChild(pElmRefresh, NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				BOOL bLockAll = VsiGetBooleanValue<BOOL>(
					VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
					VSI_PARAMETER_SYS_LOCK_ALL, m_pPdm);

				int iFailed(0);
				int iSucceeded(0);
				int iSkippedLocked(0);
				int iSkippedNoRights(0);
				int iSkippedActive(0);

				CComVariant vOwnerOperator(strOwnerOperator.GetString());

				CComPtr<IVsiStudy> pStudy;
				while (S_OK == m_pEnumStudy->Next(1, &pStudy, NULL))
				{
					CComVariant vLocked(false);

					if (bLockAll)
					{
						hr = pStudy->GetProperty(VSI_PROP_STUDY_LOCKED, &vLocked);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}

					if (S_OK != hr || VARIANT_FALSE == V_BOOL(&vLocked))
					{
						try
						{
							BOOL bAuthenticated(TRUE);
							CComHeapPtr<OLECHAR> pszName;
							if (!m_bAdvanced) // If not admin, check study owner is same
							{
								CComPtr<IVsiOperator> pOperator;
								hr = m_pOperatorMgr->GetCurrentOperator(&pOperator);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								if (S_OK == hr)
								{
									hr = pOperator->GetName(&pszName);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								}

								CComVariant vOwner(L"");
								hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								if (S_OK == hr && (0 != *V_BSTR(&vOwner)))
								{
									if (0 != wcscmp(pszName, V_BSTR(&vOwner)))
									{
										bAuthenticated = FALSE;
										++iSkippedNoRights;

										CComVariant vStudyName;
										hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vStudyName);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

										VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(
											L"Study %s skipped due to insufficient right", V_BSTR(&vStudyName)));
									}
								}
							}

							// If selected study contains active series, skip
							BOOL bActive(FALSE);
							CComVariant vOwner;
							hr = pStudy->GetProperty(VSI_PROP_STUDY_OWNER, &vOwner);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
							if (S_OK == hr)
							{
								CComPtr<IVsiOperator> pOperator;
								hr = m_pOperatorMgr->GetOperator(VSI_OPERATOR_PROP_NAME, V_BSTR(&vOwner), &pOperator);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
								if (S_OK == hr)
								{
									CComHeapPtr<OLECHAR> pszOperatorId;

									hr = pOperator->GetId(&pszOperatorId);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									CComHeapPtr<OLECHAR> pszSeriesNs;
									hr = m_pSession->CheckActiveSeries(FALSE, pszOperatorId, &pszSeriesNs);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									if (NULL != pszSeriesNs)
									{
										CComPtr<IUnknown> pUnkSeries;
										hr = m_pStudyMgr->GetItem(pszSeriesNs, &pUnkSeries);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

										CComQIPtr<IVsiSeries> pSeries(pUnkSeries);
										VSI_CHECK_INTERFACE(pSeries, VSI_LOG_ERROR, NULL);

										CComPtr<IVsiStudy> pActiveStudy;
										hr = pSeries->GetStudy(&pActiveStudy);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
										if (S_OK == hr)
										{
											CComVariant vActiveId;
											hr = pActiveStudy->GetProperty(VSI_PROP_STUDY_ID, &vActiveId);
											VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
											if (S_OK == hr)
											{
												CComVariant vId;
												hr = pStudy->GetProperty(VSI_PROP_STUDY_ID, &vId);
												VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
												if (S_OK == hr)
												{
													if (vActiveId == vId)
													{
														bActive = TRUE;
														++iSkippedActive;

														CComVariant vStudyName;
														hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vStudyName);
														VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

														VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(
															L"Study %s skipped for containing open series", V_BSTR(&vStudyName)));
													}
												}
											}
										}
									}
								}
							}

							if ((bAuthenticated || m_bAdvanced) && !bActive)
							{
								if (!strOwnerOperator.IsEmpty())
								{
									// Change owner
									hr = pStudy->SetProperty(VSI_PROP_STUDY_OWNER, &vOwnerOperator);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									// Set Last Modified By and Date
									SYSTEMTIME st;
									GetSystemTime(&st);
									COleDateTime date(st);
									CComVariant vDate;
									V_VT(&vDate) = VT_DATE;
									V_DATE(&vDate) = DATE(date);
									hr = pStudy->SetProperty(VSI_PROP_STUDY_LAST_MODIFIED_DATE, &vDate);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
									// Get current operator name
									CComVariant vOwnerCurrent(L"");
									CComPtr<IVsiOperator> pOperatorCurrent;
									hr = m_pOperatorMgr->GetCurrentOperator(&pOperatorCurrent);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
									if (NULL != pOperatorCurrent)
									{
										CComHeapPtr<OLECHAR> pszCurrentOperName;
										hr = pOperatorCurrent->GetName(&pszCurrentOperName);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

										vOwnerCurrent = pszCurrentOperName;
									}
									hr = pStudy->SetProperty(VSI_PROP_STUDY_LAST_MODIFIED_BY, &vOwnerCurrent);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									CComVariant vStudyName;
									hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vStudyName);
									VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

									VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(
										L"Study %s owner successfully changed to %s", V_BSTR(&vStudyName), V_BSTR(&vOwnerOperator)));
								}

								if (bAccessChecked)
								{
									CComVariant vValue;
									CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE type = m_studyAccess.GetStudyAccess();
									switch (type)
									{
									case CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PRIVATE:
										{
											vValue = VSI_STUDY_ACCESS_PRIVATE;
										}
										break;
									case CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_GROUP:
										{
											vValue = VSI_STUDY_ACCESS_GROUP;
										}
										break;
									case CVsiStudyAccessCombo::VSI_STUDY_ACCESS_TYPE_PUBLIC:
										{
											vValue = VSI_STUDY_ACCESS_PUBLIC;
										}
										break;
									}

									if (VT_BSTR == V_VT(&vValue))
									{
										hr = pStudy->SetProperty(VSI_PROP_STUDY_ACCESS, &vValue);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

										CComVariant vStudyName;
										hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vStudyName);
										VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

										VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(
											L"Study access for %s is successfully changed to %s", V_BSTR(&vStudyName), V_BSTR(&vValue)));
									}
								}

								CComQIPtr<IVsiPersistStudy> pStudyPersist(pStudy);
								hr = pStudyPersist->SaveStudyData(NULL, 0);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								// Add to refresh XML
								CComVariant vNs;
								hr = pStudy->GetProperty(VSI_PROP_STUDY_NS, &vNs);
								VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

								CComPtr<IXMLDOMElement> pElmItem;
								hr = pUpdateXmlDoc->createElement(CComBSTR(VSI_UPDATE_XML_ELM_ITEM), &pElmItem);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								hr = pElmItem->setAttribute(CComBSTR(VSI_UPDATE_XML_ATT_NS), vNs);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								hr = pElmRefresh->appendChild(pElmItem, NULL);
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								++iSucceeded;
							}
						}
						VSI_CATCH_(err)
						{
							++iFailed;
							// Report errors at the end
							CComVariant vStudyName;
							hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vStudyName);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

							VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(
								L"Failed to change access informaiton for Study %s", V_BSTR(&vStudyName)));
						}
					}
					else
					{
						++iSkippedLocked;

						CComVariant vStudyName;
						hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &vStudyName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(
							L"Study %s is locked and is skipped", V_BSTR(&vStudyName)));
					}

					pStudy.Release();
				}
		
				if (0 < iFailed || 0 < iSkippedLocked || 0 < iSkippedNoRights || 0 < iSkippedActive)
				{
					CString strMsg;

					if (0 < iFailed)
					{
						CString strFailed;

						strFailed.Format(
							(1 == iFailed) ?
								L"%d study failed.\r\n" :
								L"%d studies failed.\r\n",
							iFailed);

						strMsg += strFailed;
					}

					if (0 < iSkippedLocked)
					{
						CString strSkipped;
						strSkipped.Format(
							(1 == iSkippedLocked) ?
								L"%d study is locked and was skipped.\r\n" :
								L"%d studies are locked and were skipped.\r\n",
							iSkippedLocked);

						strMsg += strSkipped;
					}

					if (0 < iSkippedNoRights)
					{
						CString strSkipped;
						strSkipped.Format(
							(1 == iSkippedNoRights) ?
								L"%d study was skipped due to insufficient rights.\r\n" :
								L"%d studies were skipped due to insufficient rights.\r\n",
							iSkippedNoRights);

						strMsg += strSkipped;
					}

					if (0 < iSkippedActive)
					{
						CString strSkipped;
						strSkipped.Format(
							(1 == iSkippedActive) ?
								L"%d study containing open series was skipped.\r\n" :
								L"%d studies containing open series were skipped.\r\n",
							iSkippedActive);

						strMsg += strSkipped;
					}

					VsiMessageBox(*this, strMsg, L"Change Study Access",
						MB_OK | ((0 < iFailed) ? MB_ICONERROR : MB_ICONWARNING));
				}

				// Update Study Manager
				hr = m_pStudyMgr->Update(pUpdateXmlDoc);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				EndDialog(IDOK);
			}
		}
		VSI_CATCH(hr);

		return 0;
	}

	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&)
	{
		VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Change Study Access - Cancel");

		DeactivateTips();

		EndDialog(IDCANCEL);

		return 0;
	}

	LRESULT OnTipsGetDisplayInfo(int, LPNMHDR pnmh, BOOL&)
	{
		LPNMTTDISPINFO pttd = (LPNMTTDISPINFO)pnmh;

		::SendMessage(pnmh->hwndFrom, TTM_SETMAXTIPWIDTH, 0, 300);

		if (VSI_TIPS_OWNER == m_TipsState)
		{
			pttd->lpszText = L"Select a new study owner.";
		}
		else if (VSI_TIPS_STUDY_SHARE == m_TipsState)
		{
			pttd->lpszText = L"Select a new study sharing setting.";
		}
		else
		{
			pttd->lpszText = L"";
		}

		return 0;
	}

private:
	void ActivateTips(VSI_TIPS state)
	{
		m_TipsState = state;

		WORD wId = 0;
		switch (state)
		{
		case VSI_TIPS_OWNER:
			wId = IDC_CSI_OPERATOR_NAME;
			break;
		case VSI_TIPS_STUDY_SHARE:
			wId = IDC_CSI_STUDY_ACCESS;
			break;
		default:
			_ASSERT(0);
		}

		CVsiToolTipHelper::ActivateTips(GetDlgItem(wId));
	}

	void DeactivateTips()
	{
		m_TipsState = VSI_TIPS_NONE;

		CVsiToolTipHelper::DeactivateTips();
	}
};
