/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiDeleteConfirmation.cpp
**
**	Description:
**		Implementation of the CVsiDeleteConfirmation class
**
*******************************************************************************/

#include "stdafx.h"
#include <Richedit.h>
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiThemeColor.h>
#include <VsiMeasurementModule.h>
#include <VsiResProduct.h>
#include "VsiDeleteConfirmation.h"


CVsiDeleteConfirmation::CVsiDeleteConfirmation(
	IUnknown **ppStudy,
	int iStudy,
	IUnknown **ppSeries,
	int iSeries,
	IUnknown **ppImage,
	int iImage)
:
	m_ppStudy(ppStudy),
	m_iStudyCount(iStudy),
	m_ppSeries(ppSeries),
	m_iSeriesCount(iSeries),
	m_ppImage(ppImage),
	m_iImageCount(iImage)
{
}

CVsiDeleteConfirmation::~CVsiDeleteConfirmation()
{
}

/// <summary>
///	Build the prompt based on the number of studies selected and whether
///	there are un-archived studies in the list.
///
///	If there are un-arhcived studies selected for deleting then turn
///	on the check box to allow the user to select if we should only
///	delete archived studies.
/// </summary>
LRESULT CVsiDeleteConfirmation::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
	HRESULT hr = S_OK;
	CComVariant v;

	// Set font
	HFONT hFont;
	VsiThemeGetFont(VSI_THEME_FONT_M, &hFont);
	VsiThemeRecurSetFont(m_hWnd, hFont);

	CString strMsg;
	CString strMsmnts;
	UINT uiImage = VSI_THEME_IL_HEADER_QUESTION;

	try
	{
		int iNumMeasurements = CheckForMeasurementsOnImage();
		if (iNumMeasurements > 0)
		{
			strMsmnts.Format(L" containing %d measurements or annotations", iNumMeasurements);
		}

		int iSelected = m_iStudyCount + m_iSeriesCount + m_iImageCount;
		int iExported = 0;

		CComVariant vtbExported;
		
		for (int i = 0; i < m_iStudyCount; ++i)
		{
			CComQIPtr<IVsiStudy> pStudy(m_ppStudy[i]);
			VSI_CHECK_INTERFACE(pStudy, VSI_LOG_ERROR, NULL);

			vtbExported.Clear();
			hr = pStudy->GetProperty(VSI_PROP_STUDY_EXPORTED, &vtbExported);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr)
			{
				if (VARIANT_FALSE != V_BOOL(&vtbExported))
					++iExported;
			}
		}

		for (int i = 0; i < m_iSeriesCount; ++i)
		{
			CComQIPtr<IVsiSeries> pSeries(m_ppSeries[i]);
			VSI_CHECK_INTERFACE(pSeries, VSI_LOG_ERROR, NULL);

			vtbExported.Clear();
			hr = pSeries->GetProperty(VSI_PROP_SERIES_EXPORTED, &vtbExported);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			if (S_OK == hr)
			{
				if (VARIANT_FALSE != V_BOOL(&vtbExported))
					++iExported;
			}
		}

		if (1 == iSelected)
		{
			if (1 == m_iStudyCount)
			{
				CComQIPtr<IVsiStudy> pStudy(m_ppStudy[0]);
				VSI_CHECK_INTERFACE(pStudy, VSI_LOG_ERROR, NULL);

				v.Clear();
				hr = pStudy->GetProperty(VSI_PROP_STUDY_NAME, &v);
				VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

				LPWSTR pszStudyName = (LPWSTR)V_BSTR(&v);
				if ((0 != pszStudyName) && lstrlen(pszStudyName))
				{
					strMsg = CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_STUDIES_SELECTED));
					strMsg += pszStudyName;
					strMsg += L"\n\n";
				}
				else
				{
					strMsg = CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_ONE_STUDY_SELECTED));						
				}

				//if (1 != iExported)
				//{
				//	// Item is not archived
				//	uiImage = VSI_THEME_IL_HEADER_WARNING;
				//	bShowCheckBox = TRUE;
				//	strMsg += CString(MAKEINTRESOURCE(IDS_DELCONF_STUDY_NOT_EXPORTED_SINCE_COPY));
				//}
			}
			else if (1 == m_iSeriesCount)
			{
				CComQIPtr<IVsiSeries> pSeries(m_ppSeries[0]);
				VSI_CHECK_INTERFACE(pSeries, VSI_LOG_ERROR, NULL);

				v.Clear();
				hr = pSeries->GetProperty(VSI_PROP_SERIES_NAME, &v);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				LPWSTR pszSeriesName = (LPWSTR)V_BSTR(&v);
				if ((0 != pszSeriesName) && lstrlen(pszSeriesName))
				{
					strMsg = CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_SERIES_SELECTED));
					strMsg += pszSeriesName;
					strMsg += L"\n\n";
				}
				else
				{
					strMsg = CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_ONE_SERIES_SELECTED));					
				}

				//if (1 != iExported)
				//{
				//	// Item is not archived
				//	uiImage = VSI_THEME_IL_HEADER_WARNING;
				//	bShowCheckBox = TRUE;
				//	strMsg += L"The series has not yet been exported or has been modified since it was last copied. ";
				//}
			}
			else if (1 == m_iImageCount)
			{
				CComQIPtr<IVsiImage> pImage(m_ppImage[0]);
				VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

				v.Clear();
				hr = pImage->GetProperty(VSI_PROP_IMAGE_NAME, &v);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				LPWSTR pszImageName = (LPWSTR)V_BSTR(&v);
				if ((0 != pszImageName) && lstrlen(pszImageName))
				{
					strMsg = CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_IMAGE_SELECTED));
					if (iNumMeasurements > 0)
					{
						strMsg += L", ";
					}
					else
					{
						strMsg += L" ";
					}
					strMsg += pszImageName;
				}
				else
				{
					strMsg = CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_ONE_IMAGE_SELECTED));						
				}

				if (iNumMeasurements > 0)
				{
					strMsg += L", ";
					strMsg += strMsmnts;
				}
				strMsg += L".\n\n";
			}
			else
			{
				_ASSERT(0);
			}

			strMsg += L"This action will permanently delete all information for the selected item from your system.\n";
		}
		else
		{
			if (m_iStudyCount == iSelected)
			{
				CString strText(CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_N_STUDIES_SELECTED)));
				strMsg.Format(strText, m_iStudyCount);
			}
			else if (m_iSeriesCount == iSelected)
			{
				CString strText(CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_N_SERIES_SELECTED)));
				strMsg.Format(strText, m_iSeriesCount);
			}
			else if (m_iImageCount == iSelected)
			{
				CString strText(CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_N_IMAGES_SELECTED)));
				strMsg.Format(strText, m_iImageCount);
				if (iNumMeasurements > 0)
					strMsg += strMsmnts;

				strMsg += L".\n\n";
			}
			else
			{
				strMsg = L"You have selected the following:\n\n";

				if (1 == m_iStudyCount)
				{
					strMsg += CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_ONE_STUDY));
				}
				else if (1 < m_iStudyCount)
				{
					CString strItems;
					CString strText(CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_N_STUDIES)));
					strItems.Format(strText, m_iStudyCount);
					strMsg += strItems;
				}

				if (1 == m_iSeriesCount)
				{
					strMsg += CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_ONE_SERIES));					
				}
				else if (1 < m_iSeriesCount)
				{
					CString strItems;
					CString strText(CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_N_SERIES)));
					strItems.Format(strText, m_iSeriesCount);
					strMsg += strItems;
				}

				if (1 == m_iImageCount)
				{
					strMsg += CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_ONE_IMAGE));							
				}
				else if (1 < m_iImageCount)
				{
					CString strItems;
					CString strText(CString(MAKEINTRESOURCE(IDS_DELCONF_DELETE_N_IMAGES)));
					strItems.Format(strText, m_iImageCount);
					strMsg += strItems;
				}

				strMsg += L"\n";
			}

			//if (0 < m_iStudyCount || 0 < m_iSeriesCount)
			//{
			//	if (0 == iExported)  // None of them are archived
			//	{
			//		uiImage = VSI_THEME_IL_HEADER_WARNING;
			//		bShowCheckBox = TRUE;
			//		strMsg += L"The items selected have not yet been copied or have been modified since they were last copied. ";
			//	}
			//	else if (iSelected != iExported)  // Some are archived, some are not
			//	{
			//		uiImage = VSI_THEME_IL_HEADER_WARNING;
			//		bShowCheckBox = TRUE;
			//		strMsg += L"Some of the items selected have not yet been copied or have been modified since they were last copied. ";
			//	}
			//}

			strMsg += L"This action will permanently delete all information for the selected items from your system.\n";
		}

		strMsg += L"\nAre you sure you want to continue?";

		//CWindow wndAllowUnarch = GetDlgItem(IDC_DSC_ONLY_ARCHIVED);
		//wndAllowUnarch.ShowWindow(bShowCheckBox ? SW_SHOW : SW_HIDE);
		//wndAllowUnarch.SendMessage(
		//	BM_SETCHECK,
		//	m_bDeleteUnexportedItems ? BST_CHECKED : BST_UNCHECKED);

		CWindow wndMsg(GetDlgItem(IDC_DSC_MESSAGE));
		wndMsg.SendMessage(EM_SETWORDWRAPMODE, WBF_WORDBREAK);

		CHARFORMAT2 cf;
		ZeroMemory(&cf, sizeof(cf));
		cf.cbSize = sizeof(CHARFORMAT2);
		cf.dwMask = CFM_COLOR;
		cf.crTextColor = VSI_COLOR_L3GRAY;
		wndMsg.SendMessage(EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
		wndMsg.SendMessage(EM_REPLACESEL, FALSE, (LPARAM)(LPCWSTR)strMsg);
		wndMsg.SendMessage(EM_HIDESELECTION, 1);

		// Sets header and footer
		HIMAGELIST hilHeader(NULL);
		VsiThemeGetImageList(uiImage, &hilHeader);
		VsiThemeDialogBox(*this, TRUE, TRUE, hilHeader);
	}
	VSI_CATCH(hr);

	return 1;
}

/// <summary>
///	Clicked on the X. Assume they meant "No"
/// </summary>
LRESULT CVsiDeleteConfirmation::OnClose(UINT, WPARAM, LPARAM, BOOL&)
{
	EndDialog(IDNO);

	return 0;
}

/// <summary>
///	Clicked on Yes. Get the "delete only archived option" and close the
///	dialog
/// </summary>
LRESULT CVsiDeleteConfirmation::OnCmdYes(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Delete Confirmation - Yes");

	// Get the setting of the check mark for deleting only un-marked studies.
	//CWindow wndAllowUnarch = GetDlgItem(IDC_DSC_ONLY_ARCHIVED);
	//if (wndAllowUnarch.IsWindowVisible())
	//{
	//	LRESULT state = wndAllowUnarch.SendMessage(BM_GETCHECK, NULL, NULL);
	//	m_bDeleteUnexportedItems = (BST_CHECKED == state);
	//}

	EndDialog(IDYES);

	return 0;
}

/// <summary>
///	Clicked on No. Just close the dialog.
/// </summary>
LRESULT CVsiDeleteConfirmation::OnCmdNo(WORD, WORD, HWND, BOOL&)
{
	VSI_LOG_MSG(VSI_LOG_ACTIVITY, L"Delete Confirmation - No");

	EndDialog(IDNO);

	return 0;
}

int CVsiDeleteConfirmation::CheckForMeasurementsOnImage()
{
	HRESULT hr = S_OK;
	CComVariant vNS;
	int iNumMeasurements(0);

	try
	{
		for (int i = 0; i < m_iImageCount; ++i)
		{
			CComQIPtr<IVsiImage> pImage(m_ppImage[i]);
			VSI_CHECK_INTERFACE(pImage, VSI_LOG_ERROR, NULL);

			vNS.Clear();
			hr = pImage->GetProperty(VSI_PROP_IMAGE_ID, &vNS);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszFile;
			pImage->GetDataPath(&pszFile);

			PathRenameExtension(pszFile, L"." VSI_FILE_EXT_MEASUREMENT);

			// Open the measurement file and check if there are any measurements
			// with this image ID.
			CComPtr<IVsiMsmntFile> pMsmntFile;
			hr = pMsmntFile.CoCreateInstance(__uuidof(VsiMsmntFile));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,NULL);

			hr = pMsmntFile->Initialize(pszFile, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,NULL);

			int iCount(0);
			pMsmntFile->CountMeasurementsForImage(V_BSTR(&vNS), &iCount);
			iNumMeasurements += iCount;

			pImage.Release();
		}
	}
	VSI_CATCH(hr);

	return iNumMeasurements;
}
