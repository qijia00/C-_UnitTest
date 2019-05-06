/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudyMover.cpp
**
**	Description:
**		Implementation of CVsiStudyMover
**
*******************************************************************************/

#include "stdafx.h"
#pragma comment(lib, "Mpr.lib")
#include <VsiCommCtl.h>
#include <VsiCommCtlLib.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiImportExportXml.h>
#include <VsiParameterXml.h>
#include <VsiAnalRsltXmlTags.h>
#include <VsiLicenseUtils.h>
#include <VsiResProduct.h>
#include "VsiImpExpUtility.h"
#include "VsiStudyMover.h"

CVsiStudyMover::CVsiStudyMover() :
	m_bCancelCopy(FALSE),
	m_bAllowOverwrite(FALSE),
	m_bAutoDelete(FALSE),
	m_ullCopiedSoFar(0),
	m_ullTotalSize(0),
	m_dwStartTime(0),
	m_dwNextReportTime(0),
	m_pStudy(NULL)
{
}

CVsiStudyMover::~CVsiStudyMover()
{
}

STDMETHODIMP CVsiStudyMover::Initialize(
	HWND hParentWnd,
	const VARIANT* pvParam,
	ULONGLONG ullTotalSize,
	DWORD dwFlags)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pvParam, VSI_LOG_ERROR, NULL);
		VSI_VERIFY(VT_DISPATCH == V_VT(pvParam), VSI_LOG_ERROR, NULL);

		m_hWndOwner = hParentWnd;

		m_pXmlDoc = V_DISPATCH(pvParam);
		VSI_CHECK_INTERFACE(m_pXmlDoc, VSI_LOG_ERROR, NULL);

		m_bAllowOverwrite = (VSI_STUDY_MOVER_OVERWRITE == (dwFlags & VSI_STUDY_MOVER_OVERWRITE));
		m_bAutoDelete = (VSI_STUDY_MOVER_AUTODELETE == (dwFlags & VSI_STUDY_MOVER_AUTODELETE));
		m_ullTotalSize = ullTotalSize;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiStudyMover::Export(LPCOLESTR pszDestPath, LPCOLESTR pszCustomFolderName)
{
	m_Operation = VSI_SMO_EXPORTING;
	m_strDestinationPath = pszDestPath;
	m_strCustomFolderName = pszCustomFolderName;

	VsiProgressDialogBox(
		NULL, (LPCWSTR)MAKEINTRESOURCE(VSI_PROGRESS_TEMPLATE_MPC),
		m_hWndOwner, ExportImportCallback, this);

	return S_OK;
}

STDMETHODIMP CVsiStudyMover::Import(LPCOLESTR pszDestPath)
{
	m_Operation = VSI_SMO_IMPORTING;
	m_strDestinationPath = pszDestPath;
	m_strCustomFolderName.Empty();

	VsiProgressDialogBox(
		NULL, (LPCWSTR)MAKEINTRESOURCE(VSI_PROGRESS_TEMPLATE_MPC),
		m_hWndOwner, ExportImportCallback, this);

	return S_OK;
}

/// <summary>
///	This callback function is used by the CopyFileEx function to allow us
///	to track how well its doing. It will call this function after every
/// "Chunk" ~= 64K. We use this to update the progress dialog.
/// </summary>
static DWORD CALLBACK CopyProgressRoutine(
	LARGE_INTEGER TotalFileSize,  // total file size, in bytes
	LARGE_INTEGER TotalBytesTransferred,  // total number of bytes transferred
	LARGE_INTEGER StreamSize,  // total number of bytes for this stream
	LARGE_INTEGER StreamBytesTransferred,  // total number of bytes transferred for this stream
	DWORD dwStreamNumber,     // the current stream
	DWORD dwCallbackReason,   // reason for callback
	HANDLE hSourceFile,       // handle to the source file
	HANDLE hDestinationFile,  // handle to the destination file
	LPVOID lpData             // passed by CopyFileEx
)
{
	UNREFERENCED_PARAMETER(TotalFileSize);
	UNREFERENCED_PARAMETER(StreamSize);
	UNREFERENCED_PARAMETER(StreamBytesTransferred);
	UNREFERENCED_PARAMETER(dwStreamNumber);
	UNREFERENCED_PARAMETER(hSourceFile);
	UNREFERENCED_PARAMETER(hDestinationFile);

	switch (dwCallbackReason)
	{
	// This gets called every 64K of data copied - update the status dialog.
	case CALLBACK_CHUNK_FINISHED:
		{
			CVsiStudyMover *pStudyMover = (CVsiStudyMover*)lpData;

			// Update the progress bar to show approximately where we are
			ULONGLONG ullTotalSizeSoFar = pStudyMover->m_ullCopiedSoFar + TotalBytesTransferred.QuadPart;
			if (TotalFileSize.QuadPart == TotalBytesTransferred.QuadPart)
			{
				pStudyMover->m_ullCopiedSoFar += TotalFileSize.QuadPart;
			}

			pStudyMover->m_wndProgress.SendMessage(
				WM_VSI_PROGRESS_SETPOS, (int)(ullTotalSizeSoFar / VSI_ONE_MEGABYTE), 0);

			DWORD dwElapsedTime = GetTickCount() - pStudyMover->m_dwStartTime;
			if (dwElapsedTime > pStudyMover->m_dwNextReportTime)
			{
				WCHAR szProgressBuffer[256];

				double dbBytesPerMillisecond = (double)ullTotalSizeSoFar / (double)dwElapsedTime;

				// iTimeLeft in sec
				int iTimeLeft = (int)(((pStudyMover->m_ullTotalSize - ullTotalSizeSoFar) / dbBytesPerMillisecond) / 1000);

				if (iTimeLeft < 60)
				{
					if (iTimeLeft < 0)
					{
						iTimeLeft = 1;
					}
					else
					{
						// Round it to 5sec
						float fTime = ceil((float)iTimeLeft / 5.0f);
						iTimeLeft = (int)fTime * 5;
						if (iTimeLeft <= 0)
							iTimeLeft = 1;
					}

					swprintf_s(szProgressBuffer, _countof(szProgressBuffer), L"Copying data: About %d seconds remaining", iTimeLeft);

					// We'll update the estimates time every 4 seconds.
					pStudyMover->m_dwNextReportTime += 4000;
				}
				else if (iTimeLeft < 80)
				{
					swprintf_s(szProgressBuffer, _countof(szProgressBuffer), L"Copying data: About 1 minute remaining");

					// We'll update the estimates time every 10 seconds.
					pStudyMover->m_dwNextReportTime += 10000;
				}
				else if (iTimeLeft < 120)
				{
					swprintf_s(szProgressBuffer, _countof(szProgressBuffer), L"Copying data: About 2 minutes remaining");

					// We'll update the estimates time every 30 seconds.
					pStudyMover->m_dwNextReportTime += 30000;
				}
				else
				{
					swprintf_s(szProgressBuffer, _countof(szProgressBuffer), L"Copying data: About %d minutes remaining", (iTimeLeft + 30) / 60);

					// We'll update the estimates time every 30 seconds.
					pStudyMover->m_dwNextReportTime += 30000;
				}

				if (!pStudyMover->m_bCancelCopy)
				{
					pStudyMover->m_wndProgress.SendMessage(WM_VSI_PROGRESS_SETMSG, 0, (LPARAM)szProgressBuffer);
				}
			}

			// Check if the user pressed the Cancel button on the progress dialog.
			if (pStudyMover->m_bCancelCopy)
				return PROGRESS_CANCEL;
		}
		break;

	case CALLBACK_STREAM_SWITCH:
		// This gets called when the copy first starts. Not really anything we need to do here.
		break;
	}

	return PROGRESS_CONTINUE;
}

void CVsiStudyMover::ConvertToExtendedLengthPath(CString &strPath)
{
	if (2 < strPath.GetLength())
	{
		if (L'?' != strPath.GetAt(2))  // Not in extended length path format
		{
			if (L'\\' == strPath.GetAt(0) && L'\\' == strPath.GetAt(1))
			{
				CString strUnc(L"\\\\?\\UNC");
				strUnc += strPath.GetString() + 1;  // Skip 1 '\'
				strPath = strUnc;
			}
			else if (PathIsNetworkPath(strPath))
			{
				WCHAR pszDrive[3];
				lstrcpyn(pszDrive, strPath, 3);

				WCHAR pszRemotePath[1024];
				*pszRemotePath = 0;

				DWORD dwPathLength = _countof(pszRemotePath);
				WNetGetConnection(pszDrive, pszRemotePath, &dwPathLength);

				int iStartIndex = 2;
				if (strPath.GetAt(3) == L'\\') // in case we have 2 '\\'
				{
					iStartIndex = 3;
				}

				CString strUnc;
				strUnc.Format(L"\\\\?\\UNC%s%s", pszRemotePath + 1, strPath.GetString() + iStartIndex);
				strPath = strUnc;
			}
			else
			{
				strPath.Insert(0, L"\\\\?\\");
			}
		}
	}
}

/// <summary>
///	Copy all of the files that make up the study. The target directory
///	should have already been built. Space and access should already have
/// been verified.
/// </summary>
int CVsiStudyMover::CopyWithProgress()
{
	HRESULT hr = S_OK;
	WCHAR szDestinationPath[VSI_MAX_PATH];
	szDestinationPath[0] = 0;
	HANDLE hFind = INVALID_HANDLE_VALUE;
	BOOL bOverwrite = FALSE, bTargetLocked = FALSE, bLocked = FALSE;

	try
	{
		// Record the starting time of the process. We will need this to compute
		// estimated completion later.
		m_dwStartTime = GetTickCount();
		m_dwNextReportTime = 5000;		// The first report will be after 5 seconds of copying

		CComHeapPtr<OLECHAR> pszPath; // Source Study path
		CComVariant vAttr, vStudyPath;
		CComPtr<IXMLDOMNodeList> pListStudy;
		hr = m_pXmlDoc->selectNodes(
			CComBSTR(L"//" VSI_STUDY_XML_ELM_STUDIES L"/" VSI_STUDY_XML_ELM_STUDY),
			&pListStudy);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting studies");

		long lStudyCount = 0;
		hr = pListStudy->get_length(&lStudyCount);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

		for (long i = 0; i < lStudyCount; i++)
		{
			CComPtr<IXMLDOMNode> pNode;
			hr = pListStudy->get_item(i, &pNode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting studies");

			CComQIPtr<IXMLDOMElement> pStudyElement(pNode);

			hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_NAME), &vAttr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_NAME));

			if (lstrlen(V_BSTR(&vAttr)) > 0)
			{
				CString strMsg;
				CString strText(MAKEINTRESOURCE(IDS_STYMVR_COPYING_STUDY_NAME));
				strMsg.Format(strText, V_BSTR(&vAttr));
				m_wndProgress.SetWindowText(strMsg);
			}
			else
			{
				m_wndProgress.SetWindowText(CString(MAKEINTRESOURCE(IDS_STYMVR_COPYING_STUDY)));
			}

			// Overwrite?
			bOverwrite = FALSE;

			vAttr.Clear();
			hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_OVERWRITE), &vAttr);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_OVERWRITE));

			if (S_OK == hr)
			{
				vAttr.ChangeType(VT_BOOL);
				bOverwrite = VARIANT_FALSE != V_BOOL(&vAttr);
			}

			// If the current study would cause an overwrite and the overwrite flag is not
			// set then skip this study.
			if (!m_bAllowOverwrite && bOverwrite)
			{
				hr = pStudyElement->setAttribute(
					CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
					CComVariant(VSI_SEI_COPY_FAILED));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pStudyElement->setAttribute(
					CComBSTR(VSI_STUDY_XML_ATT_FAILURE_CODE),
					CComVariant(VSI_SEI_FAILED_OVERWRITE_DISABLED));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				continue;
			}

			if (VSI_SMO_IMPORTING == m_Operation)
			{
				// Target locked?
				bTargetLocked = FALSE;

				vAttr.Clear();
				hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_LOCKED), &vAttr);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_LOCKED));

				if (S_OK == hr)
				{
					vAttr.ChangeType(VT_BOOL);
					bTargetLocked = VARIANT_FALSE != V_BOOL(&vAttr);
				}

				// Skip over locked studies on import.
				if (bOverwrite && bTargetLocked)
				{
					hr = pStudyElement->setAttribute(
						CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
						CComVariant(VSI_SEI_COPY_SKIP));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					hr = pStudyElement->setAttribute(
						CComBSTR(VSI_STUDY_XML_ATT_FAILURE_CODE),
						CComVariant(VSI_SEI_FAILED_TARGET_LOCKED));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					continue;
				}
			}

			// Create the destination folder where the study will be copied to. If we are
			// exporting then we want a directory name that tells the average human reader
			// what stored here. If we are importing or copying then we need a directory
			// name compatible with the application
			if (!CreateDestination(pStudyElement, szDestinationPath, _countof(szDestinationPath)))
			{
				hr = pStudyElement->setAttribute(
					CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
					CComVariant(VSI_SEI_COPY_FAILED));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				hr = pStudyElement->setAttribute(
					CComBSTR(VSI_STUDY_XML_ATT_FAILURE_CODE),
					CComVariant(VSI_SEI_FAILED_DISK_ERROR));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				return VSI_SEI_COPY_FAILED;
			}

			// Loop through each file in the source directory and copy it to the destination
			WIN32_FIND_DATA ff;
			CString strSourceName;
			CString strDestName;

			// Get study path
			hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_PATH), &vStudyPath);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_PATH));
			LPCWSTR pszPath = V_BSTR(&vStudyPath);

			// Construct a string to represent the base path to use for
			// finding all of the files in the study.
			CString strSourceFolder;
			strSourceFolder.Format(L"%s\\*.*", pszPath);
			ConvertToExtendedLengthPath(strSourceFolder);

			hFind = FindFirstFile(strSourceFolder, &ff);
			BOOL bWorking = (INVALID_HANDLE_VALUE != hFind);

			// if failed to find this file off the batch, mark it as failed and continue forward
			if (!bWorking)
			{
				hr = pStudyElement->setAttribute(
					CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
					CComVariant(VSI_SEI_COPY_FAILED));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			while (bWorking)	// For each file in the study.
			{
				// Skip the dots, include subdirectories.
				if (0 == ff.cFileName[0] ||	L'.' == ff.cFileName[0])
				{
					NULL;
				}
				else if (0 == (ff.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
				{
					// File copy

					strSourceName.Format(L"%s\\%s", pszPath, ff.cFileName);
					strDestName.Format(L"%s\\%s", szDestinationPath, ff.cFileName);

					VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Copying file: from %s", strSourceName.GetString()));
					VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Copying file: to %s", strDestName.GetString()));

					ConvertToExtendedLengthPath(strSourceName);
					ConvertToExtendedLengthPath(strDestName);

					BOOL bRtn = CopyFileEx(
						strSourceName, strDestName, CopyProgressRoutine,
						(LPVOID)this, &m_bCancelCopy, 0);

					if (FALSE == bRtn)
					{
						VSI_WIN32_VERIFY(bRtn, VSI_LOG_WARNING, L"Failed to copy file");
						if (m_bCancelCopy)	// The user canceled the copy.
						{
							hr = pStudyElement->setAttribute(
								CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
								CComVariant(VSI_SEI_COPY_CANCELED));
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
						else  // The copy failed
						{
							hr = pStudyElement->setAttribute(
								CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
								CComVariant(VSI_SEI_COPY_FAILED));
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}

						FindClose(hFind);
						hFind = INVALID_HANDLE_VALUE;

						if (!RollbackStudyCopy(szDestinationPath))
						{
							hr = pStudyElement->setAttribute(
								CComBSTR(VSI_STUDY_XML_ATT_ROLLBACK_STATUS),
								CComVariant(CString(MAKEINTRESOURCE(IDS_STYMVR_ROLLBACK_OLD_STUDY_FAILED))));
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
						}
					}

					// If we are importing then make sure the write protect bit is cleared so the
					// user will be able to edit it later,
					if (VSI_SMO_IMPORTING == m_Operation)
					{
						DWORD dwAttributes = GetFileAttributes(strDestName);
						if (FILE_ATTRIBUTE_READONLY & dwAttributes)
						{
							dwAttributes &= ~FILE_ATTRIBUTE_READONLY;
							SetFileAttributes(strDestName, dwAttributes);
						}
					}
				}
				else
				{
					// Folder copy
					if (!CopySeriesFolder(pStudyElement, ff.cFileName, pszPath, szDestinationPath))
					{
						// Exit on failure
						FindClose(hFind);
						hFind = INVALID_HANDLE_VALUE;

						if (!RollbackStudyCopy(szDestinationPath))
						{
							hr = pStudyElement->setAttribute(
								CComBSTR(VSI_STUDY_XML_ATT_ROLLBACK_STATUS),
								CComVariant(CString(MAKEINTRESOURCE(IDS_STYMVR_ROLLBACK_OLD_STUDY_FAILED))));
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,
								VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_ROLLBACK_STATUS));
						}
					}
				}

				CComVariant vStatus;
				hr = pStudyElement->getAttribute(
					CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
					&vStatus);
				if (S_OK != hr)
				{
					hr = pStudyElement->setAttribute(
						CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
						CComVariant(VSI_SEI_COPY_OK));
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,
						VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));
				}

				bWorking = FindNextFile(hFind, &ff);
			}

			// Done with the current directory. Close the handle and return.
			if (INVALID_HANDLE_VALUE != hFind)
			{
				FindClose(hFind);
				hFind = INVALID_HANDLE_VALUE;
			}

			if (VSI_SMO_IMPORTING == m_Operation)
			{
				if (bOverwrite)
				{
					// Copy series from tmp to target study

					CComVariant vTarget;
					hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_TARGET), &vTarget);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					CString strTempStudyFolder(V_BSTR(&vTarget));

					// Find study folder being replaced
					LPWSTR pszTempStudyFolder = strTempStudyFolder.GetBuffer();
					LPWSTR pszSpt = wcsrchr(pszTempStudyFolder, L'.');
					if (pszSpt != NULL)
					{
						CString strOrigStudyFolder(pszTempStudyFolder, pszSpt - pszTempStudyFolder);
						strTempStudyFolder.ReleaseBuffer();

						// Get data path from created Temp folder path
						pszSpt = wcsrchr(szDestinationPath, L'\\');
						if (NULL != pszSpt)
						{
							CString strDataPath(szDestinationPath, (int)(pszSpt - szDestinationPath));

							CString strLocalStudyFolder;
							strLocalStudyFolder.Format(L"%s\\%s", strDataPath.GetString(), strOrigStudyFolder.GetString());

							// Copy series from tmp to target study
							if (CopySeriesFromTempToTarget(szDestinationPath, strLocalStudyFolder.GetString()))
							{
								hr = pStudyElement->setAttribute(
									CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
									CComVariant(VSI_SEI_COPY_OK));
								VSI_CHECK_HR(hr, VSI_LOG_ERROR,
									VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));
							}
							else
							{
								hr = pStudyElement->setAttribute(
									CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
									CComVariant(VSI_SEI_COPY_FAILED));
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								hr = pStudyElement->setAttribute(
									CComBSTR(VSI_STUDY_XML_ATT_FAILURE_CODE),
									CComVariant(VSI_SEI_FAILED_DISK_ERROR));
								VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

								return VSI_SEI_COPY_FAILED;
							}
						}
					}
				}
			}

			// The study was copied successfully. If we had to create an overwrite roll back directory
			// then delete it now.
			if (!RemoveRollbackDirectory(szDestinationPath))
			{
				// This case is handled properly next time we try to create a rollback
			}

			// If we are exporting and the user has selected to auto-delete then this is where
			// we do it.
			if (VSI_SMO_EXPORTING == m_Operation)
			{
				// Put a copy of the measurement package definition file associated with the study
				// into the target directory so with will be available when the study is imported
				// to a work station.
				//_stprintf(szSourceName, _T("%s\\%s"), CStudyPaths::GetSystemFilesDirectory(), it->m_szMeasurementFile);
				//_stprintf(szDestName, _T("%s\\%s"), szDestinationPath, it->m_szMeasurementFile);
				//CopyFile(szSourceName, szDestName, FALSE);

				vAttr.Clear();
				hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_LOCKED), &vAttr);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_LOCKED));

				bLocked = FALSE;
				if (S_OK == hr)
				{
					bLocked = (0 == lstrcmpi(VSI_VAL_TRUE, V_BSTR(&vAttr)));
				}

				// Delete not locked study
				if (m_bAutoDelete && !bLocked)
				{
					CString strPath(pszPath);
					ConvertToExtendedLengthPath(strPath);

					if (VsiRemoveDirectory(strPath))
					{
						hr = pStudyElement->setAttribute(
							CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
							CComVariant(VSI_SEI_COPY_OK_AUTO_DEL));
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
					else
					{
						hr = pStudyElement->setAttribute(
							CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
							CComVariant(VSI_SEI_COPY_OK_AUTO_DEL_FAILED));
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}
			}
		}
	}
	VSI_CATCH(hr);

	return 1;
}

/// <summary>
///	Create the destination study folder.
/// </summary>
BOOL CVsiStudyMover::CreateDestination(
	IXMLDOMElement *pStudyElement,
	LPWSTR pszDestinationPath, int iPathLength)
{
	BOOL bRet = TRUE;

	HRESULT hr = S_OK;

	try
	{
		// Determine whether folder name is user provided or constructed by default
		if (!m_strCustomFolderName.IsEmpty())
		{
			swprintf_s(pszDestinationPath, iPathLength, L"%s\\%s",
				m_strDestinationPath, m_strCustomFolderName);
		}
		else
		{
			CComVariant vTarget;
			hr = pStudyElement->getAttribute(CComBSTR(VSI_STUDY_XML_ATT_TARGET), &vTarget);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure getting attribute (%s)", VSI_STUDY_XML_ATT_TARGET));

			swprintf_s(pszDestinationPath, iPathLength, L"%s\\%s",
				m_strDestinationPath, (LPWSTR)V_BSTR(&vTarget));
		}

		// Try to create the directory. If we succeeded then we are done.
		if (CreateDirectory(pszDestinationPath, NULL))
			return TRUE;

		// If it failed for any reason other than it already exists then we'll give up now.
		DWORD err = GetLastError();
		if (ERROR_ALREADY_EXISTS != err)
		{
			hr = pStudyElement->setAttribute(
				CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
				CComVariant(VSI_SEI_COPY_FAILED));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,
				VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));

			return FALSE;
		}

		if (VSI_SMO_IMPORTING == m_Operation)
		{
			// If we are here then the target directory exists.
			// Let's try to rename it and then retry the create.
			CString strRenamePath;

			// Delete old rollback
			if (!RemoveRollbackDirectory(pszDestinationPath))
			{
				hr = pStudyElement->setAttribute(
					CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
					CComVariant(VSI_SEI_COPY_FAILED));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));

				return FALSE;
			}

			CreateRollbackName(pszDestinationPath, strRenamePath);
			// Create rollback
			if (!MoveFile(pszDestinationPath, strRenamePath))
			{
				hr = pStudyElement->setAttribute(
					CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
					CComVariant(VSI_SEI_COPY_FAILED));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));

				return FALSE;
			}

			// Now try the create again. If it fails this time then we probably have a serious
			// problem with the drive. Best to give up now.
			if (!CreateDirectory(pszDestinationPath, NULL))
			{
				// Logs failed reasons
				VSI_LOG_MSG(VSI_LOG_WARNING,
					VsiFormatMsg(L"CreateDirectory failed: %s", pszDestinationPath));
				CVsiException e(
					HRESULT_FROM_WIN32(GetLastError()),
					VSI_LOG_WARNING,
					L"CreateDirectory failed",
					__FILE__,
					__LINE__);
				VSI_ERROR_LOG(e);

				hr = pStudyElement->setAttribute(
					CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
					CComVariant(VSI_SEI_COPY_FAILED));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));

				return FALSE;
			}
		}
	}
	VSI_CATCH(hr);

	return bRet;
}

/// <summary>
///	Create the path name for the rollback directory used during overwrites of
///	existing studies. Appends a known extension to the existing path.
/// </summary>
/// <param name="pszDestinationPath">Pointer to the study's path</param>
/// <param name="pszRollbackName">Pointer to the buffer for the rollback path</param>
/// <param name="iPathLength">pszRollbackName length</param>
BOOL CVsiStudyMover::CreateRollbackName(
	LPCWSTR pszDestinationPath,
	CString &strRollbackName)
{
	strRollbackName.Format(L"%s.Rollback", pszDestinationPath);

	return TRUE;
}

/// <summary>
///	Create the destination folder for individual series
///	Copy all files into this folder
/// </summary>
/// <param name="pszDestinationPath">Pointer to the series source path (not including trailing slash)</param>
/// <param name="rollbackName">Pointer to the destination folder path. Assumed to be VSI_MAX_PATH long</param>
/// <returns>TRUE on success, FALSE on failure</returns>
BOOL CVsiStudyMover::CopySeriesFolder(
	IXMLDOMElement *pStudyElement,
	LPCWSTR pszSieresFolder,
	LPCWSTR pszSourcePath,
	LPCWSTR pszDestinationPath)
{
	HRESULT	hr = S_OK;
	HANDLE hFind = INVALID_HANDLE_VALUE;
	WIN32_FIND_DATA ff;
	CString strDestFolder;
	CString strDestFolderFinal;
	CString strDestFolderDel;

	BOOL bCopySeriesFailed = FALSE;
	try
	{
		long lSeriesCount(0);
		long lImageCount(0);

		if (VSI_SMO_IMPORTING == m_Operation)
		{
			if ((m_pStudy == NULL) || (!m_pStudy.IsEqualObject(pStudyElement)))
			{
				m_pStudy = pStudyElement;
			}
		}

		// Get series count
		CComPtr<IXMLDOMNodeList> pListSeries;
		hr = pStudyElement->selectNodes(CComBSTR(VSI_SERIES_XML_ELM_SERIES), &pListSeries);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting series");

		hr = pListSeries->get_length(&lSeriesCount);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting series");

		BOOL bCopy = FALSE;
		CComPtr<IXMLDOMNode> pSeriesNode;

		if (0 < lSeriesCount)
		{
			CString strFind;
			strFind.Format(VSI_SERIES_XML_ELM_SERIES L"[@" VSI_SERIES_XML_ATT_FOLDER_NAME L"=\'%s\']", pszSieresFolder);

			hr = pStudyElement->selectSingleNode(CComBSTR(strFind), &pSeriesNode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting series");

			bCopy = S_OK == hr;

			if (bCopy && pSeriesNode != NULL)
			{
				CComPtr<IXMLDOMNodeList> pListImages;
				hr = pSeriesNode->selectNodes(CComBSTR(VSI_IMAGE_XML_ELM_IMAGE), &pListImages);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting images");

				if (S_OK == hr)
				{
					hr = pListImages->get_length(&lImageCount);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR,	L"Failure getting image count");
				}
			}
		}

		if (bCopy)
		{
			// Create destination folder for the series
			CString strSourceFolder;

			strSourceFolder.Format(L"%s\\%s", pszSourcePath, pszSieresFolder);
			strDestFolder.Format(L"%s\\%s", pszDestinationPath, pszSieresFolder);

			ConvertToExtendedLengthPath(strDestFolder);

			if (INVALID_FILE_ATTRIBUTES != GetFileAttributes(strDestFolder))
			{
				strDestFolderFinal = strDestFolder;
				strDestFolder += L".tmp";
			}

			if (!CreateDirectory(strDestFolder, NULL))
			{
				// Logs failed reasons
				VSI_LOG_MSG(VSI_LOG_WARNING,
					VsiFormatMsg(L"CreateDirectory failed: %s", strDestFolder.GetString()));
				CVsiException e(
					HRESULT_FROM_WIN32(GetLastError()),
					VSI_LOG_WARNING,
					L"CreateDirectory failed",
					__FILE__,
					__LINE__);
				VSI_ERROR_LOG(e);

				hr = pStudyElement->setAttribute(
					CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
					CComVariant(VSI_SEI_COPY_FAILED));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,
					VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));

				return FALSE;
			}

			CString strSourceName;
			CString strDestName;

			// Construct a string to represent the base path to use for
			// finding all of the files in the series.
			CString strFind;
			strFind.Format(L"%s\\*.*", strSourceFolder.GetString());

			ConvertToExtendedLengthPath(strFind);

			hFind = FindFirstFile(strFind, &ff);
			BOOL bWorking = (INVALID_HANDLE_VALUE != hFind);

			while (bWorking)	// For each file in the study.
			{
				// Skip the dots and subdirectories.
				if (0 == ff.cFileName[0] ||
					L'.' == ff.cFileName[0] ||
					(0 != (ff.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)))
				{
					NULL;
				}
				else
				{
					strSourceName.Format(L"%s\\%s", strSourceFolder.GetString(), ff.cFileName);
					strDestName.Format(L"%s\\%s", strDestFolder.GetString(), ff.cFileName);

					bCopy = FALSE;

					// Check common files
					CString ppszMustCopy[] = {
						VSI_FILE_MEASUREMENT,
						CString(MAKEINTRESOURCE(IDS_SERIES_FILE_NAME))
					};

					for (int i = 0; i < _countof(ppszMustCopy); ++i)
					{
						if (0 == lstrcmp(ppszMustCopy[i], ff.cFileName))
						{
							bCopy = TRUE;
							break;
						}
					}

					if (!bCopy)
					{
						if (NULL != wcsstr(ff.cFileName, L".tiff") || NULL != wcsstr(ff.cFileName, L".Tif"))
						{
							bCopy = TRUE;
						}
					}

					if (!bCopy)
					{
						if (0 == lImageCount)
						{
							bCopy = TRUE;
						}
						else
						{
							_ASSERT(NULL != pSeriesNode);  // Shouldn't happen

							WCHAR szImageName[MAX_PATH];
							wcscpy_s(szImageName, _countof(szImageName), ff.cFileName);

							LPWSTR pszDot = wcschr(szImageName, L'.');
							if (NULL != pszDot)
							{
								*pszDot = 0;
							}

							CString strFind;
							strFind.Format(VSI_IMAGE_XML_ELM_IMAGE L"[@" VSI_IMAGE_XML_ATT_NAME L"=\'%s\']", szImageName);

							CComPtr<IXMLDOMNode> pImageNode;
							hr = pSeriesNode->selectSingleNode(CComBSTR(strFind), &pImageNode);
							VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting series");

							bCopy = S_OK == hr;
						}
					}

					if (bCopy)
					{
						VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Copying file: from %s", strSourceName.GetString()));
						VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"Copying file: to %s", strDestName.GetString()));

						ConvertToExtendedLengthPath(strSourceName);
						ConvertToExtendedLengthPath(strDestName);

						BOOL bRtn = CopyFileEx(
							strSourceName, strDestName, CopyProgressRoutine,
							(LPVOID)this, &m_bCancelCopy, 0);

						if (FALSE == bRtn)
						{
							if (!m_bCancelCopy)	// The copy failed
							{
								hr = pStudyElement->setAttribute(
									CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
									CComVariant(VSI_SEI_COPY_FAILED));
								VSI_CHECK_HR(hr, VSI_LOG_ERROR,
									VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));
							}
							else  // The user canceled the copy.
							{
								hr = pStudyElement->setAttribute(
									CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
									CComVariant(VSI_SEI_COPY_CANCELED));
								VSI_CHECK_HR(hr, VSI_LOG_ERROR,
									VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));
							}

							bCopySeriesFailed = TRUE;
						}

						// If we are importing then make sure the write protect bit is cleared so the
						// user will be able to edit it later,
						if (VSI_SMO_IMPORTING == m_Operation)
						{
							DWORD dwAttributes = GetFileAttributes(strDestName);
							if (INVALID_FILE_ATTRIBUTES != dwAttributes)
							{
								if (FILE_ATTRIBUTE_READONLY & dwAttributes)
								{
									dwAttributes &= ~FILE_ATTRIBUTE_READONLY;
									SetFileAttributes(strDestName, dwAttributes);
								}
							}
						}
					}
				}

				bWorking = FindNextFile(hFind, &ff);
			}

			if (!bCopySeriesFailed)
			{
				FindClose(hFind);
				hFind = INVALID_HANDLE_VALUE;

				if (strDestFolderFinal.IsEmpty())
				{
					// Don't rollback
					strDestFolder.Empty();
				}
				else
				{
					strDestFolderDel = strDestFolderFinal + L".del";

					ConvertToExtendedLengthPath(strDestFolder);
					ConvertToExtendedLengthPath(strDestFolderFinal);
					ConvertToExtendedLengthPath(strDestFolderDel);

					BOOL bRet = MoveFile(strDestFolderFinal, strDestFolderDel);
					if (bRet)
					{
						// Rename temp folder to final name
						bRet = MoveFile(strDestFolder, strDestFolderFinal);
						if (bRet)
						{
							// Don't rollback
							strDestFolder.Empty();
						}
						else
						{
							hr = pStudyElement->setAttribute(
								CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
								CComVariant(VSI_SEI_COPY_FAILED));
							VSI_CHECK_HR(hr, VSI_LOG_ERROR,
								VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));
						}
					}
					else
					{
						hr = pStudyElement->setAttribute(
							CComBSTR(VSI_STUDY_XML_ATT_COPY_STATUS),
							CComVariant(VSI_SEI_COPY_FAILED));
						VSI_CHECK_HR(hr, VSI_LOG_ERROR,
							VsiFormatMsg(L"Failure setting attribute (%s)", VSI_STUDY_XML_ATT_COPY_STATUS));
					}
				}
			}
		}
	}
	VSI_CATCH(hr);

	// Done with the current directory. Close the handle and return.
	if (INVALID_HANDLE_VALUE != hFind)
		FindClose(hFind);

	if (!strDestFolder.IsEmpty())
	{
		ConvertToExtendedLengthPath(strDestFolder);

		BOOL bRet = VsiRemoveDirectory(strDestFolder);
		if (!bRet)
		{
			// Not critical
			VSI_LOG_MSG(
				VSI_LOG_WARNING,
				VsiFormatMsg(
					L"Failed to remove series folder - %s",
					strDestFolder.GetString()));
		}
	}

	if (!strDestFolderDel.IsEmpty())
	{
		ConvertToExtendedLengthPath(strDestFolderDel);

		BOOL bRet = VsiRemoveDirectory(strDestFolderDel);
		if (!bRet)
		{
			// Not critical
			VSI_LOG_MSG(
				VSI_LOG_WARNING,
				VsiFormatMsg(
					L"Failed to remove series folder - %s",
					strDestFolderDel.GetString()));
		}
	}

	if (bCopySeriesFailed)
		return FALSE;

	return TRUE;
}

/// <summary>
///	Copies series folders from original location ( the one to be overwritten) to temporary
///	location excluding already copied series
/// </summary>
BOOL CVsiStudyMover::CopySeriesFromTempToTarget(
	LPCWSTR pszSourcePath,
	LPCWSTR pszDestPath)
{
	HRESULT	hr = S_OK;
	HANDLE hFind = INVALID_HANDLE_VALUE;

	try
	{
		// Construct a string to represent the base path to use for
		// finding all of the files in the study.
		CString strSourceFolder;
		strSourceFolder.Format(L"%s\\*.*", pszSourcePath);

		WIN32_FIND_DATA ff;
		hFind = FindFirstFile(strSourceFolder, &ff);
		BOOL bWorking = (INVALID_HANDLE_VALUE != hFind);

		while (bWorking)	// For each file in the study.
		{
			// Skip the dots, include subdirectories.
			if (0 == ff.cFileName[0] ||	L'.' == ff.cFileName[0])
			{
				NULL;
			}
			else if (0 < (ff.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			{
				// 1. Backup old series to .del if exists
				// 2. Move tmp series to target location
				// 3. If failed, restore .del
				// 4. If succeeded, delete .del

				CString strSourceFolder;
				CString strDestFolder;
				CString strSeriesFolderDel;

				strSourceFolder.Format(L"%s\\%s", pszSourcePath, ff.cFileName);
				strDestFolder.Format(L"%s\\%s", pszDestPath, ff.cFileName);

				// Backup old
				ConvertToExtendedLengthPath(strDestFolder);

				if (INVALID_FILE_ATTRIBUTES != GetFileAttributes(strDestFolder))
				{
					strSeriesFolderDel = strDestFolder + L".del";

					ConvertToExtendedLengthPath(strSeriesFolderDel);

					// Remove directory if exist
					if (VsiGetFileExists(strSeriesFolderDel))
					{
						VsiRemoveDirectory(strSeriesFolderDel);
						Sleep(400);
					}

					BOOL bRet = MoveFile(strDestFolder, strSeriesFolderDel);
					if (!bRet)
					{
						//DWORD err = GetLastError();
						// Exit on failure
						FindClose(hFind);
						hFind = INVALID_HANDLE_VALUE;

						return FALSE;
					}
				}

				// Move series folder
				{
					BOOL bRet = TRUE;

					// Try 2 sec incase the files are still lock
					for (int i = 0; i < 5; ++i)
					{
						bRet = MoveFile(strSourceFolder, strDestFolder);
						if (bRet)
							break;

						Sleep(400);
					}

					if (!bRet)  // FAILED!!!
					{
						// Restore backup - Try 2 sec incase the files are still lock
						for (int i = 0; i < 5; ++i)
						{
							bRet = MoveFile(strSeriesFolderDel, strDestFolder);
							if (bRet)
								break;

							Sleep(400);
						}

						if (!bRet)
						{
							// Very bad - need better log on report
							VSI_LOG_MSG(
								VSI_LOG_WARNING,
								VsiFormatMsg(
									L"Failed to restore series folder - %s",
									strSeriesFolderDel.GetString()));
						}

						// Exit on failure
						FindClose(hFind);
						hFind = INVALID_HANDLE_VALUE;

						return FALSE;
					}
				}

				// Delete the ".del" folder
				if (!strSeriesFolderDel.IsEmpty())
				{
					ConvertToExtendedLengthPath(strSeriesFolderDel);

					BOOL bRet = VsiRemoveDirectory(strSeriesFolderDel);
					if (!bRet)
					{
						VSI_LOG_MSG(
							VSI_LOG_WARNING,
							VsiFormatMsg(
								L"Failed to delete overwritten study folder - %s",
								strSeriesFolderDel.GetString()));
					}
				}
			}

			bWorking = FindNextFile(hFind, &ff);
		}
	}
	VSI_CATCH(hr);

	// Done with the current directory. Close the handle and return.
	if (INVALID_HANDLE_VALUE != hFind)
	{
		FindClose(hFind);
	}

	return TRUE;
}

/// <summary>
///	The user has either canceled the operation or there was an error. This
///	function will try to cleanup any partially copied studies and perform a
///	rollback on a study that was being overwritten.
/// </summary>
BOOL CVsiStudyMover::RollbackStudyCopy(LPCWSTR pszDestinationPath)
{
	BOOL bRtnCode(TRUE);

	if (VSI_SMO_EXPORTING == m_Operation || VSI_SMO_IMPORTING == m_Operation)
	{
		BOOL bRet(TRUE);

		CString strDestinationPath(pszDestinationPath);

		ConvertToExtendedLengthPath(strDestinationPath);

		// If we were overwriting an existing study then
		// rollback the old directory name.

		// Don't support extended-length path for local study
		CString strRenamePath;
		CreateRollbackName(strDestinationPath, strRenamePath);

		ConvertToExtendedLengthPath(strRenamePath);

		if (INVALID_FILE_ATTRIBUTES != GetFileAttributes(strRenamePath))
		{
			// Tries 5sec to remove the folder
			for (int i = 0; i < 10; ++i)
			{
				bRet = VsiRemoveDirectory(strDestinationPath);
				if (bRet)
				{
					break;
				}
				Sleep(500);
			}

			if (bRet)
			{
				ConvertToExtendedLengthPath(strDestinationPath);

				if (!MoveFile(strRenamePath, strDestinationPath))
				{
					// Logs failed reasons
					VSI_LOG_MSG(VSI_LOG_WARNING,
						VsiFormatMsg(L"MoveFile failed: existing - %s", strRenamePath));
					VSI_LOG_MSG(VSI_LOG_WARNING,
						VsiFormatMsg(L"MoveFile failed: new - %s", strDestinationPath));
					CVsiException e(
						HRESULT_FROM_WIN32(GetLastError()),
						VSI_LOG_WARNING,
						L"MoveFile failed",
						__FILE__,
						__LINE__);
					VSI_ERROR_LOG(e);

					bRet = FALSE;
				}
			}
			else
			{
				// Logs failed reasons
				VSI_LOG_MSG(VSI_LOG_WARNING,
					VsiFormatMsg(L"RemoveDirectory failed: %s", strDestinationPath));
				CVsiException e(
					HRESULT_FROM_WIN32(GetLastError()),
					VSI_LOG_WARNING,
					L"RemoveDirectory failed",
					__FILE__,
					__LINE__);
				VSI_ERROR_LOG(e);
			}
		}

		if (!bRet)
		{
			//caller function will display proper error message
			bRtnCode = FALSE;
		}
	}

	return bRtnCode;
}

/// <summary>
///	Remove the rollback directory for the specified study path. If there is no
///	rollback then nothing is done.
/// </summary>
/// <param name="pszDestinationPath">Pointer to the study's path</param>
/// <returns>Returns a pointer to the directory built</returns>
BOOL CVsiStudyMover::RemoveRollbackDirectory(
	LPCWSTR pszDestinationPath)
{
	BOOL bRtnCode(TRUE);

	CString strRenamePath;
	CreateRollbackName(pszDestinationPath, strRenamePath);

	ConvertToExtendedLengthPath(strRenamePath);

	if (INVALID_FILE_ATTRIBUTES != GetFileAttributes(strRenamePath))
	{
		// It exists. Try to remove it.
		if (!VsiRemoveDirectory(strRenamePath))
		{
			// Logs failed reasons
			VSI_LOG_MSG(VSI_LOG_WARNING,
				VsiFormatMsg(L"RemoveDirectory failed: %s", strRenamePath.GetString()));
			CVsiException e(
				HRESULT_FROM_WIN32(GetLastError()),
				VSI_LOG_WARNING,
				L"RemoveDirectory failed",
				__FILE__,
				__LINE__);
			VSI_ERROR_LOG(e);

			bRtnCode = FALSE;
		}
	}

	return bRtnCode;
}

/// <summary>
/// Thread function to launch the copy process.
/// </summary>
UINT __stdcall CVsiStudyMover::StudyMoverThreadProc(void* pParam)
{
	HRESULT hr;
	CVsiStudyMover *pStudyMover = (CVsiStudyMover*)pParam;

	try
	{
		hr = CoInitializeEx(0, COINIT_MULTITHREADED);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"CoInitializeEx failed");

		pStudyMover->CopyWithProgress();

		// Make sure files are flushed
		VsiFlushVolume(pStudyMover->m_strDestinationPath);
	}
	VSI_CATCH(hr);

	pStudyMover->m_wndProgress.PostMessage(WM_CLOSE);

	CoUninitialize();

	return 0;
}

/// <summary>
/// This is the callback function used by the progress dialog to perform the import job
/// </summary>
DWORD CALLBACK CVsiStudyMover::ExportImportCallback(HWND hWnd, UINT uiState, LPVOID pData)
{
	CVsiStudyMover *pMover = (CVsiStudyMover*)pData;

	if (VSI_PROGRESS_STATE_INIT == uiState)
	{
		// Cache progress dialog handle
		pMover->m_wndProgress = hWnd;

		// Set progress bar range to the total size of all of the studies to
		// be copied.
		int iMegs = (int)(pMover->m_ullTotalSize / VSI_ONE_MEGABYTE);

		// For really small studies...
		if (iMegs < 1)
			iMegs = 1;
		SendMessage(hWnd, WM_VSI_PROGRESS_SETMAX, iMegs, 0);

		VSI_THREAD_START ts;
		ts.uiFlags = VSI_THREAD_START_DATA | VSI_THREAD_START_DEBUGNAME;
		ts.pszThreadDebugger = L"Export Import Callback";

		switch (pMover->m_Operation)
		{
		case VSI_SMO_IMPORTING:
			SetWindowText(hWnd, CString(MAKEINTRESOURCE(IDS_STYMVR_COPY_STUDY_FROM)));
			ts.pszThread = L"Copy Study From";
			break;
		case VSI_SMO_EXPORTING:
			SetWindowText(hWnd, CString(MAKEINTRESOURCE(IDS_STYMVR_COPY_STUDY_TO)));
			ts.pszThread = L"Copy Study To";
			break;
		default:
			_ASSERT(0);
			break;
		}

		SendMessage(hWnd, WM_VSI_PROGRESS_SETMSG, 0, (LPARAM)L"Copying data...");

		// Launch worker thread
		ts.threadFunction = StudyMoverThreadProc;
		ts.data = (LPVOID)pMover;
		VsiThreadStart(&ts);
	}
	else if (VSI_PROGRESS_STATE_CANCEL == uiState)
	{
		pMover->m_bCancelCopy = TRUE;

		SendMessage(hWnd, WM_VSI_PROGRESS_SETMSG, 0, (LPARAM)L"Cancelling...");
	}
	else if (VSI_PROGRESS_STATE_END == uiState)
	{
		// Nothing
	}

	return 0;
}