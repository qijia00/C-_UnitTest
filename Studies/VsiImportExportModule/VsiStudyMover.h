/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudyMover.h
**
**	Description:
**		Declaration of the CVsiStudyMover
**
*******************************************************************************/

#pragma once

#include <VsiStudyModule.h>
#include <VsiImportExportModule.h>
#include <VsiGlobalDef.h>
#include "resource.h"       // main symbols



// CVsiStudyMover
class ATL_NO_VTABLE CVsiStudyMover :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CVsiStudyMover, &CLSID_VsiStudyMover>,
	public IVsiStudyMover
{
private:
	HWND m_hWndOwner;
	BOOL m_bAllowOverwrite;
	BOOL m_bAutoDelete;
	CString m_strDestinationPath;  // parent folder path with trailing slash
	CString m_strCustomFolderName; // user provided folder name for the copying of
								   // a single study

	CComQIPtr<IXMLDOMDocument> m_pXmlDoc;
	CComPtr<IXMLDOMElement> m_pStudy;

public:
	// Defines for CVsiStudyMover operation.
	typedef enum
	{
		VSI_SMO_NO_OPERATION,
		VSI_SMO_IMPORTING,
		VSI_SMO_EXPORTING,
	} VSI_STUDY_MOVER_OPERATION;

	VSI_STUDY_MOVER_OPERATION m_Operation;

	BOOL m_bCancelCopy;  // Set this TRUE to abort the current file copy.

	// Used by the progress dialog to estimate how much longer the process
	// is going to take.
	ULONGLONG m_ullCopiedSoFar;
	ULONGLONG m_ullTotalSize;

	// Used by progress dialog to estimate progress and to know when to update
	// the estimate string
	DWORD m_dwStartTime;
	DWORD m_dwNextReportTime;
	CWindow m_wndProgress;

public:
	CVsiStudyMover();
	~CVsiStudyMover();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSISTUDYMOVER)

	BEGIN_COM_MAP(CVsiStudyMover)
		COM_INTERFACE_ENTRY(IVsiStudyMover)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

	VSI_DECLARE_USE_MESSAGE_LOG()

public:
	// IVsiStudyMover
	STDMETHOD(Initialize)(HWND hParentWnd, const VARIANT* pvParam, ULONGLONG ullTotalSize, DWORD dwFlags);
	STDMETHOD(Export)(LPCOLESTR pszDestPath, LPCOLESTR pszCustomFolderName = NULL);
	STDMETHOD(Import)(LPCOLESTR pszDestPath);

private:
	void ConvertToExtendedLengthPath(CString &strPath); 
	int CopyWithProgress();

	BOOL CreateDestination(IXMLDOMElement *pStudyElement,
		LPWSTR pDestinationPath, int iPathLength);

	BOOL CreateRollbackName(LPCWSTR pDestinationPath, CString &strRollbackName);
	BOOL CopySeriesFolder(IXMLDOMElement *pStudyElement, LPCWSTR pszSieresFolder, LPCWSTR pSourcePath, LPCWSTR pDestinationPath);
	BOOL RollbackStudyCopy(LPCWSTR pszDestinationPath);
	BOOL RemoveRollbackDirectory(LPCWSTR pDestinationPath);

	BOOL CopySeriesFromTempToTarget(LPCWSTR pszSourcePath, LPCWSTR pszDestPath);

	static UINT __stdcall StudyMoverThreadProc(void* pParam);
	static DWORD CALLBACK ExportImportCallback(HWND hWnd, UINT uiState, LPVOID pData);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiStudyMover), CVsiStudyMover)
