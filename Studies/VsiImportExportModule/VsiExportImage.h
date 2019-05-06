/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiExportImage.h
**
**	Description:
**		Declaration of the CVsiExportImage
**
*******************************************************************************/

#pragma once

#include <shldisp.h>
#include <VsiRes.h>
#include <VsiAppModule.h>
#include <VsiMessageFilter.h>
#include <VsiLogoViewBase.h>
#include <VsiAppViewModule.h>
#include <VsiStudyModule.h>
#include <VsiModeModule.h>
#include <VsiImportExportModule.h>
#include <vfw.h>
#include <VsiPathNameEditWnd.h>
#include <VsiImportExportUIHelper.h>
#include <VsiParamMeasurement.h>
#include "VsiImpExpUtility.h"
#include "VsiExportTable.h"
#include "VsiExportReport.h"
#include "resource.h"


#define WM_VSI_REFRESH_TREE					(WM_USER + 100)
#define WM_VSI_SET_SELECTION				(WM_USER + 101)
#define WM_VSI_EXPORT_IMAGE					(WM_USER + 102)
#define WM_VSI_EXPORT_IMAGE_COMPLETED		(WM_USER + 103)


class CVsiExportImageCineDlg;
class CVsiExportImageFrameDlg;
class CVsiExportImageDicomDlg;
class CVsiExportImageReportDlg;
class CVsiExportImageGraphDlg;
class CVsiExportImageTableDlg;
class CVsiExportImagePhysioDlg;


class ATL_NO_VTABLE CVsiExportImage :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVsiExportImage, &CLSID_VsiExportImage>,
	public CDialogImpl<CVsiExportImage>,
	public CVsiLogoViewBase<CVsiExportImage>,
	public CVsiImportExportUIHelper<CVsiExportImage>,
	public IVsiView,
	public IVsiPreTranslateMsgFilter,
	public IVsiExportImage
{
	friend CVsiExportImageCineDlg;
	friend CVsiExportImageFrameDlg;
	friend CVsiExportImageDicomDlg;
	friend CVsiExportImageReportDlg;
	friend CVsiExportImageGraphDlg;
	friend CVsiExportImageTableDlg;
	friend CVsiExportImagePhysioDlg;
private:
	CComPtr<IVsiApp> m_pApp;
	CComPtr<IVsiCommandManager> m_pCmdMgr;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiModeManager> m_pModeMgr;
	CComPtr<IVsiCursorManager> m_pCursorMgr;

	CComPtr<IVsiAutoCompleteSource> m_pAutoCompleteSource;
	CComPtr<IAutoComplete> m_pAutoComplete;

	DWORD m_dwFlags;

	CString m_strExportFolder;

	// VSI_EXPORT_IMAGE_FLAG_CINE
	// VSI_EXPORT_IMAGE_FLAG_IMAGE
	// VSI_EXPORT_IMAGE_FLAG_DICOM
	// VSI_EXPORT_IMAGE_FLAG_RAW
	// VSI_EXPORT_IMAGE_FLAG_REPORT
	// VSI_EXPORT_IMAGE_FLAG_TABLE
	DWORD m_dwExportTypeMain;
	
	BOOL m_bHideMsmnts;

	std::unique_ptr<CVsiExportImageCineDlg> m_pSubPanelCine;
	std::unique_ptr<CVsiExportImageFrameDlg> m_pSubPanelFrame;
	std::unique_ptr<CVsiExportImageDicomDlg> m_pSubPanelDicom;
	std::unique_ptr<CVsiExportImageReportDlg> m_pSubPanelReport;
	std::unique_ptr<CVsiExportImageGraphDlg> m_pSubPanelGraph;
	std::unique_ptr<CVsiExportImageTableDlg> m_pSubPanelTable;
	std::unique_ptr<CVsiExportImagePhysioDlg> m_pSubPanelPhysio;

	CString	m_csPath;	// Highlighted path

	class CVsiImage
	{
	public:
		CComQIPtr<IVsiImage> m_pImage;
		CComQIPtr<IVsiMode> m_pMode;

		CVsiImage()
		{
		}

		CVsiImage(IUnknown *pImage, IUnknown *pMode)
		{
			m_pImage = pImage;
			m_pMode = pMode;
		}

		// Copy constructor
		CVsiImage::CVsiImage(const CVsiImage& right)
		{
			*this = right;
		}
		virtual CVsiImage& operator=(const CVsiImage& right)
		{
			if (this == &right)
				return *this;

			m_pImage = right.m_pImage;
			m_pMode = right.m_pMode;

			return *this;
		}
	};

	typedef std::list<CVsiImage> CVsiListImage;
	typedef CVsiListImage::iterator CVsiListImageIter;

	CVsiListImage m_listImage;

	CVsiExportTable m_ExportTable;

	CVsiExportReport m_ExportReport;
	CComQIPtr<IXMLDOMDocument> m_pXmlReportDoc;
	CComQIPtr<IVsiAnalysisExport> m_pAnalysisExport;

	CVsiPathNameEditWnd m_wndName;

	int m_iOffset;

	int m_i3Ds;	

private:
	// Export progress
	class CVsiExportProgress
	{
	public:
		CWindow m_wndProgress;
		BOOL m_bCancel;
		CComPtr<IXMLDOMDocument> m_pXmlDoc;
		CVsiListImageIter m_iterExport;
		int m_iImage;
		int m_iUniqueDigit;
		BOOL m_bError;
		BOOL m_bForceReport;
		
		CVsiExportProgress() :
			m_bCancel(FALSE),
			m_iImage(0),
			m_iUniqueDigit(1),
			m_bError(FALSE),
			m_bForceReport(FALSE)
		{
		}
	};
	
	CVsiExportProgress *m_pExportProgress;

public:
	enum { IDD = IDD_EXPORT_IMAGE_VIEW };

	CVsiExportImage();
	~CVsiExportImage();

	LPCTSTR GetPath() { return m_csPath;}

	DECLARE_REGISTRY_RESOURCEID(IDR_VSIEXPORTIMAGE)

	BEGIN_COM_MAP(CVsiExportImage)
		COM_INTERFACE_ENTRY(IVsiView)
		COM_INTERFACE_ENTRY(IVsiExportImage)
		COM_INTERFACE_ENTRY(IVsiPreTranslateMsgFilter)
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
	// IVsiView
	STDMETHOD(Activate)(IVsiApp *pApp, HWND hwndParent);
	STDMETHOD(Deactivate)();
	STDMETHOD(GetWindow)(HWND *phWnd);
	STDMETHOD(GetAccelerator)(HACCEL *phAccel);
	STDMETHOD(GetIsBusy)(
		DWORD dwStateCurrent,
		DWORD dwState,
		BOOL bTryRelax,
		BOOL *pbBusy);

	// IVsiPreTranslateMsgFilter
	STDMETHOD(PreTranslateMessage)(MSG *pMsg, BOOL *pbHandled);

	// IVsiExportImage
	STDMETHOD(Initialize)(DWORD dwFlags);
	STDMETHOD(AddImage)(IUnknown *pUnkImage, IUnknown *pUnkMode);
	STDMETHOD(SetTable)(HWND hwndTable);
	STDMETHOD(SetReport)(IUnknown *pUnkXmlDoc);
	STDMETHOD(ExportImage)(IUnknown *pApp, IUnknown *pUnkImage, IUnknown *pUnkMode, LPCOLESTR pszPath);
	STDMETHOD(ExportReport)(IUnknown *pApp, IUnknown *pUnkXmlDoc, LPCOLESTR pszPath);
	STDMETHOD(SetAnalysisExport)(IUnknown *pAnalysisExport);

	// CVsiImportExportUIHelper
	BOOL GetAvailablePath(LPWSTR pszBuffer, int iBufferLength);
	int CheckForOverwrite(LPCWSTR pszPath, LPCWSTR pszFile);
	int GetNumberOfSelectedItems();
	HWND GetSaveAsControl()
	{
		return m_wndName;
	}

protected:
	BEGIN_MSG_MAP(CVsiExportImage)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
		MESSAGE_HANDLER(WM_VSI_REFRESH_TREE, OnRefreshTree)
		MESSAGE_HANDLER(WM_VSI_SET_SELECTION, OnSetSelection)
		MESSAGE_HANDLER(WM_VSI_EXPORT_IMAGE, OnExportImage)
		MESSAGE_HANDLER(WM_VSI_EXPORT_IMAGE_COMPLETED, OnExportImageCompleted)

		COMMAND_HANDLER(IDOK, BN_CLICKED, OnBnClickedOK)
		COMMAND_HANDLER(IDCANCEL, BN_CLICKED, OnBnClickedCancel)

		COMMAND_HANDLER(IDC_IE_CINE_LOOP, BN_CLICKED, OnBnClickedTypeMain)
		COMMAND_HANDLER(IDC_IE_IMAGE, BN_CLICKED, OnBnClickedTypeMain)
		COMMAND_HANDLER(IDC_IE_DICOM, BN_CLICKED, OnBnClickedTypeMain)
		COMMAND_HANDLER(IDC_IE_REPORT, BN_CLICKED, OnBnClickedTypeMain)
		COMMAND_HANDLER(IDC_IE_TABLE, BN_CLICKED, OnBnClickedTypeMain)
		COMMAND_HANDLER(IDC_IE_GRAPH, BN_CLICKED, OnBnClickedTypeMain)
		COMMAND_HANDLER(IDC_IE_PHYSIO, BN_CLICKED, OnBnClickedTypeMain)

		COMMAND_HANDLER(IDC_IE_SAVEAS, EN_CHANGE, OnFileNameChanged)

		NOTIFY_HANDLER(IDC_IE_FOLDER_TREE, TVN_SELCHANGED, OnTreeSelChanged);

		CHAIN_MSG_MAP(CVsiImportExportUIHelper<CVsiExportImage>)
		CHAIN_MSG_MAP(CVsiLogoViewBase<CVsiExportImage>)

		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnWindowPosChanged(UINT, WPARAM, LPARAM lp, BOOL &bHandled);
	LRESULT OnRefreshTree(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnSetSelection(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnExportImage(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnExportImageCompleted(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnBnClickedOK(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedCancel(WORD, WORD, HWND, BOOL&);
	LRESULT OnBnClickedTypeMain(WORD, WORD, HWND, BOOL&);
	LRESULT OnFileNameChanged(WORD, WORD, HWND, BOOL&);
	LRESULT OnTreeSelChanged(int, LPNMHDR, BOOL&);

private:
	static DWORD CALLBACK ExportImageCallback(HWND hWnd, UINT uiState, LPVOID pData);

	CWindow GetPanel(DWORD dwExportType);
	CWindow GetTypeControl();
	int GetTypeCount(int *piFrames = NULL, int *piLoops = NULL, int *pi3Ds = NULL) throw(...);
	void UpdateTypeMain(BOOL bForceUpdate = FALSE);
	void UpdateRequiredSize();
	void UpdateUI();
	int SetTextBoxLimit();
	LPCWSTR GetFileExtension();

	// Testing functions
	BOOL IsAllLoopsOfModes(ULONG mode) throw(...);
	BOOL IsAllLoopsOf3dModes() throw(...);
	BOOL IsAnyOf3dVolumeModes() throw(...);
	BOOL IsAnyOfModes(ULONG mode) throw(...);
	BOOL IsAllOfRf(VSI_RF_EXPORT_TYPE rfType) throw(...);
	BOOL IsAnyOfRf(VSI_RF_EXPORT_TYPE rfType) throw(...);
	BOOL IsAllOfEKV() throw(...);
	BOOL IsThisOfRf(IVsiImage *pImage, IVsiMode *pMode, VSI_RF_EXPORT_TYPE rfType, BOOL *pbEkv) throw(...);	

	double EstimateSize();
	double EstimateRFSize(IVsiImage *pImage, IVsiMode *pMode, BOOL bFrame, VSI_RF_EXPORT_TYPE rfType);
	
	HRESULT CreateMsmntReport(IVsiApp *pApp, IXMLDOMDocument **ppXmlMsmntDoc, CVsiListImage &imageList);
	HRESULT CopyCurves(
		IXMLDOMDocument *pXmlAnalysisReportDoc,
		IVsiImage *pImage,
		IVsiAnalysisSet *pAnalysisSet);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiExportImage), CVsiExportImage)
