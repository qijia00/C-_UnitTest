/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiReportExportStatus.h
**
**	Description:
**		Declaration of the CVsiReportExportStatus
**
********************************************************************************/

#include <VsiRes.h>


class CVsiReportExportStatus : public CDialogImpl<CVsiReportExportStatus>
{
public:
	enum { IDD = IDD_EXPORT_REPORT };

	typedef enum
	{
		REPORT_TYPE_STUDY_EXPORT = 1,
		REPORT_TYPE_STUDY_IMPORT,
		REPORT_TYPE_IMAGE_EXPORT,
		REPORT_TYPE_DICOM_EXPORT,
	} REPORT_TYPE;	

private:
	CComQIPtr<IXMLDOMDocument> m_pXmlDoc;
	REPORT_TYPE m_type;

public:
	CVsiReportExportStatus(IUnknown *pXmlDoc, REPORT_TYPE type);
	~CVsiReportExportStatus();

protected:
	// Message map and handlers
	BEGIN_MSG_MAP(CVsiReportExportStatus)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		COMMAND_ID_HANDLER(IDOK, OnCmdOK)

		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnClose(UINT, WPARAM, LPARAM, BOOL&);

	LRESULT OnCmdOK(WORD, WORD, HWND, BOOL&);

private:
	void AddStudyItem(IXMLDOMElement *pStudyElement);
	int ReportStudySection(DWORD dwCode, LPCWSTR pszTitle);
	void AddImageItem(IXMLDOMElement *pImageElement);
	int ReportImageSection(DWORD dwCode, LPCWSTR pszTitle);
};

