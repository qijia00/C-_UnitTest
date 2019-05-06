/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiCSVWriter.h
**
**	Description:
**		Declaration of the CVsiCSVWriter
**
*******************************************************************************/

#pragma once

#include <VsiPdmModule.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiMsmntCommon.h>
#include <VsiResProduct.h>

class CVsiCSVWriter
{
public:
	typedef enum
	{
		VSI_CSV_WRITER_FILTER_NONE			= 0x00000000,
		VSI_CSV_WRITER_FILTER_MEASUREMENTS	= 0x00000001,
		VSI_CSV_WRITER_FILTER_CALCULATIONS	= 0x00000002,
	} VSI_CSV_WRITER_FILTER;

private:
	VSI_CSV_WRITER_FILTER m_filter;

	typedef struct tagVSI_STUDY_ID
	{
		WCHAR	m_szStudyID[80];
		long	m_lIndex;
	} VSI_STUDY_ID;

	typedef struct tagVSI_IMAGE_INFO
	{
		CString	m_strImageLabel;
		CString	m_strFrequency;
	} VSI_IMAGE_INFO;

	#define STUDY_ID_LIST std::vector<VSI_STUDY_ID>
	#define STUDY_ID_LIST_ITER STUDY_ID_LIST::iterator
	#define IMAGE_ID_NAME_MAP std::map<CString, VSI_IMAGE_INFO>

	STUDY_ID_LIST m_StudyIDList;

	CComPtr<IXMLDOMDocument> m_pDoc;
	CComPtr<IVsiApp> m_pApp;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiModeManager> m_pModeMgr;
	WCHAR m_szSptr[2];

	BOOL m_bFemale;
	BOOL m_bPregnant;

	CVsiVecMsmntPVInfo m_vPVInfo;
	IMAGE_ID_NAME_MAP m_mapImages;

public:
	CVsiCSVWriter() :
		m_filter(VSI_CSV_WRITER_FILTER_NONE),
		m_bFemale(FALSE),
		m_bPregnant(FALSE)
	{
	}
	~CVsiCSVWriter()
	{
		_ASSERT(NULL == m_pDoc);
		_ASSERT(NULL == m_pApp);
		_ASSERT(NULL == m_pPdm);
		_ASSERT(NULL == m_pModeMgr);
	}

	HRESULT ConvertFile(
		IXMLDOMDocument *pXmlDoc,
		IVsiApp *pApp,
		VSI_CSV_WRITER_FILTER filter,
		LPCWSTR pszDestinationFile);

private:
	HRESULT GetFieldSeparator(LPWSTR szSptr, int iCount);
	HRESULT GetFieldSeparator();

	int GetStudyIndex(WCHAR *pszStudyID);
	int GetNumberOfStudies();
	long BuildStudiesList();
	
	HRESULT ConvertHeaderSection(HANDLE hFile);
	HRESULT ConvertStudy(HANDLE hFile, IXMLDOMElement *pStudyElement, BOOL bVevo2100);
	HRESULT ConvertSeries(HANDLE hFile, IXMLDOMElement *pSeriesElement, BOOL bVevo2100);

	HRESULT ExportPVLoopData(CString& strOut);
	DWORD GetLongestLoop();

	HRESULT ConvertProtocolMeasurements(CString& strOut, IXMLDOMElement *pCollectionElement, BOOL bGeneric);
	HRESULT ConvertProtocolCalculations(CString& strOut, IXMLDOMElement *pProtocolElement);
	HRESULT ConvertProtocolSection(CString& strOut, IXMLDOMElement *pPkgElement);
	HRESULT ConvertPackageInfo(CString& strOut, IXMLDOMElement *pPkgElement);	
	HRESULT ExportMsmntData(CString& strOut, IXMLDOMElement *pMeasurementElement);
	HRESULT ExportHistogramData(CString& strOut, IXMLDOMElement *pDataSegmentElement);
	HRESULT ExportContrastData(CString& strOut, IXMLDOMElement *pDataSegmentElement);
	HRESULT ExportCardiacData(CString& strOut, IXMLDOMElement *pDataSegmentElement);
	HRESULT ExportPVData(CString& strOut, IXMLDOMElement *pDataSegmentElement);
	HRESULT ExportBLVData(CString& strOut, IXMLDOMElement *pDataSegmentElement);
	HRESULT ExportMLVData(CString& strOut, IXMLDOMElement *pDataSegmentElement);
	HRESULT ExportPAData(CString& strOut, IXMLDOMElement *pDataSegmentElement);
	HRESULT FillImageMap(IXMLDOMElement *pSeriesElement);
	HRESULT WriteCSVFile(HANDLE hFile, CString& strOut);

	void ReportNoMsmnts(CString& strOut);

	void FormatVariantDouble(CComVariant &vNum, int iPrecision = 0) throw(...);
};

DEFINE_ENUM_FLAG_OPERATORS(CVsiCSVWriter::VSI_CSV_WRITER_FILTER);
