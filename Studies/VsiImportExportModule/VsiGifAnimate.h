/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiGifAnimate.h
**
**	Description:
**		Code to enable animation of gif files
**
********************************************************************************/

#pragma once

#include <VsiImportExportModule.h>
#include "resource.h"
#include "VsiGifEncode.h"


// CVsiGifAnimate

class ATL_NO_VTABLE CVsiGifAnimate :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CVsiGifAnimate, &CLSID_VsiGifAnimate>,
	public CVsiThread<CVsiGifAnimate>,
	public IVsiGifAnimateAvi,
	public IVsiGifAnimateDirect
{
private:
	CWindow m_hwndProgress;
	CString m_strInputFileName;
	CString m_strOutputFileName;	
	BOOL m_bError;

	// For IVsiGifAnimateDirect
	CVsiGifEncode m_gifEncode;
	CComPtr<IVsiColorQuantize> m_pColorReduce[4];

	BYTE *m_ppbtTransparentMask[4+1];
	BYTE *m_ppbtImageDataSrc[4+1];
	BYTE *m_ppbtImageDataDst[4+1];
	DWORD m_dwNumSourceFrames;
	BOOL m_bFirstDataSet;
	DWORD m_dwNumThreads;
	
	DWORD m_dwXRes;
	DWORD m_dwYRes;

public:
	CVsiGifAnimate();
	~CVsiGifAnimate();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSIGIFANIMATE)

	BEGIN_COM_MAP(CVsiGifAnimate)
		COM_INTERFACE_ENTRY(IVsiGifAnimateAvi)
		COM_INTERFACE_ENTRY(IVsiGifAnimateDirect)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

	// IVsiGifAnimateDirect
	STDMETHOD(Start)(
		DWORD dwXRes,
		DWORD dwYRes,
		DWORD dwRate,
		DWORD dwScale,
		LPCOLESTR pszOutputFileName);
	STDMETHOD(AddFrame)(
		BYTE *pbtData,
		DWORD dwSize);
	STDMETHOD(Stop)();

	// IVsiGifAnimateAvi
	STDMETHOD(ConvertAviToGif)(
		HWND hwndProgress, 
		LPCOLESTR pszInputFileName, 
		LPCOLESTR pszOutputFileName); 
	STDMETHOD(ConvertAviToGifDirect)(
		HWND hwndProgress, 
		LPCOLESTR pszInputFileName, 
		LPCOLESTR pszOutputFileName); 

	// CVsiThread
	DWORD ThreadProc();

	BOOL IsBusy()
	{
		return (S_OK == IsThreadRunning());
	}

	VSI_DECLARE_USE_MESSAGE_LOG()

private:
	void CreateCompareMask(
		const BYTE *pbtSource0, 
		const BYTE *pbtSource1, 
		BYTE *pbtMask, 
		DWORD dwLength);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiGifAnimate), CVsiGifAnimate)
