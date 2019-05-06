/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudy.h
**
**	Description:
**		Declaration of the CVsiStudy
**
*******************************************************************************/

#pragma once


#include <ATLComTime.h>
#include <VsiPropMap.h>
#include <VsiStudyModule.h>
#include "resource.h"       // main symbols


class ATL_NO_VTABLE CVsiStudy :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CVsiStudy, &CLSID_VsiStudy>,
	public IPersistStreamInit,
	public IVsiSAXContentHandlerImpl,
	public IVsiPersistStudy,
	public IVsiStudy
{
private:
	CCriticalSection m_cs;

	CComBSTR m_strFile;

	// Session info
	DWORD m_dwModified;

	VSI_STUDY_ERROR m_dwErrorCode;

	// Properties
	CVsiMapProp m_mapProp;

	// Series
	std::map<CString, CComPtr<IVsiSeries>> m_mapSeries;  // id to IVsiSeries

	// Series sorting support - must match VsiDataListWnd::CVsiSeries sorting!!!
	class CVsiSeries
	{
	public:
		IVsiSeries* m_pSeries;
		class CVsiSort
		{
		public:
			VSI_PROP_SERIES m_sortBy;
			BOOL m_bDescending;
		};
		CVsiSort *m_pSort;

		CVsiSeries() :
			m_pSeries(NULL),
			m_pSort(NULL)
		{
		}
		CVsiSeries(
			IVsiSeries *pSeries,
			CVsiSort *pSort)
		:
			m_pSeries(pSeries),
			m_pSort(pSort)
		{
		}
		bool operator<(const CVsiSeries& right) const
		{
			int iCmp = 0;

			m_pSeries->Compare(right.m_pSeries, m_pSort->m_sortBy, &iCmp);

			if (m_pSort->m_bDescending)
				iCmp *= -1;

			return (iCmp < 0);
		}
	};
	typedef std::list<CVsiSeries> CVsiListSeries;

public:
	CVsiStudy();
	~CVsiStudy();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSISTUDY)

	BEGIN_COM_MAP(CVsiStudy)
		COM_INTERFACE_ENTRY(IVsiStudy)
		COM_INTERFACE_ENTRY(IVsiPersistStudy)
		COM_INTERFACE_ENTRY(ISAXContentHandler)
		COM_INTERFACE_ENTRY(IPersistStreamInit)
		COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct()
	{
		return CoCreateFreeThreadedMarshaler(
			GetControllingUnknown(), &m_pUnkMarshaler.p);
	}

	void FinalRelease()
	{
		m_pUnkMarshaler.Release();
	}

	CComPtr<IUnknown> m_pUnkMarshaler;

	VSI_DECLARE_USE_MESSAGE_LOG()

public:
	// IVsiStudy
	STDMETHOD(Clone)(IVsiStudy **ppStudy);

	STDMETHOD(GetErrorCode)(VSI_STUDY_ERROR *pdwErrorCode);

	STDMETHOD(GetDataPath)(LPOLESTR *ppszPath);
	STDMETHOD(SetDataPath)(LPCOLESTR pszPath);

	STDMETHOD(Compare)(IVsiStudy *pStudy, VSI_PROP_STUDY dwProp, int *piCmp);

	STDMETHOD(GetProperty)(VSI_PROP_STUDY dwProp, VARIANT *pvtValue);
	STDMETHOD(SetProperty)(VSI_PROP_STUDY dwProp, const VARIANT *pvtValue);

	STDMETHOD(AddSeries)(IVsiSeries *pSeries);
	STDMETHOD(RemoveSeries)(LPCOLESTR pszId);
	STDMETHOD(RemoveAllSeries)();

	STDMETHOD(GetSeriesCount)(LONG *plCount, BOOL *pbCancel);
	STDMETHOD(GetSeriesEnumerator)(
		VSI_PROP_SERIES sortBy,
		BOOL bDescending,
		IVsiEnumSeries **ppEnum,
		BOOL *pbCancel);
	STDMETHOD(GetSeries)(LPCOLESTR pszId, IVsiSeries **ppSeries);
	STDMETHOD(GetItem)(LPCOLESTR pszNs, IUnknown **ppUnk);

	// IVsiPersistStudy
	STDMETHOD(LoadStudyData)(LPCOLESTR pszFile, DWORD dwFlags, BOOL *pbCancel);
	STDMETHOD(SaveStudyData)(LPCOLESTR pszFile, DWORD dwFlags);

	// ISAXContentHandler
	STDMETHOD(startElement)(
		const wchar_t *pwszNamespaceUri,
		int cchNamespaceUri,
		const wchar_t *pwszLocalName,
		int cchLocalName,
		const wchar_t *pwszQName,
		int cchQName,
		ISAXAttributes *pAttributes);

	// IPersistStreamInit
	STDMETHOD(GetClassID)(CLSID *pClassID);
	STDMETHOD(IsDirty)();
	STDMETHOD(InitNew)();
	STDMETHOD(Load)(LPSTREAM pStm);
	STDMETHOD(Save)(LPSTREAM pStm, BOOL fClearDirty);
	STDMETHOD(GetSizeMax)(ULARGE_INTEGER* pcbSize);

private:
	HRESULT SaveStudyDataInternal(LPCOLESTR pszFile);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiStudy), CVsiStudy)
