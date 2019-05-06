/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiSeries.h
**
**	Description:
**		Declaration of the CVsiSeries
**
*******************************************************************************/

#pragma once


#include <VsiSaxUtils.h>
#include <VsiPropMap.h>
#include <VsiStudyModule.h>
#include "resource.h"       // main symbols


class ATL_NO_VTABLE CVsiSeries :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CVsiSeries, &CLSID_VsiSeries>,
	public IPersistStreamInit,
	public IVsiSAXContentHandlerImpl,
	public IVsiPersistSeries,
	public IVsiSeries
{
private:
	CCriticalSection m_cs;

	IVsiStudy *m_pStudy;  // weak ref
	CComBSTR m_strFile;

	DWORD m_dwModified;
	VSI_SERIES_ERROR m_dwErrorCode;

	// Properties
	CVsiMapProp m_mapProp;

	// Labels
	typedef std::map<CString, CString> CVsiMapPropLabel;
	CVsiMapPropLabel m_mapPropLabel;

	// Images
	typedef std::map<CString, CComPtr<IVsiImage> > CVsiMapImage;
	CVsiMapImage m_mapImage;  // id to IVsiImage

	// Images sorting support - must match VsiDataListWnd::CVsiImage sorting!!!
	class CVsiImage
	{
	public:
		IVsiImage* m_pImage;
		class CVsiSort
		{
		public:
			VSI_PROP_IMAGE m_sortBy;
			BOOL m_bDescending;
			IVsiModeManager *m_pModeMgr;
		};
		CVsiSort *m_pSort;

		CVsiImage() :
			m_pImage(NULL),
			m_pSort(NULL)
		{
		}
		CVsiImage(
			IVsiImage *pImage,
			CVsiSort *pSort)
		:
			m_pImage(pImage),
			m_pSort(pSort)
		{
		}
		bool operator<(const CVsiImage& right) const
		{
			int iCmp = 0;

			m_pImage->Compare(right.m_pImage, m_pSort->m_sortBy, &iCmp);

			if (m_pSort->m_bDescending)
				iCmp *= -1;

			return (iCmp < 0);
		}

	private:
		int CompareDates(const CVsiImage& right) const
		{	
			int iCmp = 0;
			HRESULT hr1, hr2;

			CComVariant v1, v2;
			hr1 = m_pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &v1);
			hr2 = right.m_pImage->GetProperty(VSI_PROP_IMAGE_ACQUIRED_DATE, &v2);
			if (S_OK == hr1 && S_OK == hr2)
			{
				COleDateTime cDate1(v1), cDate2(v2);
				if (cDate1 < cDate2)
				{
					iCmp = -1;
				}
				else if (cDate2 < cDate1)
				{
					iCmp = 1;
				}
				else  // Same - checks creation date
				{
					CComVariant vCreate1, vCreate2;
					m_pImage->GetProperty(VSI_PROP_IMAGE_CREATED_DATE, &vCreate1);
					right.m_pImage->GetProperty(VSI_PROP_IMAGE_CREATED_DATE, &vCreate2);
					cDate1 = vCreate1;
					cDate2 = vCreate2;
					if (cDate1 < cDate2)
					{
						iCmp = -1;
					}
					else if (cDate2 < cDate1)
					{
						iCmp = 1;
					}
				}
			}
			else if (S_OK == hr1)
			{
				iCmp = 1;
			}
			else if (S_OK == hr2)
			{
				iCmp = -1;
			}

			return iCmp;
		}
	};
	typedef std::list<CVsiImage> CVsiListImage;

public:
	CVsiSeries();
	~CVsiSeries();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSISERIES)

	BEGIN_COM_MAP(CVsiSeries)
		COM_INTERFACE_ENTRY(IVsiSeries)
		COM_INTERFACE_ENTRY(IVsiPersistSeries)
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
	// IVsiSeries
	STDMETHOD(Clone)(IVsiSeries **ppSeries);

	STDMETHOD(GetErrorCode)(VSI_SERIES_ERROR *pdwErrorCode);

	STDMETHOD(GetDataPath)(LPOLESTR *ppszPath);
	STDMETHOD(SetDataPath)(LPCOLESTR pszPath);
	STDMETHOD(GetStudy)(IVsiStudy **ppStudy);
	STDMETHOD(SetStudy)(IVsiStudy *pStudy);

	STDMETHOD(Compare)(IVsiSeries *pSeries, VSI_PROP_SERIES dwProp, int *piCmp);

	STDMETHOD(GetProperty)(VSI_PROP_SERIES dwProp, VARIANT *pvValue);
	STDMETHOD(SetProperty)(VSI_PROP_SERIES dwProp, const VARIANT *pvValue);
	STDMETHOD(GetPropertyLabel)(
		VSI_PROP_SERIES dwProp,
		LPOLESTR *ppszLabel);

	STDMETHOD(AddImage)(IVsiImage *pImage);
	STDMETHOD(RemoveImage)(LPCOLESTR pszId);
	STDMETHOD(RemoveAllImages)();

	STDMETHOD(GetImageCount)(LONG *plCount, BOOL *pbCancel);
	STDMETHOD(GetImageEnumerator)(
		VSI_PROP_IMAGE sortBy,
		BOOL bDescending,
		IVsiEnumImage **ppEnum,
		BOOL *pbCancel);
	STDMETHOD(GetImageFromIndex)(
		int iIndex,
		VSI_PROP_IMAGE sortBy,
		BOOL bDescending,
		IVsiImage **ppImage);
	STDMETHOD(GetImage)(LPCOLESTR pszId, IVsiImage **ppImage);
	STDMETHOD(GetItem)(LPCOLESTR pszNs, IUnknown **ppUnk);

	// IVsiPersistSeries
	STDMETHOD(LoadSeriesData)(LPCOLESTR pszFile, DWORD dwFlags, BOOL *pbCancel);
	STDMETHOD(SaveSeriesData)(LPCOLESTR pszFile, DWORD dwFlags);
	STDMETHOD(LoadPropertyLabel)(LPCOLESTR pszFile);

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
	void UpdateNamespace() throw(...);
	HRESULT SavePropertyLabel(IUnknown *pUnkContent);
	HRESULT SaveSeriesDataInternal(LPCOLESTR pszFile);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiSeries), CVsiSeries)
