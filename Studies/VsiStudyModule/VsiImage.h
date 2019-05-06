/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiImage.h
**
**	Description:
**		Declaration of the CVsiImage
**
*******************************************************************************/

#pragma once


#include <VsiPropMap.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiStudyModule.h>
#include "resource.h"       // main symbols


class ATL_NO_VTABLE CVsiImage :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CVsiImage, &CLSID_VsiImage>,
	public IPersistStreamInit,
	public IVsiPersistImage,
	public IVsiImage
{
private:
	IVsiSeries *m_pSeries;  // weak ref
	CComBSTR m_strFile;

	DWORD m_dwModified;
	VSI_IMAGE_ERROR m_dwErrorCode;

	// Properties
	CVsiMapProp m_mapProp;

public:
	CVsiImage();
	~CVsiImage();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSIIMAGE)

	BEGIN_COM_MAP(CVsiImage)
		COM_INTERFACE_ENTRY(IVsiImage)
		COM_INTERFACE_ENTRY(IVsiPersistImage)
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
	// IVsiImage
	STDMETHOD(Clone)(IVsiImage **ppImage);
	STDMETHOD(GetErrorCode)(VSI_IMAGE_ERROR *pdwErrorCode);
	STDMETHOD(GetDataPath)(LPOLESTR *ppszPath);
	STDMETHOD(SetDataPath)(LPCOLESTR pszPath);
	STDMETHOD(GetSeries)(IVsiSeries **ppSeries);
	STDMETHOD(SetSeries)(IVsiSeries *pSeries);
	STDMETHOD(GetProperty)(VSI_PROP_IMAGE dwProp, VARIANT* pvtValue);
	STDMETHOD(SetProperty)(VSI_PROP_IMAGE dwProp, const VARIANT *pvtValue);
	STDMETHOD(GetThumbnailFile)(LPOLESTR *ppszFile);
	STDMETHOD(GetStudy)(IVsiStudy **ppStudy);
	STDMETHOD(Compare)(IVsiImage *pImage, VSI_PROP_IMAGE dwProp, int *piCmp);

	// IVsiPersistImage
	STDMETHOD(LoadImageData)(
		LPCOLESTR pszFile,
		IUnknown *pUnkMode,
		IUnknown *pUnkPdm,
		DWORD dwFlags);
	STDMETHOD(SaveImageData)(
		LPCOLESTR pszFile,
		IUnknown *pUnkMode,
		IUnknown *pUnkPdm,
		DWORD dwFlags);
	STDMETHOD(ResaveProperties)(LPCOLESTR pszFile);

	// IPersistStreamInit
	STDMETHOD(GetClassID)(CLSID *pClassID);
	STDMETHOD(IsDirty)();
	STDMETHOD(InitNew)();
	STDMETHOD(Load)(LPSTREAM pStm);
	STDMETHOD(Save)(LPSTREAM pStm, BOOL fClearDirty);
	STDMETHOD(GetSizeMax)(ULARGE_INTEGER* pcbSize);

private:
	HRESULT SaxParseDom(IXMLDOMDocument *pXmlDoc, LPCWSTR pszFile);
	void UpdateNamespace() throw(...);
	HRESULT DoConversion();
	int CompareDates(IVsiImage* pImage);
	HRESULT SaveImageDataInternal(
		LPCOLESTR pszFile,
		IUnknown *pUnkMode,
		IUnknown *pUnkPdm,
		DWORD dwFlags);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiImage), CVsiImage)
