/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiStudyManager.h
**
**	Description:
**		Declaration of the CVsiStudyManager
**
*******************************************************************************/

#pragma once

#include <VsiPropMap.h>
#include <IVsiEventSourceImpl.h>
#include <VsiAppModule.h>
#include <VsiPdmModule.h>
#include <VsiStudyModule.h>
#include "resource.h"       // main symbols


class ATL_NO_VTABLE CVsiStudyManager :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CVsiStudyManager, &CLSID_VsiStudyManager>,
	public IVsiEventSourceImpl<CVsiStudyManager>,
	public IVsiStudyManager
{
private:
	CComQIPtr<IVsiApp> m_pApp;
	CComPtr<IVsiPdm> m_pPdm;

	CCriticalSection m_cs;

	std::map<CString, CComPtr<IVsiStudy>> m_mapStudy;  // id to IVsiStudy

	CString m_strPath;

	class CVsiFilter
	{
	public:
		CString m_strOperator;
		CString m_strGroup;
	    BOOL m_bPublic;

		CVsiFilter(LPCWSTR pszOperator, LPCWSTR pszGroup, BOOL bPublic) :
			m_strOperator(pszOperator),
			m_strGroup(pszGroup),
			m_bPublic(bPublic)
		{
		}
	};

	std::unique_ptr<CVsiFilter> m_pFilter;

public:
	CVsiStudyManager();
	~CVsiStudyManager();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSISTUDYMANAGER)

	BEGIN_COM_MAP(CVsiStudyManager)
		COM_INTERFACE_ENTRY(IVsiStudyManager)
		COM_INTERFACE_ENTRY(IVsiEventSource)
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
	// IVsiStudyManager
	STDMETHOD(Initialize)(IVsiApp *pApp);
	STDMETHOD(Uninitialize)();

	STDMETHOD(LoadData)(
		LPCOLESTR pszPath,
		DWORD dwTypes,
		const VSI_LOAD_STUDY_FILTER *pFilter);
	STDMETHOD(LoadStudy)(LPCOLESTR pszPath, BOOL bFailAllow, IVsiStudy** ppStudy);

	STDMETHOD(GetStudyCount)(LONG *plCount);
	STDMETHOD(GetStudyEnumerator)(BOOL bExcludeCorruptedAndNotCompatible, IVsiEnumStudy **ppEnum);
	STDMETHOD(GetStudy)(LPCOLESTR pszNs, IVsiStudy **ppStudy);
	STDMETHOD(GetItem)(LPCOLESTR pszNs, IUnknown **ppUnk);
	STDMETHOD(GetIsItemPresent)(LPCOLESTR pszNs);
	STDMETHOD(GetIsStudyPresent)(
		VSI_PROP_STUDY dwIdType,
		LPCOLESTR pszCreatedId,
		LPCOLESTR pszOperatorName,
		IVsiStudy **ppStudy);

	STDMETHOD(AddStudy)(IVsiStudy *pStudy);
	STDMETHOD(RemoveStudy)(LPCOLESTR pszNs);

	STDMETHOD(Update)(IXMLDOMDocument *pXmlDoc);
	STDMETHOD(GetStudyCopyNumber)(LPCOLESTR pszCreatedId, LPCOLESTR pszOwner, int* piNum);
	STDMETHOD(MoveSeries)(IVsiSeries *pSrcSeries, IVsiStudy *pDesStudy, LPOLESTR *ppszResult);

private:
	bool IsStudyVisible(
		IVsiStudy *pStudy,
		IVsiOperatorManager *pOperatorMgr,
		std::set<CString> &setOperatorInGroup,
		std::set<CString> &setOperatorNotInGroup);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiStudyManager), CVsiStudyManager)
