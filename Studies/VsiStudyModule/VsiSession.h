/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiSession.h
**
**	Description:
**		Declaration of the CVsiSession
**
*******************************************************************************/

#pragma once


#include <VsiAppModule.h>
#include <VsiPdmModule.h>
#include <VsiModeModule.h>
#include <VsiStudyModule.h>
#include "resource.h"       // main symbols


class ATL_NO_VTABLE CVsiSession :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CVsiSession, &CLSID_VsiSession>,
	public IVsiSessionJournal,
	public IVsiSession
{
private:
	CCriticalSection m_cs;
	CCriticalSection m_csStudy;
	CCriticalSection m_csSeries;
	CCriticalSection m_csImage;
	CCriticalSection m_csMode;
	CCriticalSection m_csModeView;
	VSI_SESSION_MODE m_mode;

	CComQIPtr<IVsiApp> m_pApp;
	CComPtr<IVsiPdm> m_pPdm;
	CComPtr<IVsiOperatorManager> m_pOperatorMgr;

	CComPtr<IVsiStudy> m_pStudy;		// Primary
	CComPtr<IVsiSeries> m_pSeries;		// Primary
	CComPtr<IVsiStudy> m_ppStudy[2];
	CComPtr<IVsiSeries> m_ppSeries[2];
	CComPtr<IVsiImage> m_ppImage[2];
	CComPtr<IVsiMode> m_ppMode[2];
	CComPtr<IVsiModeView> m_ppModeView[2];

	int m_iSlotActive;	// Active slot index

	typedef std::set<CString> CVsiSetItem;
	typedef CVsiSetItem::iterator CVsiSetItemIter;

	CVsiSetItem m_setReviewed;

public:
	CVsiSession();
	~CVsiSession();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSISESSION)

	BEGIN_COM_MAP(CVsiSession)
		COM_INTERFACE_ENTRY(IVsiSession)
		COM_INTERFACE_ENTRY(IVsiSessionJournal)
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
	// IVsiSession
	STDMETHOD(Initialize)(IVsiApp *pApp);
	STDMETHOD(Uninitialize)();

	STDMETHOD(GetSessionMode)(VSI_SESSION_MODE *pMode);
	STDMETHOD(SetSessionMode)(VSI_SESSION_MODE mode);

	STDMETHOD(GetPrimaryStudy)(IVsiStudy **ppStudy);
	STDMETHOD(SetPrimaryStudy)(IVsiStudy *pStudy, BOOL bJournal);
	STDMETHOD(GetPrimarySeries)(IVsiSeries **ppSeries);
	STDMETHOD(SetPrimarySeries)(IVsiSeries *pSeries, BOOL bJournal);
	STDMETHOD(GetActiveSlot)(int *piSlot);
	STDMETHOD(SetActiveSlot)(int iSlot);

	STDMETHOD(GetStudy)(int iSlot, IVsiStudy **ppStudy);
	STDMETHOD(SetStudy)(int iSlot, IVsiStudy *pStudy, BOOL bJournal);
	STDMETHOD(GetSeries)(int iSlot, IVsiSeries **ppSeries);
	STDMETHOD(SetSeries)(int iSlot, IVsiSeries *pSeries, BOOL bJournal);
	STDMETHOD(GetImage)(int iSlot, IVsiImage **ppImage);
	STDMETHOD(SetImage)(int iSlot, IVsiImage *pImage, BOOL bSetReviewed, BOOL bJournal);
	STDMETHOD(GetMode)(int iSlot, IUnknown **ppUnkMode);
	STDMETHOD(SetMode)(int iSlot, IUnknown *pUnkMode);
	STDMETHOD(GetModeView)(int iSlot, IUnknown **ppUnkModeView);
	STDMETHOD(SetModeView)(int iSlot, IUnknown *pUnkModeView);
	STDMETHOD(GetIsStudyLoaded)(LPCOLESTR pszNs, int *piSlot);
	STDMETHOD(GetIsSeriesLoaded)(LPCOLESTR pszNs, int *piSlot);
	STDMETHOD(GetIsImageLoaded)(LPCOLESTR pszNs, int *piSlot);
	STDMETHOD(GetIsItemLoaded)(LPCOLESTR pszNs,	int *piSlot);
	STDMETHOD(GetIsItemReviewed)(LPCOLESTR pszNs);

	STDMETHOD(ReleaseSlot)(int iSlot);

	STDMETHOD(CheckActiveSeries)(BOOL bReset, LPCOLESTR pszOperatorId, LPOLESTR *ppszNs);

	// IVsiSessionJournal
	STDMETHOD(Record)(
		DWORD dwAction,
		const VARIANT *pvParam1,
		const VARIANT *pvParam2);

private:
	BOOL IsStudyIdentical(IVsiStudy *pStudy1, IVsiStudy *pStudy2) throw(...);
	BOOL IsSeriesIdentical(IVsiSeries *pSeries1, IVsiSeries *pSeries2) throw(...);
	BOOL IsImageIdentical(IVsiImage *pImage1, IVsiImage *pImage2) throw(...);
	void DeleteActiveSeriesFile();
};

OBJECT_ENTRY_AUTO(__uuidof(VsiSession), CVsiSession)
