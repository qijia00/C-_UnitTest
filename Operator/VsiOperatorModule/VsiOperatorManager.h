/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorManager.h
**
**	Description:
**		Declaration of the CVsiOperatorManager
**
*******************************************************************************/

#pragma once

#include <VsiStl.h>
#include <VsiAppModule.h>
#include <VsiOperatorModule.h>
#include "resource.h"


class ATL_NO_VTABLE CVsiOperatorManager :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CVsiOperatorManager, &CLSID_VsiOperatorManager>,
	public IVsiOperatorManager,
	public IVsiPersistOperatorManager
{
private:
	CComQIPtr<IVsiServiceProvider> m_pServiceProvider;
	CComPtr<IVsiPdm> m_pPdm;

	CCriticalSection m_cs;
	CString m_strDefaultDataFile;

	BOOL m_bServiceMode;

	std::map<CString, CComPtr<IVsiOperator>> m_mapIdToOperator;  // id to IVsiOperator
	std::map<CString, CComPtr<IVsiOperator>, lessNoCaseLPCWSTR> m_mapNameToOperator;  // name to IVsiOperator
	
	CString m_strCurrentOperator;

	typedef enum
	{
		VSI_SERVICE_STATUS_ALLOWED = 1,
		VSI_SERVICE_STATUS_BLOCKED
	} VSI_SERVICE_STATUS;
	std::map<CString, VSI_SERVICE_STATUS> m_mapServiceStatus;  // Service and status

	BOOL m_bModified;
	BOOL m_bAdminAuthenticated;

	std::map<CString, SYSTEMTIME> m_mapAuthenticated;  // operators currently logged on

public:
	CVsiOperatorManager();
	~CVsiOperatorManager();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSIOPERATORMANAGER)

	BEGIN_COM_MAP(CVsiOperatorManager)
		COM_INTERFACE_ENTRY(IVsiOperatorManager)
		COM_INTERFACE_ENTRY(IVsiPersistOperatorManager)
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
	// IVsiOperatorManager
	STDMETHOD(Initialize)(IVsiApp *pApp, LPCOLESTR pszDefaultDataFile);
	STDMETHOD(Uninitialize)();
	STDMETHOD(GetOperatorDataPath)(
		VSI_OPERATOR_PROP prop,
		LPCOLESTR pszProp,
		LPOLESTR *ppszPath);
	STDMETHOD(GetOperators)(VSI_OPERATOR_FILTER_TYPE types, IVsiEnumOperator **ppEnum);
	STDMETHOD(GetOperator)(
		VSI_OPERATOR_PROP prop,
		LPCOLESTR pszProp,
		IVsiOperator **ppOperator);
	STDMETHOD(GetCurrentOperator)(IVsiOperator **ppOperator);
	STDMETHOD(SetCurrentOperator)(LPCOLESTR pszId);
	STDMETHOD(GetCurrentOperatorDataPath)(LPOLESTR *ppszPath);
	STDMETHOD(GetRelationshipWithCurrentOperator)(
		VSI_OPERATOR_PROP prop,
		LPCOLESTR pszProp,
		VSI_OPERATOR_RELATIONSHIP *pRelationship);
	STDMETHOD(AddOperator)(IVsiOperator *pOperator, LPCOLESTR pszIdCloneSettingsFrom);
	STDMETHOD(RemoveOperator)(LPCOLESTR pszId);
	STDMETHOD(OperatorModified)();
	STDMETHOD(GetOperatorCount)(VSI_OPERATOR_FILTER_TYPE ulFlags, DWORD *pdwCount);
	STDMETHOD(SetIsAuthenticated)(LPCOLESTR pszId, BOOL bAuthenticated);
	STDMETHOD(GetIsAuthenticated)(LPCOLESTR pszId, BOOL *pbAuthenticated);
	STDMETHOD(GetIsAdminAuthenticated)(BOOL *pbAuthenticated);
	STDMETHOD(ClearAuthenticatedList)();

	// IVsiPersistOperatorManager
	STDMETHOD(IsDirty)();
	STDMETHOD(LoadOperator)(LPCOLESTR pszDataFile);
	STDMETHOD(SaveOperator)(LPCOLESTR pszDataFile);

private:
	bool IsOperatorType(IVsiOperator *pOperator, VSI_OPERATOR_FILTER_TYPE types);
	HRESULT LoadOperators(IXMLDOMElement* pXmlElemRoot);
	HRESULT LoadOperatorSettings(IXMLDOMNodeList* pXmlNodeList, bool bPasswordEncrypted);
	HRESULT SaveOperators(IXMLDOMElement* pXmlElem);
	void UpdatePdm(IVsiOperator *pOperator) throw(...);
	void GetHashedDate(CString &strHashed) throw(...);
	HRESULT ValidateOperatorSetup(LPCWSTR pszName);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiOperatorManager), CVsiOperatorManager)
