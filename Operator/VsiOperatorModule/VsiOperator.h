/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperator.h
**
**	Description:
**		Declaration of the CVsiOperator
**
*******************************************************************************/

#pragma once

#include "resource.h"       // main symbols
#include <VsiOperatorModule.h>

class ATL_NO_VTABLE CVsiOperator :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CVsiOperator, &CLSID_VsiOperator>,
	public IVsiPersistOperator,
	public IVsiOperator
{
private:
	CString m_strId;
	CString m_strName;
	CString m_strPassword;
	CString m_strGroup;
	VSI_OPERATOR_TYPE m_dwType;
	VSI_OPERATOR_STATE m_dwState;
	VSI_OPERATOR_DEFAULT_STUDY_ACCESS m_dwDefaultStudyAccess;
	struct CVsiSessionInfo : VSI_OPERATOR_SESSION_INFO
	{
		CVsiSessionInfo()
		{
			ZeroMemory(&stLastLogin, sizeof(SYSTEMTIME));
			ZeroMemory(&stLastLogout, sizeof(SYSTEMTIME));
			dwTotalSeconds = 0;
		}
	} m_sessionInfo;

	CCriticalSection m_cs;

public:
	CVsiOperator();
	~CVsiOperator();

	DECLARE_REGISTRY_RESOURCEID(IDR_VSIOPERATOR)

	BEGIN_COM_MAP(CVsiOperator)
		COM_INTERFACE_ENTRY(IVsiOperator)
		COM_INTERFACE_ENTRY(IVsiPersistOperator)
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
	// IVsiOperator
	STDMETHOD(GetId)(LPOLESTR *ppszId);
	STDMETHOD(SetId)(LPCOLESTR pszId);
	STDMETHOD(GetName)(LPOLESTR *ppszName);
	STDMETHOD(SetName)(LPCOLESTR pszName);
	STDMETHOD(GetPassword)(LPOLESTR *ppszPassword);
	STDMETHOD(SetPassword)(LPCOLESTR pszPassword, BOOL bHashed);
	STDMETHOD(TestPassword)(LPCOLESTR pszPassword);
	STDMETHOD(HasPassword)();
	STDMETHOD(GetGroup)(LPOLESTR *ppszGroup);
	STDMETHOD(SetGroup)(LPCOLESTR pszGroup);
	STDMETHOD(GetType)(VSI_OPERATOR_TYPE *pdwType);
	STDMETHOD(SetType)(VSI_OPERATOR_TYPE dwType);
	STDMETHOD(GetState)(VSI_OPERATOR_STATE *pdwState);
	STDMETHOD(SetState)(VSI_OPERATOR_STATE dwState);
	STDMETHOD(GetDefaultStudyAccess)(VSI_OPERATOR_DEFAULT_STUDY_ACCESS *pdwStudyAccess);
	STDMETHOD(SetDefaultStudyAccess)(VSI_OPERATOR_DEFAULT_STUDY_ACCESS dwStudyAccess);
	STDMETHOD(GetSessionInfo)(VSI_OPERATOR_SESSION_INFO *pInfo);
	STDMETHOD(SetSessionInfo)(const VSI_OPERATOR_SESSION_INFO *pInfo);

	// IVsiPersistOperator
	STDMETHOD(LoadXml)(IXMLDOMNode* pXMLNode);
	STDMETHOD(SaveXml)(IXMLDOMNode* pXMLNode);
	STDMETHOD(SaveSax)(IUnknown* pUnkISAXContentHandler);
};

OBJECT_ENTRY_AUTO(__uuidof(VsiOperator), CVsiOperator)
