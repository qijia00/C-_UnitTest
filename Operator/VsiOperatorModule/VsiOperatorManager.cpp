/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiOperatorManager.cpp
**
**	Description:
**		Implementation of CVsiOperatorManager
**
*******************************************************************************/

#include "stdafx.h"
#include <VsiSaxUtils.h>
#include <ATLComTime.h>
#include <VsiResProduct.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>
#include <VsiServiceProvider.h>
#include <VsiPdmModule.h>
#include <VsiParameterNamespace.h>
#include <VsiParameterSystem.h>
#include <VsiParameterUtils.h>
#include <VsiStudyModule.h>
#include <VsiMeasurementModule.h>
#include <VsiServiceKey.h>
#include "VsiOperatorXml.h"
#include "VsiOperatorManager.h"


/// <summary>
/// Operator Manager constructor
/// </summary>
CVsiOperatorManager::CVsiOperatorManager() :
	m_bModified(FALSE),
	m_bAdminAuthenticated(FALSE)
{
	VSI_LOG_MSG(VSI_LOG_INFO, L"Operator Manager Created");
}

/// <summary>
/// Operator Manager destructor
/// </summary>
CVsiOperatorManager::~CVsiOperatorManager()
{
	_ASSERT(m_pPdm == NULL);

	ClearAuthenticatedList();

	VSI_LOG_MSG(VSI_LOG_INFO, L"Operator Manager Destroyed");
}

/// <summary>
/// The IVsiOperatorManager::Initialize method initializes the Operator Manager
/// </summary>
/// <param name="pApp">IVsiApp interface</param>
/// <param name="bServiceMode">TRUE - Service mode is on; FALSE - Service mode is off</param>
/// <returns>S_OK - The call was successful; E_FAIL - An undetermined error occurred</returns>
STDMETHODIMP CVsiOperatorManager::Initialize(IVsiApp *pApp, LPCOLESTR pszDefaultDataFile)
{
	HRESULT hr = S_OK;

	try
	{
		m_pServiceProvider = pApp;
		VSI_CHECK_INTERFACE(m_pServiceProvider, VSI_LOG_ERROR, NULL);

		m_strDefaultDataFile = pszDefaultDataFile;
		
		// Register service
		hr = m_pServiceProvider->RegisterService(
			VSI_SERVICE_OPERATOR_MGR,
			__uuidof(IVsiOperatorManager),
			static_cast<IVsiOperatorManager*>(this));
		VSI_CHECK_HR(hr, VSI_LOG_ERROR,
			L"Failure while initializing Application Controller");

		// Get PDM
		hr = m_pServiceProvider->QueryService(
			VSI_SERVICE_PDM,
			__uuidof(IVsiPdm),
			(IUnknown**)&m_pPdm);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
	}
	VSI_CATCH_(err)
	{
		hr = err;
		if (FAILED(hr))
			VSI_ERROR_LOG(err);

		Uninitialize();
	}

	return hr;
}

/// <summary>
/// The IVsiOperatorManager::Uninitialize the Operator Manager
/// </summary>
/// <returns>S_OK - The call was successful</returns>
STDMETHODIMP CVsiOperatorManager::Uninitialize()
{
	HRESULT hr;

	CCritSecLock csl(m_cs);
	m_mapIdToOperator.clear();
	m_mapNameToOperator.clear();

	m_pPdm.Release();

	if (m_pServiceProvider != NULL)
	{
		// Un-register service
		hr = m_pServiceProvider->UnRegisterService(VSI_SERVICE_OPERATOR_MGR);
		if (FAILED(hr))
			VSI_LOG_MSG(VSI_LOG_WARNING,
				L"Failure un-initializing Operator Manager");

		m_pServiceProvider.Release();
	}

	return S_OK;
}

STDMETHODIMP CVsiOperatorManager::GetOperatorDataPath(
	VSI_OPERATOR_PROP prop,
	LPCOLESTR pszProp,
	LPOLESTR *ppszPath)
{
	HRESULT hr(S_FALSE);
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(ppszPath, VSI_LOG_ERROR, L"ppszPath");

		if (!m_strDefaultDataFile.IsEmpty())
		{
			CComHeapPtr<OLECHAR> pszName;

			switch (prop)
			{
			case VSI_OPERATOR_PROP_NONE:
				break;
			case VSI_OPERATOR_PROP_NAME:
				{
					auto iter = m_mapNameToOperator.find(pszProp);
					if (iter != m_mapNameToOperator.end())
					{
						pszName.Attach(AtlAllocTaskWideString(pszProp));
					}
				}
				break;
			case VSI_OPERATOR_PROP_ID:
				{
					auto iter = m_mapIdToOperator.find(pszProp);
					if (iter != m_mapIdToOperator.end())
					{
						IVsiOperator *pOperatorTest = iter->second;

						hr = pOperatorTest->GetName(&pszName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}
				}
				break;
			default:
				_ASSERT(0);
			}

			if (NULL != pszName)
			{
				CString strPath;
				strPath.Format(
					L"%s\\%s",
					m_strDefaultDataFile.Left(m_strDefaultDataFile.ReverseFind(L'\\')),
					pszName.operator wchar_t *());
				*ppszPath = AtlAllocTaskWideString(strPath);

				hr = S_OK;
			}
			else
			{
				*ppszPath = AtlAllocTaskWideString(
					m_strDefaultDataFile.Left(
						m_strDefaultDataFile.ReverseFind(L'\\')));

				hr = S_OK;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// The IVsiOperatorManager::GetOperators method provides a enumeration object for the operators 
/// </summary>
/// <param name="ppEnum">Returns the pointer to the IVsiEnumOperator interface</param>
/// <returns>S_OK - The call was successful; E_INVALIDARG - ppEnum is invalid; E_FAIL - An undetermined error occurred</returns>
STDMETHODIMP CVsiOperatorManager::GetOperators(
	VSI_OPERATOR_FILTER_TYPE filters,
	IVsiEnumOperator **ppEnum)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	typedef CComObject<
		CComEnum<
			IVsiEnumOperator,
			&__uuidof(IVsiEnumOperator),
			IVsiOperator*, _CopyInterface<IVsiOperator>
		>
	> CVsiEnumOperator;

	int i = 0;
	std::unique_ptr<IVsiOperator*[]> ppOperators;

	try
	{
		VSI_CHECK_ARG(ppEnum, VSI_LOG_ERROR, L"ppEnum");

		if (VSI_OPERATOR_FILTER_TYPE_SERVICE_MODE & filters)
		{
			CString strHashed;
			GetHashedDate(strHashed);

			auto iter = m_mapServiceStatus.find(strHashed);
			if (m_mapServiceStatus.end() != iter)
			{
				if (VSI_SERVICE_STATUS_BLOCKED == iter->second)
				{
					filters &= ~VSI_OPERATOR_FILTER_TYPE_SERVICE_MODE;
				}
			}
		}

		DWORD dwCount(0);
		hr = GetOperatorCount(filters, &dwCount);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		ppOperators.reset(new IVsiOperator*[max(1, dwCount)]);

		for (auto iter = m_mapNameToOperator.begin();
			iter != m_mapNameToOperator.end();
			++iter)
		{
			if (IsOperatorType(iter->second, filters))
			{
				CComPtr<IVsiOperator> pOperator(iter->second);
				ppOperators[i++] = pOperator.Detach();
			}
		}

		std::unique_ptr<CVsiEnumOperator> pEnum(new CVsiEnumOperator);
		VSI_CHECK_MEM(pEnum, VSI_LOG_ERROR, L"GetOperators failed");

		hr = pEnum->Init(ppOperators.get(), ppOperators.get() + i, NULL, AtlFlagTakeOwnership);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"GetOperators failed");

		// pEnum is the owner
		ppOperators.release();

		hr = pEnum->QueryInterface(
			__uuidof(IVsiEnumOperator),
			(void**)ppEnum);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"GetOperators failed");

		pEnum.release();
	}
	VSI_CATCH(hr);

	if (NULL != ppOperators)
	{
		for (; i > 0; i--)
		{
			if (NULL != ppOperators[i])
			{
				ppOperators[i]->Release();
			}
		}
	}

	return hr;
}

/// <summary>
/// The IVsiOperatorManager::GetOperator method returns the interface of the requested operator 
/// </summary>
/// <param name="prop">Request type: VSI_OPERATOR_PROP_ID or VSI_OPERATOR_PROP_NAME</param>
/// <param name="pszProp">Request name</param>
/// <param name="ppOperator">Returns the pointer to the IVsiOperator interface</param>
/// <returns>S_OK - The call was successful; E_INVALIDARG - pszProp or ppOperator is invalid; E_FAIL - An undetermined error occurred</returns>
STDMETHODIMP CVsiOperatorManager::GetOperator(
	VSI_OPERATOR_PROP prop,
	LPCOLESTR pszProp,
	IVsiOperator **ppOperator)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszProp, VSI_LOG_ERROR, L"pszProp");
		VSI_CHECK_ARG(ppOperator, VSI_LOG_ERROR, L"ppOperator");

		switch (prop)
		{
		case VSI_OPERATOR_PROP_ID:
			{
				auto iter = m_mapIdToOperator.find(pszProp);
				if (iter != m_mapIdToOperator.end())
				{
					*ppOperator = iter->second;
					(*ppOperator)->AddRef();

					hr = S_OK;
				}
			}
			break;

		case VSI_OPERATOR_PROP_NAME:
			{
				auto iter = m_mapNameToOperator.find(pszProp);
				if (iter != m_mapNameToOperator.end())
				{
					*ppOperator = iter->second;
					(*ppOperator)->AddRef();

					hr = S_OK;
				}
			}
			break;

		default:
			_ASSERT(0);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// The IVsiOperatorManager::GetCurrentOperator method returns the interface of the logon operator 
/// </summary>
/// <param name="ppOperator">Returns the pointer to the IVsiOperator interface</param>
/// <returns>S_OK - The call was successful; E_INVALIDARG - ppOperator is invalid; E_FAIL - An undetermined error occurred</returns>
STDMETHODIMP CVsiOperatorManager::GetCurrentOperator(IVsiOperator **ppOperator)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(ppOperator, VSI_LOG_ERROR, L"ppOperator");

		auto iter = m_mapIdToOperator.find(m_strCurrentOperator);
		if (iter != m_mapIdToOperator.end())
		{
			*ppOperator = iter->second;
			(*ppOperator)->AddRef();

			hr = S_OK;
		}
		else
		{
			*ppOperator = NULL;
			hr = S_FALSE;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// The IVsiOperatorManager::SetCurrentOperator method set the current logon operator 
/// </summary>
/// <param name="pszId">Identifier of the operator, NULL to logoff</param>
/// <returns>S_OK - The call was successful; E_FAIL - An undetermined error occurred</returns>
STDMETHODIMP CVsiOperatorManager::SetCurrentOperator(LPCOLESTR pszId)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		if (NULL != pszId)
		{
			auto iter = m_mapIdToOperator.find(pszId);
			VSI_VERIFY(iter != m_mapIdToOperator.end(), VSI_LOG_ERROR, NULL);

			// Update session info
			VSI_OPERATOR_SESSION_INFO si = { 0 };

			hr = iter->second->GetSessionInfo(&si);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			GetSystemTime(&si.stLastLogin);
			ZeroMemory(&si.stLastLogout, sizeof(SYSTEMTIME));

			hr = iter->second->SetSessionInfo(&si);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Authenticated List
			ClearAuthenticatedList();

			// Save current operator
			m_strCurrentOperator = pszId;

			// Updates PDM
			UpdatePdm(iter->second);

			bool bDirty = false;
			if (IsOperatorType(iter->second, VSI_OPERATOR_FILTER_TYPE_SERVICE_MODE))
			{
				CString strHashed;
				GetHashedDate(strHashed);

				auto iter = m_mapServiceStatus.find(strHashed);
				if (m_mapServiceStatus.end() == iter)
				{
					m_mapServiceStatus[strHashed] = VSI_SERVICE_STATUS_ALLOWED;
					bDirty = true;
				}
			}

			// Check whether we need to expire service mode password
			if (0 < m_mapServiceStatus.size())
			{
				CString strHashed;
				GetHashedDate(strHashed);

				for (auto iter = m_mapServiceStatus.begin(); iter != m_mapServiceStatus.end(); ++iter)
				{
					if (iter->second == VSI_SERVICE_STATUS_ALLOWED)
					{
						if (iter->first != strHashed)
						{
							iter->second = VSI_SERVICE_STATUS_BLOCKED;
							bDirty = true;
						}
					}
				}

				if (bDirty)
				{
					SaveOperator(NULL);
				}
			}
		}
		else
		{
			if (!m_strCurrentOperator.IsEmpty())
			{
				auto iter = m_mapIdToOperator.find(m_strCurrentOperator);
				if (iter != m_mapIdToOperator.end())
				{
					// Update session info
					VSI_OPERATOR_SESSION_INFO si = { 0 };

					hr = iter->second->GetSessionInfo(&si);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					GetSystemTime(&si.stLastLogout);

					FILETIME ftLogin = { 0 };
					SystemTimeToFileTime(&si.stLastLogin, &ftLogin);

					ULARGE_INTEGER uliLogin = { ftLogin.dwLowDateTime, ftLogin.dwHighDateTime };

					FILETIME ftLogout = { 0 };
					SystemTimeToFileTime(&si.stLastLogout, &ftLogout);

					ULARGE_INTEGER uliLogout = { ftLogout.dwLowDateTime, ftLogout.dwHighDateTime };

					if (uliLogout.QuadPart > uliLogin.QuadPart)
					{
						ULARGE_INTEGER uliDiff;
						uliDiff.QuadPart = uliLogout.QuadPart - uliLogin.QuadPart;
						uliDiff.QuadPart = (ULONGLONG)(uliDiff.QuadPart / 1E7);

						si.dwTotalSeconds += uliDiff.LowPart;
					}

					hr = iter->second->SetSessionInfo(&si);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}

			// Authenticated List
			ClearAuthenticatedList();

			// Remove current operator
			m_strCurrentOperator.Empty();

			// Updates PDM
			UpdatePdm(NULL);
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// Get the current logged in operator data path.
/// </summary>
/// <param name="ppszPath">Data path output.</param>
/// <returns>S_OK - The call was successful; S_FALSE - No current operator; E_FAIL - An undetermined error occurred</returns>
STDMETHODIMP CVsiOperatorManager::GetCurrentOperatorDataPath(LPOLESTR *ppszPath)
{
	HRESULT hr(S_FALSE);
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(ppszPath, VSI_LOG_ERROR, L"ppszPath");

		CString strPath;

		CComPtr<IVsiOperator> pOperatorCurrent;
		hr = GetCurrentOperator(&pOperatorCurrent);
		VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

		VSI_OPERATOR_TYPE dwType;
		hr = pOperatorCurrent->GetType(&dwType);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		BOOL bServiceAccount = 0 != (VSI_OPERATOR_TYPE_SERVICE_MODE & dwType);			

		if (bServiceAccount)
		{
			// Get a temp path
			WCHAR szTempPath[MAX_PATH];
			int iRet = GetTempPath(_countof(szTempPath), szTempPath);
			VSI_WIN32_VERIFY(0 != iRet, VSI_LOG_ERROR, NULL);

			PathAppend(szTempPath, L"ServiceAccount");

			*ppszPath = AtlAllocTaskWideString(szTempPath);

			BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
				VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
			if (bSecureMode)
			{
				hr = ValidateOperatorSetup(NULL);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}

			hr = S_OK;
		}
		else
		{
			if (!m_strCurrentOperator.IsEmpty())
			{
				auto iter = m_mapIdToOperator.find(m_strCurrentOperator);
				if (iter != m_mapIdToOperator.end())
				{
					CComHeapPtr<OLECHAR> pszOperatorName;
					hr = iter->second->GetName(&pszOperatorName);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					strPath.Format(
						L"%s\\%s",
						m_strDefaultDataFile.Left(m_strDefaultDataFile.ReverseFind(L'\\')).GetString(),
						pszOperatorName);
					*ppszPath = AtlAllocTaskWideString(strPath);

					BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
						VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
						VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
					if (bSecureMode)
					{
						hr = ValidateOperatorSetup(pszOperatorName);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					}

					hr = S_OK;
				}
				else
				{
					hr = S_FALSE;
				}
			}
			else
			{
				hr = S_FALSE;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorManager::GetRelationshipWithCurrentOperator(
	VSI_OPERATOR_PROP prop,
	LPCOLESTR pszProp,
	VSI_OPERATOR_RELATIONSHIP *pRelationship)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszProp, VSI_LOG_ERROR, NULL);
		VSI_CHECK_ARG(pRelationship, VSI_LOG_ERROR, NULL);

		*pRelationship = VSI_OPERATOR_RELATIONSHIP_NONE;

		if (m_strCurrentOperator.IsEmpty())
		{
			hr = S_FALSE;
		}
		else
		{
			CComHeapPtr<OLECHAR> pszId;
			IVsiOperator *pOperatorTest(NULL);

			switch (prop)
			{
			case VSI_OPERATOR_PROP_NAME:
				{
					auto iter = m_mapNameToOperator.find(pszProp);
					if (iter != m_mapNameToOperator.end())
					{
						pOperatorTest = iter->second;

						hr = pOperatorTest->GetId(&pszId);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						pszProp = pszId;
					}
				}
				break;
			case VSI_OPERATOR_PROP_ID:
				{
					auto iter = m_mapIdToOperator.find(pszProp);
					if (iter != m_mapIdToOperator.end())
					{
						pOperatorTest = iter->second;
					}
				}
				break;
			default:
				_ASSERT(0);
			}

			if (m_strCurrentOperator == pszProp)
			{
				*pRelationship = VSI_OPERATOR_RELATIONSHIP_SAME;
			}
			else if (NULL != pOperatorTest)
			{
				IVsiOperator *pOperatorCurrent(NULL);
				auto iter = m_mapIdToOperator.find(m_strCurrentOperator);
				if (iter != m_mapIdToOperator.end())
				{
					pOperatorCurrent = iter->second;
				}

				if (NULL != pOperatorCurrent)
				{
					CComHeapPtr<OLECHAR> pszGroupCurrent;
					hr = pOperatorCurrent->GetGroup(&pszGroupCurrent);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
					if (S_OK == hr && (0 != *pszGroupCurrent))
					{
						CComHeapPtr<OLECHAR> pszGroupTest;
						hr = pOperatorTest->GetGroup(&pszGroupTest);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (S_OK == hr)
						{
							if (0 == wcscmp(pszGroupCurrent, pszGroupTest))
							{
								*pRelationship = VSI_OPERATOR_RELATIONSHIP_SAME_GROUP;
							}
						}
					}
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorManager::AddOperator(
	IVsiOperator *pOperator, LPCOLESTR pszIdCloneSettingsFrom)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pOperator, VSI_LOG_ERROR, L"pOperator");
		VSI_CHECK_INTERFACE(m_pPdm, VSI_LOG_ERROR, L"m_pPdm");

		CComHeapPtr<OLECHAR> pszId;
		CComHeapPtr<OLECHAR> pszName;

		hr = pOperator->GetId(&pszId);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		hr = pOperator->GetName(&pszName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		m_mapIdToOperator[pszId.operator wchar_t *()] = pOperator;
		m_mapNameToOperator[pszName.operator wchar_t *()] = pOperator;

		// Clone settings from another operator
		if (NULL != pszIdCloneSettingsFrom && 0 != *pszIdCloneSettingsFrom)
		{
			CComPtr<IVsiOperator> pOperatorClone;
			hr = GetOperator(VSI_OPERATOR_PROP_ID, pszIdCloneSettingsFrom, &pOperatorClone);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			CComHeapPtr<OLECHAR> pszNameClone;
			hr = pOperatorClone->GetName(&pszNameClone);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				
			CString strSource;
			strSource.Format(
				L"%s\\%s",
				m_strDefaultDataFile.Left(m_strDefaultDataFile.ReverseFind(L'\\')).GetString(),
				pszNameClone.operator wchar_t *());
 
			if (0 == _waccess(strSource, 0))
			{
				CString strTarget;
				strTarget.Format(
					L"%s\\%s",
					m_strDefaultDataFile.Left(m_strDefaultDataFile.ReverseFind(L'\\')).GetString(),
					pszName.operator wchar_t *());
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
 
				VSI_LOG_MSG(VSI_LOG_INFO, L"Clone operator settings:");
				VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"  Source: %s", strSource.GetString()));
				VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"  Target: %s", strTarget.GetString()));
 
				// Copy folder
				BOOL bRet = VsiCopyDirectory(strSource, strTarget);
				VSI_VERIFY(bRet, VSI_LOG_ERROR, L"Clone operator settings failed!");
			} 
		}
		else  // Set up User settings
		{
			BOOL bSecureMode = VsiGetBooleanValue<BOOL>(
				VSI_PDM_ROOT_APP, VSI_PDM_GROUP_SETTINGS, 
				VSI_PARAMETER_SYS_SECURE_MODE, m_pPdm);
			if (bSecureMode)
			{
				hr = ValidateOperatorSetup(pszName);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			}
		}
	
		VsiSetParamEvent(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
			VSI_PARAMETER_EVENTS_GENERAL_OPERATOR_LIST_UPDATE, m_pPdm);

		m_bModified = TRUE;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorManager::RemoveOperator(LPCOLESTR pszId)
{
	HRESULT hr = S_FALSE;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszId, VSI_LOG_ERROR, L"pszId");
		VSI_CHECK_INTERFACE(m_pPdm, VSI_LOG_ERROR, L"m_pPdm");

		auto iter = m_mapIdToOperator.find(pszId);
		if (iter != m_mapIdToOperator.end())
		{
			CComHeapPtr<OLECHAR> pszName;
			CComPtr<IVsiOperator> pOperator(iter->second);
			pOperator->GetName(&pszName);

			m_mapIdToOperator.erase(iter);

			// This is the active operator?
			if (m_strCurrentOperator == pszId)
			{
				m_strCurrentOperator.Empty();
				UpdatePdm(NULL);
			}

			CString strName(pszName);
			auto iter2 = m_mapNameToOperator.find(strName);
			if (iter2 != m_mapNameToOperator.end())
			{
				m_mapNameToOperator.erase(iter2);
			}

			// Change operator folder to be .bak
			CString strSource;
			strSource.Format(
				L"%s\\%s",
				m_strDefaultDataFile.Left(m_strDefaultDataFile.ReverseFind(L'\\')).GetString(),
				strName.GetString());

			if (0 == _waccess(strSource, 0))
			{
				CString strTarget;
				CString strPath = m_strDefaultDataFile.Left(m_strDefaultDataFile.ReverseFind(L'\\'));
				hr = VsiGetUniqueFolderName(strPath, strName, L".bak", strTarget.GetBufferSetLength(MAX_PATH), MAX_PATH);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				strTarget.ReleaseBuffer();

				VSI_LOG_MSG(VSI_LOG_INFO, L"Backing up old operator settings:");
				VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"  Source: %s", strSource.GetString()));
				VSI_LOG_MSG(VSI_LOG_INFO, VsiFormatMsg(L"  Target: %s", strTarget.GetString()));

				BOOL bRet = MoveFile(strSource, strTarget);
				VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, L"Backing up failed!");
			}

			VsiSetParamEvent(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
				VSI_PARAMETER_EVENTS_GENERAL_OPERATOR_LIST_UPDATE, m_pPdm);

			m_bModified = TRUE;

			hr = S_OK;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorManager::OperatorModified()
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_INTERFACE(m_pPdm, VSI_LOG_ERROR, L"m_pPdm");

		VsiSetParamEvent(VSI_PDM_ROOT_APP, VSI_PDM_GROUP_EVENTS,
			VSI_PARAMETER_EVENTS_GENERAL_OPERATOR_LIST_UPDATE, m_pPdm);

		m_bModified = TRUE;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorManager::GetOperatorCount(VSI_OPERATOR_FILTER_TYPE filters, DWORD *pdwCount)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pdwCount, VSI_LOG_ERROR, L"pdwCount");

		*pdwCount = 0;

		DWORD dwCount(0);

		for (auto iter = m_mapIdToOperator.begin();
			iter != m_mapIdToOperator.end();
			++iter)
		{
			if (IsOperatorType(iter->second, filters))
			{
				++dwCount;
			}
		}

		*pdwCount = dwCount;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorManager::SetIsAuthenticated(
	LPCOLESTR pszId, BOOL bAuthenticated)
{
	HRESULT hr(S_OK);
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pszId, VSI_LOG_ERROR, NULL);

		auto iterOperator = m_mapIdToOperator.find(pszId);
		if (iterOperator != m_mapIdToOperator.end())
		{
			if (bAuthenticated)
			{
				SYSTEMTIME st;
				GetSystemTime(&st);

				// Save the time when it was authenticated
				m_mapAuthenticated[(LPCWSTR)pszId] = st;

				if (!m_bAdminAuthenticated)
				{
					// Checks whether it's Admin
					VSI_OPERATOR_TYPE dwType(VSI_OPERATOR_TYPE_NONE);
					IVsiOperator *pOperator = iterOperator->second;
					hr = pOperator->GetType(&dwType);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (VSI_OPERATOR_TYPE_ADMIN & dwType)
					{
						m_bAdminAuthenticated = TRUE;
					}
				}
			}
			else
			{
				auto iterLogin = m_mapAuthenticated.find(pszId);
				if (iterLogin != m_mapAuthenticated.end())
				{
					m_mapAuthenticated.erase(iterLogin);
				}

				m_bAdminAuthenticated = FALSE;

				for (auto iter = m_mapAuthenticated.begin();
					iter != m_mapAuthenticated.end();
					++iter)
				{
					auto iterOperator = m_mapIdToOperator.find(iter->first);
					if (iterOperator != m_mapIdToOperator.end())
					{
						IVsiOperator *pOperator = iterOperator->second;

						VSI_OPERATOR_TYPE dwType(VSI_OPERATOR_TYPE_NONE);
						hr = pOperator->GetType(&dwType);
						VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

						if (VSI_OPERATOR_TYPE_ADMIN & dwType)
						{
							m_bAdminAuthenticated = TRUE;
							break;
						}
					}
				}
			}
		}
		else
		{
			// Operator id not valid
			hr = S_FALSE;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorManager::GetIsAuthenticated(LPCOLESTR pszId, BOOL *pbAuthenticated)
{
	HRESULT hr = S_OK;
	CCritSecLock csl(m_cs);

	try
	{
		VSI_CHECK_ARG(pbAuthenticated, VSI_LOG_ERROR, NULL);

		*pbAuthenticated = FALSE;

		auto iterLogin = m_mapAuthenticated.find(pszId);
		if (iterLogin != m_mapAuthenticated.end())
		{
			*pbAuthenticated = TRUE;
		}
		else
		{
			// Check whether it is an valid operator
			auto iter = m_mapIdToOperator.find(pszId);
			if (iter == m_mapIdToOperator.end())
			{
				hr = S_FALSE;
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorManager::GetIsAdminAuthenticated(BOOL *pbAuthenticated)
{
	HRESULT hr = S_OK;

	try
	{
		VSI_CHECK_ARG(pbAuthenticated, VSI_LOG_ERROR, NULL);

		*pbAuthenticated = m_bAdminAuthenticated;
	}
	VSI_CATCH(hr);

	return hr;
}

STDMETHODIMP CVsiOperatorManager::ClearAuthenticatedList()
{
	CCritSecLock csl(m_cs);

	m_mapAuthenticated.clear();
	
	m_bAdminAuthenticated = FALSE;

	return S_OK;
}

STDMETHODIMP CVsiOperatorManager::IsDirty()
{
	return m_bModified ? S_OK : S_FALSE;
}

/// <summary>
/// CVsiOperatorManager::LoadOperator
/// </summary>
STDMETHODIMP CVsiOperatorManager::LoadOperator(LPCOLESTR pszDataFile)
{
	HRESULT hrRet = S_OK;
	HRESULT hr = S_OK;
	CComPtr<IXMLDOMDocument> pXmlDocOperator;
	CComQIPtr<IXMLDOMElement> pXMLElem;
	CComQIPtr<IXMLDOMElement> pXMLElemValue;
	CComPtr<IXMLDOMNodeList> pXMLNodeList;
	CComPtr<IXMLDOMNode> pXMLNode;
	CComVariant vtNodeValue;

	CCritSecLock cs(m_cs);

	try
	{
		VSI_LOG_MSG(VSI_LOG_INFO, L"Load operator started");

		if (NULL == pszDataFile)
		{
			pszDataFile = m_strDefaultDataFile;
		}

		hr = VsiCreateDOMDocument(&pXmlDocOperator);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Load operator failed");

		// Load the XML document
		VARIANT_BOOL bLoaded = VARIANT_FALSE;
		hr = pXmlDocOperator->load(CComVariant(pszDataFile), &bLoaded);
		// Handle error
		if (hr != S_OK || bLoaded == VARIANT_FALSE)
		{
			HRESULT hr1;
			CComPtr<IXMLDOMParseError> pParseError;
			hr1 = pXmlDocOperator->get_parseError(&pParseError);
			if (SUCCEEDED(hr1))
			{
				long lError;
				hr1 = pParseError->get_errorCode(&lError);
				CComBSTR bstrDesc;
				hr1 = pParseError->get_reason(&bstrDesc);
				if (SUCCEEDED(hr1))
				{
					// Display error
					VSI_LOG_MSG(VSI_LOG_WARNING,
						VsiFormatMsg(L"Read operator file '%s' failed - %s", pszDataFile, bstrDesc));
				}
				else
				{
					VSI_LOG_MSG(VSI_LOG_WARNING,
						VsiFormatMsg(L"Read operator file '%s' failed", pszDataFile));
				}
			}
			else
			{
				VSI_LOG_MSG(VSI_LOG_WARNING,
					VsiFormatMsg(L"Read operator file '%s' failed", pszDataFile));
			}

			hrRet = hr;
		}
		else
		{
			ClearAuthenticatedList();

			m_mapIdToOperator.clear();
			m_mapNameToOperator.clear();

			// find the root node in the XML document
			CComPtr<IXMLDOMElement> pXmlElemRoot;
			hr = pXmlDocOperator->get_documentElement(&pXmlElemRoot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Load operator failed");

			LoadOperators(pXmlElemRoot);
		}

		// Service Mode account
		{
			CComPtr<IVsiOperator> pOperator;

			hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"LoadOperator - failure creating operator");

			CString strName(MAKEINTRESOURCE(IDS_LOGIN_ACCOUNT_SERVICE_MODE));

			hr = pOperator->SetId(VSI_OPERATOR_ID_SERVICE_MODE);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pOperator->SetName(strName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pOperator->SetType(VSI_OPERATOR_TYPE_SERVICE_MODE | VSI_OPERATOR_TYPE_ADMIN);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_mapIdToOperator[VSI_OPERATOR_ID_SERVICE_MODE] = pOperator;
			m_mapNameToOperator[strName] = pOperator;
		}

		m_bModified = FALSE;
	}
	VSI_CATCH(hr);

	return SUCCEEDED(hr) ? hrRet : hr;
}

bool CVsiOperatorManager::IsOperatorType(IVsiOperator *pOperator, VSI_OPERATOR_FILTER_TYPE filters)
{
	bool bIsType(false);
	HRESULT hr(S_OK);

	try
	{
		VSI_OPERATOR_TYPE dwType;
		hr = pOperator->GetType(&dwType);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
			
		if (VSI_OPERATOR_TYPE_SERVICE_MODE & dwType)
		{
			if (VSI_OPERATOR_FILTER_TYPE_SERVICE_MODE & filters)
			{
				bIsType = true;
			}
		}
		else
		{
			if ((VSI_OPERATOR_FILTER_TYPE_STANDARD & filters) &&
				(dwType & VSI_OPERATOR_TYPE_STANDARD))
			{
				if (VSI_OPERATOR_FILTER_TYPE_PASSWORD & filters)
				{
					CComHeapPtr<OLECHAR> pszPassword;
					hr = pOperator->HasPassword();
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						bIsType = true;
					}
				}
				else if (VSI_OPERATOR_FILTER_TYPE_NO_PASSWORD & filters)
				{
					CComHeapPtr<OLECHAR> pszPassword;
					hr = pOperator->HasPassword();

					if (S_FALSE == hr)
					{
						bIsType = true;
					}
				}
				else
				{
					bIsType = true;
				}
			}
			else if ((VSI_OPERATOR_FILTER_TYPE_ADMIN & filters) &&
				(dwType & VSI_OPERATOR_TYPE_ADMIN))
			{
				if (VSI_OPERATOR_FILTER_TYPE_PASSWORD & filters)
				{
					CComHeapPtr<OLECHAR> pszPassword;
					hr = pOperator->HasPassword();
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

					if (S_OK == hr)
					{
						bIsType = true;
					}
				}
				else if (VSI_OPERATOR_FILTER_TYPE_NO_PASSWORD & filters)
				{
					CComHeapPtr<OLECHAR> pszPassword;
					hr = pOperator->HasPassword();

					if (S_FALSE == hr)
					{
						bIsType = true;
					}
				}
				else
				{
					bIsType = true;
				}
			}
		}
	}
	VSI_CATCH(hr);

	return bIsType;
}

/// <summary>
/// CVsiOperatorManager::LoadOperators
/// </summary>
HRESULT CVsiOperatorManager::LoadOperators(IXMLDOMElement* pXmlElemRoot)
{
	HRESULT hr = S_OK;
	CComPtr<IXMLDOMNodeList> pXmlNodeList;

	CCritSecLock cs(m_cs);

	try
	{
		// Settings
		VSI_LOG_MSG(VSI_LOG_INFO, L"Load operators list XML started");

		// Check if root element is VSI_OPERATOR_ELM_OPERATORS
		CComBSTR bstrName;
		hr = pXmlElemRoot->get_tagName(&bstrName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		if (0 != lstrcmpi(bstrName, VSI_OPERATOR_ELM_OPERATORS))
		{
			VSI_FAIL(VSI_LOG_ERROR, NULL);
		}

		CComVariant vVersion;
		hr = pXmlElemRoot->getAttribute(CComBSTR(VSI_OPERATOR_VERSION), &vVersion);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		bool bPasswordEncrypted = (S_OK == hr);

		// Retrieve service records
		CComVariant vServiceRecords;
		hr = pXmlElemRoot->getAttribute(CComBSTR(VSI_OPERATOR_ATT_SERVICE_RECORD), &vServiceRecords);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		CString strServiceRecords = (LPCWSTR)V_BSTR(&vServiceRecords);

		// Parse service records
		m_mapServiceStatus.clear();
		int iPos = 0;
		CString strServiceRecord = strServiceRecords.Tokenize(L",", iPos);
		while (!strServiceRecord.IsEmpty())
		{
			m_mapServiceStatus[strServiceRecord.Left(40)] =
				strServiceRecord.Right(1) == L"A" ? VSI_SERVICE_STATUS_ALLOWED : VSI_SERVICE_STATUS_BLOCKED;
			strServiceRecord = strServiceRecords.Tokenize(L",", iPos);
		} 

		hr = pXmlElemRoot->selectNodes(CComBSTR(VSI_OPERATOR_ELM_OPERATOR), &pXmlNodeList);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		LoadOperatorSettings(pXmlNodeList, bPasswordEncrypted);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		VSI_LOG_MSG(VSI_LOG_INFO, L"Load operators list XML completed");
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// Load an operator settings
/// </summary>
HRESULT CVsiOperatorManager::LoadOperatorSettings(IXMLDOMNodeList* pXmlNodeList, bool bPasswordEncrypted)
{
	HRESULT hr = S_OK;
	CComQIPtr<IXMLDOMElement> pXmlElem;
	CComQIPtr<IXMLDOMElement> pXmlElemValue;
	CComPtr<IXMLDOMNode> pXmlNode;
	CComVariant vtNodeValue;
	CComPtr<IVsiOperator> pOperator;
	long lCount = 0;

	CCritSecLock cs(m_cs);

	try
	{
		hr = pXmlNodeList->get_length(&lCount);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"LoadOperator - failure getting number of elements");

		// Retrieve each element in the list.
		for (long i = 0; i < lCount; i++)
		{
			pXmlNode.Release();
			hr = pXmlNodeList->get_item(i, &pXmlNode);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"LoadOperator - failure getting nodes");

			// Add to operator map
			pOperator.Release();
			hr = pOperator.CoCreateInstance(__uuidof(VsiOperator));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"LoadOperator - failure creating operator");

			CComQIPtr<IVsiPersistOperator> pOperatorXml(pOperator);
			if (pOperatorXml != NULL)  // optional
			{
				hr = pOperatorXml->LoadXml(pXmlNode);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"LoadOperator - failure loading operator");
			}

			if (!bPasswordEncrypted)
			{
				// Encrypt the password
				CComHeapPtr<OLECHAR> pszPasswordNotEncrypted;
				hr = pOperator->GetPassword(&pszPasswordNotEncrypted);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				if (0 != *pszPasswordNotEncrypted)
				{
					hr = pOperator->SetPassword(pszPasswordNotEncrypted, FALSE);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
				}
			}

			CComHeapPtr<OLECHAR> pszId;
			CComHeapPtr<OLECHAR> pszName;

			hr = pOperator->GetId(&pszId);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pOperator->GetName(&pszName);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			m_mapIdToOperator[pszId.operator wchar_t *()] = pOperator;
			m_mapNameToOperator[pszName.operator wchar_t *()] = pOperator;
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// CVsiOperatorManager::SaveOperator
/// </summary>
/// <returns>
///		E_ACCESSDENIED
///		HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD)
/// </returns>
STDMETHODIMP CVsiOperatorManager::SaveOperator(LPCOLESTR pszDataFile)
{
	HRESULT hr = S_OK;
	BOOL bBackup = FALSE;
	BOOL bRestoreOnFail = FALSE;
	WCHAR szBackup[MAX_PATH];
	BOOL bRet;
	HANDLE hTransaction = INVALID_HANDLE_VALUE;

	CCritSecLock cs(m_cs);

	try
	{
		*szBackup = 0;

		if (NULL == pszDataFile)
		{
			pszDataFile = m_strDefaultDataFile;
		}

		VSI_VERIFY(0 != *pszDataFile, VSI_LOG_ERROR, NULL);

		// This function will return NULL in WinXP, which is fine. The API will 
		// revert to the non-transacted methods internally if necessary.
		hTransaction = VsiTxF::VsiCreateTransaction(
			NULL,
			0,		
			TRANSACTION_DO_NOT_PROMOTE,	
			0,		
			0,		
			INFINITE,
			L"Operator xml save transaction");
		VSI_WIN32_VERIFY(INVALID_HANDLE_VALUE != hTransaction, VSI_LOG_ERROR, L"Failure creating transaction");

		// Backup old
		if (0 == _waccess(pszDataFile, 0))
		{
			// Backup path
			errno_t err = wcscpy_s(szBackup, pszDataFile);
			VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");

			err = wcscat_s(szBackup, L".bak");
			VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");

			bRet = VsiTxF::VsiCopyFileTransacted(
				pszDataFile,
				szBackup,
				NULL,
				NULL,
				NULL,
				0,
				hTransaction);
			VSI_WIN32_VERIFY(bRet, VSI_LOG_ERROR, NULL);

			bBackup = TRUE;
		}

		// Create a temporary file using the DocOperator, then use the transacted 
		// file methods to move the temp file back to the original.
		// Use commutl to create unique temp file.
		WCHAR szTempFileName[VSI_MAX_PATH];
		errno_t err = wcscpy_s(szTempFileName, pszDataFile);
		VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");
		VsiPathRemoveExtension(szTempFileName);

		WCHAR szUniqueName[VSI_MAX_PATH];
		hr = VsiCreateUniqueName(L"", L"", szUniqueName, VSI_MAX_PATH);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure getting unique name");

		err = wcscat_s(szTempFileName, szUniqueName);
		VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");
		err = wcscat_s(szTempFileName, L".xml");
		VSI_VERIFY(0 == err, VSI_LOG_ERROR, L"Failure copying string");

		// Saves operator list
		{
			CComPtr<IXMLDOMDocument> pXmlDocOperator;
			CComPtr<IXMLDOMElement> pElmRoot;
			CComQIPtr<IXMLDOMElement> pXMLElem;
			CComQIPtr<IXMLDOMElement> pXMLElemValue;
			CComPtr<IXMLDOMNodeList> pXMLNodeList;
			CComPtr<IXMLDOMNode> pXMLNode;
			CComVariant vtNodeValue;

			// Create XML doc
			hr = VsiCreateDOMDocument(&pXmlDocOperator);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"Failure creating DOMDocument");

			CComPtr<IXMLDOMProcessingInstruction> pPI;
			hr = pXmlDocOperator->createProcessingInstruction(
				CComBSTR(L"xml"), CComBSTR(L"version='1.0' encoding='utf-8'"), &pPI);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			hr = pXmlDocOperator->appendChild(pPI, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			// Operator Root element
			hr = pXmlDocOperator->createElement(CComBSTR(VSI_OPERATOR_ELM_OPERATORS), &pElmRoot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	VsiFormatMsg(L"Failure creating '%s' element", VSI_OPERATOR_ELM_OPERATORS));

			hr = pElmRoot->setAttribute(CComBSTR(VSI_OPERATOR_VERSION), CComVariant(VSI_OPERATOR_VERSION_LATEST));
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);

			if (0 < m_mapServiceStatus.size())
			{
				CString strServiceRecord;
				for (auto iter = m_mapServiceStatus.begin(); iter != m_mapServiceStatus.end(); ++iter)
				{
					if (!strServiceRecord.IsEmpty())
					{
						strServiceRecord += L",";
					}

					strServiceRecord += iter->first;
					strServiceRecord += (VSI_SERVICE_STATUS_ALLOWED == iter->second) ? L"A" : L"B";
				}

				hr = pElmRoot->setAttribute(CComBSTR(VSI_OPERATOR_ATT_SERVICE_RECORD), CComVariant(strServiceRecord));
				VSI_CHECK_HR(hr, VSI_LOG_ERROR,	NULL);
			}

			hr = SaveOperators(pElmRoot);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	VsiFormatMsg(L"Failure appending '%s' element", VSI_OPERATOR_ELM_OPERATORS));

			hr = pXmlDocOperator->appendChild(pElmRoot, NULL);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	VsiFormatMsg(L"Failure appending '%s' element", VSI_OPERATOR_ELM_OPERATORS));

			// Save file
			CComVariant vt(szTempFileName);
			hr = pXmlDocOperator->save(vt);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR,	VsiFormatMsg(L"Failure saving operator file - %s", pszDataFile));
		}
		
		bRet = VsiTxF::VsiMoveFileTransacted(
			szTempFileName,
			pszDataFile,	
			NULL,	
			NULL,	
			MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING,
			hTransaction);
		if (!bRet && (NULL == hTransaction))
		{
			// Only want to restore on fail when we are not using transacted methods 
			// and the move fails.
			bRestoreOnFail = TRUE;
		}
		VSI_VERIFY(bRet, VSI_LOG_ERROR, L"Create transaction failed");

		bRet = VsiTxF::VsiCommitTransaction(hTransaction);
		VSI_VERIFY(bRet, VSI_LOG_ERROR, L"Create transaction failed");

		m_bModified = FALSE;
	}
	VSI_CATCH(hr);

	if (INVALID_HANDLE_VALUE != hTransaction)
	{
		BOOL bRet = VsiTxF::VsiCloseTransaction(hTransaction);
		if (!bRet)
		{
			VSI_LOG_MSG(VSI_LOG_ERROR, L"Failure closing transaction");
		}
	}

	if (FAILED(hr))
	{
		bRet = TRUE;
		CString strMsg;

		if (bBackup && bRestoreOnFail)
		{
			// Restores file
			if (0 == _waccess(pszDataFile, 0))
			{
				bRet = DeleteFile(pszDataFile);
				if (bRet)
				{
					bRet = MoveFile(szBackup, pszDataFile);
				}
			}
		}

		if (bRet)
		{
			strMsg = L"There is an error in saving the user configuration settings.";
		}
		else
		{
			// Very bad!!!
			strMsg = L"There is an error in recovering the user configuration settings.\n"
				L"Save your data and shut down the system.\n";
		}

		VSI_LOG_MSG(VSI_LOG_INFO, strMsg);
	}

	return hr;
}

/// <summary>
/// CVsiOperatorManager::SaveOperators
/// </summary>
HRESULT CVsiOperatorManager::SaveOperators(IXMLDOMElement* pXmlElemRoot)
{
	HRESULT hr = S_OK;

	try
	{
		auto iterOperator = m_mapIdToOperator.begin();

		for (; iterOperator != m_mapIdToOperator.end(); ++iterOperator)
		{
			IVsiOperator *pOperator = iterOperator->second;

			CComQIPtr<IVsiPersistOperator> pOperatorXml(pOperator);
			if (pOperatorXml != NULL)  // Optional
			{
				CComHeapPtr<OLECHAR> pszId;
				hr = pOperator->GetId(&pszId);
				VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

				// Skip service mode 
				if (0 != wcscmp(pszId, VSI_OPERATOR_ID_SERVICE_MODE))
				{
					hr = pOperatorXml->SaveXml(pXmlElemRoot);
					VSI_CHECK_HR(hr, VSI_LOG_ERROR, L"SaveXml - failure loading operator");
				}
			}
		}
	}
	VSI_CATCH(hr);

	return hr;
}

/// <summary>
/// Sets active operator name to PDM
/// <summary>
void CVsiOperatorManager::UpdatePdm(IVsiOperator *pOperator) throw(...)
{
	HRESULT hr;
	CComVariant vName;

	// Updates PDM
	CComPtr<IVsiParameterRange> pParamOperator;
	hr = m_pServiceProvider->QueryService(
		VSI_SERVICE_PDM L"/" VSI_PDM_ROOT_APP L"/"
			VSI_PDM_GROUP_SETTINGS L"/" VSI_PARAMETER_SYS_OPERATOR,
		__uuidof(IVsiParameterRange),
		(IUnknown**)&pParamOperator);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	if (NULL != pOperator)
	{
		// Gets name
		CComHeapPtr<OLECHAR> pszName;
		hr = pOperator->GetName(&pszName);
		VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

		vName = pszName;
	}
	else
	{
		vName = L"";
	}

	hr = pParamOperator->SetValue(vName, FALSE);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	hr = OperatorModified();
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);
}

void CVsiOperatorManager::GetHashedDate(CString &strHashed) throw(...)
{
	SYSTEMTIME stLocal;
	GetLocalTime(&stLocal);

	CString strDate;
	strDate.Format(L"%4d%2d%2d", stLocal.wYear, stLocal.wMonth, stLocal.wDay);

	HRESULT hr = VsiCreateSha1(
		(const BYTE*)strDate.GetString(),
		strDate.GetLength() * sizeof(WCHAR),
		strHashed.GetBufferSetLength(41),
		41);
	VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

	strHashed.ReleaseBuffer();
}

// Validate operator folder contains custom measurements and presets configuration files
HRESULT CVsiOperatorManager::ValidateOperatorSetup(LPCWSTR pszName)
{
	HRESULT hr = S_OK;
	HANDLE hFile(INVALID_HANDLE_VALUE);

	try
	{
		CString strOperatorsPath(m_strDefaultDataFile.Left(m_strDefaultDataFile.ReverseFind(L'\\')).GetString());
		CString strDataPath(strOperatorsPath.Left(strOperatorsPath.ReverseFind(L'\\')).GetString());

		// Get the directory where the custom measurement packages are stored - to copy to operator's folder
		CString strCustomMeasurementsDir;
		strCustomMeasurementsDir.Format(L"%s\\%s", strDataPath.GetString(), VSI_FOLDER_MEASUREMENTS);

		if (0 != _waccess(strCustomMeasurementsDir, 0))
		{
			// Use the measurement folder under the exe path
			WCHAR szPath[MAX_PATH];
			GetModuleFileName(NULL, szPath, _countof(szPath));

			LPWSTR pszSpt = wcsrchr(szPath, L'\\');
			*(pszSpt + 1) = 0;

			PathAppend(szPath, VSI_FOLDER_MEASUREMENTS);
			strCustomMeasurementsDir = szPath;
		}

		strCustomMeasurementsDir += L"\\";

		// Get the directory where the factory default presets are stored - to copy to operator's folder
		WCHAR szPath[MAX_PATH];
		GetModuleFileName(NULL, szPath, _countof(szPath));

		LPWSTR pszSpt = wcsrchr(szPath, L'\\');
		*(pszSpt + 1) = 0;
		PathAppend(szPath, VSI_FOLDER_PRESETS);

		CString strCustomPresetsDir(szPath);
		strCustomPresetsDir += L"\\";

		// Get Operators root path
		CString strOperatorsRootDir;
		strOperatorsRootDir.Format(L"%s\\%s\\", strDataPath.GetString(), VSI_OPERATORS_FOLDER);

		// Get operator folder
		CString strOperatorDir;
		if (NULL != pszName)
		{
			strOperatorDir.Format(L"%s%s", strOperatorsRootDir.GetString(),	pszName);
		}
		else
		{
			// Double check service mode account
			CComPtr<IVsiOperator> pOperatorCurrent;
			hr = GetCurrentOperator(&pOperatorCurrent);
			VSI_CHECK_S_OK(hr, VSI_LOG_ERROR, NULL);

			VSI_OPERATOR_TYPE dwType;
			hr = pOperatorCurrent->GetType(&dwType);
			VSI_CHECK_HR(hr, VSI_LOG_ERROR, NULL);

			BOOL bServiceAccount = 0 != (VSI_OPERATOR_TYPE_SERVICE_MODE & dwType);			

			if (bServiceAccount)
			{
				WCHAR szTempPath[MAX_PATH];
				int iRet = GetTempPath(_countof(szTempPath), szTempPath);
				VSI_WIN32_VERIFY(0 != iRet, VSI_LOG_ERROR, NULL);

				PathAppend(szTempPath, L"ServiceAccount");

				strOperatorDir = szTempPath;
			}
		}

		// Get measurement package folder
		CString strOperatorMeasurementsDir;
		strOperatorMeasurementsDir.Format(
			L"%s\\" VSI_FOLDER_MEASUREMENTS,
			strOperatorDir.GetString());

		// Get preset folder
		CString strOperatorPresetsDir;
		strOperatorPresetsDir.Format(
			L"%s\\" VSI_FOLDER_PRESETS,
			strOperatorDir.GetString());

		if (0 != _waccess(strOperatorMeasurementsDir, 0))
		{
			// Copy DTD files
			CString strDtd(strCustomMeasurementsDir);
			strDtd += L"*.dtd";

			BOOL bRet = VsiCopyFiles(strDtd, strOperatorMeasurementsDir);
			VSI_VERIFY(bRet, VSI_LOG_ERROR, L"copy custom packages from system folder failed!");
		}

		if (0 != _waccess(strOperatorPresetsDir, 0))
		{
			// Copy default system presets application(axml file)
			WCHAR szPath[MAX_PATH];
			GetModuleFileName(NULL, szPath, _countof(szPath));
 
			LPWSTR pszSpt = wcsrchr(szPath, L'\\');
			*(pszSpt + 1) = 0;
 
			PathAppend(szPath, VSI_FOLDER_PRESETS);
			CString strSystemPresetsDir = szPath;

			VsiRemoveDirectory(strOperatorPresetsDir, FALSE);

			CString strSearch(strSystemPresetsDir);
			strSearch += L"\\*.*";
 
			WIN32_FIND_DATA ff;
			hFile = FindFirstFile(strSearch, &ff);
			BOOL bValid = (INVALID_HANDLE_VALUE != hFile);
			while (bValid)
			{
				if (ff.cFileName[0] != L'.')
				{			
					if (ff.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
					{
						// Search application under Subfolder
						CString strSubFolder = strSystemPresetsDir + L"\\" + ff.cFileName + L"\\";
						CString strFileSpec = strSubFolder + L"*.*xml";

						CString strTargetFolder = strOperatorPresetsDir + L"\\" + ff.cFileName + L"\\";

						BOOL bRet = VsiCopyFiles(strFileSpec, strTargetFolder);
						VSI_VERIFY(bRet, VSI_LOG_ERROR, L"copy application from system preset folder failed!");
					}
				}
				bValid = FindNextFile(hFile, &ff);
			} 
		}
	}
	VSI_CATCH(hr);

	if (INVALID_HANDLE_VALUE != hFile)
	{
		FindClose(hFile);
	}

	return hr;
}