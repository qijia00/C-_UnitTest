/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  dllmain.cpp
**
**	Description:
**		DLL main
**
********************************************************************************/

#include "stdafx.h"
#include "resource.h"
#include <VsiGlobalDef.h>
#include <VsiOperatorModule.h>
#include "dllmain.h"
#include "xdlldata.h"

CVsiOperatorModule _AtlModule;

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	// Use resource DLL
	VSI_USE_RESOURCE_DLL(dwReason);

	// Disables the DLL_THREAD_ATTACH and DLL_THREAD_DETACH notifications
	if (DLL_PROCESS_ATTACH == dwReason)
	{
		// Setups the debug heap manager
		_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);

		// Do not call DisableThreadLibraryCalls from a DLL that is linked to the static C run-time library
		DisableThreadLibraryCalls(hInstance);
	}

#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(hInstance, dwReason, lpReserved))
		return FALSE;
#endif

	hInstance;
	return _AtlModule.DllMain(dwReason, lpReserved); 
}
