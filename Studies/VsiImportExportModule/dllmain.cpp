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
#include <VsiImportExportModule.h>
#include "dllmain.h"
#include "xdlldata.h"

CVsiImportExportModule _AtlModule;

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	// Use resource DLL
	VSI_USE_RESOURCE_DLL(dwReason);

	if (DLL_PROCESS_ATTACH == dwReason)
	{
		// Setups the debug heap manager
		_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);

		// Disables the DLL_THREAD_ATTACH and DLL_THREAD_DETACH notifications
		DisableThreadLibraryCalls(hInstance);
	}

#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(hInstance, dwReason, lpReserved))
		return FALSE;
#endif

	hInstance;
	return _AtlModule.DllMain(dwReason, lpReserved); 
}
