/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  dllmain.h
**
**	Description:
**		Declaration of module class.
**
********************************************************************************/

class CVsiStudyModule : public CAtlDllModuleT< CVsiStudyModule >
{
public :
	DECLARE_LIBID(LIBID_VsiStudyModuleLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_VSISTUDYMODULE, "{281E78C8-7E3B-44E9-997A-F460FB05F52E}")
};

extern class CVsiStudyModule _AtlModule;
